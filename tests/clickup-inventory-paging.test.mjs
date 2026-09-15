import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const start = source.indexOf('function getClickUpInventory_(params)');
const end = source.indexOf('\nfunction getClickUpMilestoneClosing_', start);
assert.ok(start >= 0 && end > start, 'inventory function must be present');

const headers = [
  'mes', 'cliente', 'consultor', 'status', 'project_key', 'project_url',
  'view_id', 'list_id', 'folder_id', 'space_id', 'tasks_concluidas',
  'tasks_pendentes', 'marcos_concluidos', 'marcos_pendentes', 'fases_total',
  'progresso', 'data_ultima_atualizacao', 'dias_sem_atualizacao',
  'clickup_json', 'ultima_sync_clickup', 'sync_status_clickup', 'sync_error_clickup'
];
const rows = Array.from({ length: 337 }, (_, index) => headers.map((header) => {
  if (header === 'mes') return 'HIST';
  if (header === 'cliente') return `Cliente ${index + 1}`;
  if (header === 'project_key') return `HIST|${index + 1}`;
  if (header === 'tasks_concluidas') return index;
  if (header === 'clickup_json') return 'x'.repeat(45000);
  return '';
}));
const allValues = [headers, ...rows];
const rangeCalls = [];
const sheet = {
  getLastRow: () => allValues.length,
  getLastColumn: () => headers.length,
  getRange(row, column, rowCount, columnCount) {
    rangeCalls.push({ row, column, rowCount, columnCount });
    return {
      getValues: () => allValues.slice(row - 1, row - 1 + rowCount)
        .map((item) => item.slice(column - 1, column - 1 + columnCount))
    };
  }
};
const context = {
  requireUser_: () => ({ role: 'admin' }),
  getClickUpInventorySheet_: () => sheet,
  toInt_: (value, fallback) => {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : fallback;
  },
  rowToObject_: (header, row) => Object.fromEntries(header.map((name, index) => [name, row[index]])),
  sanitizeText_: (value) => String(value ?? '').trim(),
  canUserAccessProjectItem_: () => true
};
vm.runInNewContext(source.slice(start, end), context);

test('reads the 337 historical projects in bounded pages without reading clickup_json', () => {
  const collected = [];
  let offset = 0;
  let response;
  do {
    response = context.getClickUpInventory_({ lean: '1', paged: '1', offset, limit: 100 });
    collected.push(...response.projetos);
    offset = response.next_offset;
  } while (response.has_more);

  assert.equal(collected.length, 337);
  assert.equal(response.scanned_total, 337);
  assert.equal(response.next_offset, 337);
  assert.equal(response.done, true);
  assert.deepEqual(rangeCalls.filter((call) => call.row > 1 && call.column === 1).map((call) => call.rowCount), [100, 100, 100, 37]);

  const jsonColumn = headers.indexOf('clickup_json') + 1;
  const dataReads = rangeCalls.filter((call) => call.row > 1);
  assert.ok(dataReads.every((call) => jsonColumn < call.column || jsonColumn >= call.column + call.columnCount));
  assert.ok(collected.every((item) => JSON.parse(item.clickup_json).compacto_gestao === true));
});
