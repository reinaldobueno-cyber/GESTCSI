import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFile } from 'node:fs/promises';

const html = await readFile(new URL('../index.html', import.meta.url), 'utf8');

function functionBody(name, next) {
  const start = html.indexOf(`function ${name}(`);
  const end = html.indexOf(next, start + 10);
  assert.ok(start >= 0 && end > start, name);
  return html.slice(start, end);
}

const source = functionBody('gestaoMetricCell', 'function gestaoTotalConcluidos(')
  + functionBody('gestaoRenderClickupProductivity', 'var GESTAO_CLICKUP_INVENTORY');

function render(error) {
  const empty = { projetos: [], label: 'PROJETOS <= 2025', key: 'legacy', tasksPendentes: 0, marcosPendentes: 0, atrasados: 0, semUpdate: 0 };
  const current = { ...empty, projetos: Array(215).fill({}), label: 'PROJETOS 2026', key: 'current' };
  const total = { ...current, label: 'TOTAL CLICKUP', key: 'total' };
  const stats = { projetos: 0, tasksConcluidas: 0, tasksPendentes: 0, marcosConcluidos: 0, marcosPendentes: 0, atrasados: 0, semUpdate: 0 };
  const context = {
    gestaoBuildClickupProductivity: () => ({
      base: current.projetos,
      buckets: { legacy: empty, current, total },
      consultores: [{ key: 'c1', nome: 'Consultor', legacy: stats, current: { ...stats, projetos: 12 }, total: { ...stats, projetos: 12 }, exemplos: [] }]
    }),
    gestaoTotalConcluidos: () => 0,
    gestaoAuditButton: () => '<button>auditoria</button>',
    gestaoRenderClickupUserActivityShell: () => '',
    gestaoOpenConsultorAudit: () => false,
    jsQuote: (value) => value,
    escapeHtml: (value) => String(value),
    fn: (value) => String(value)
  };
  vm.runInNewContext(source, context);
  return context.gestaoRenderClickupProductivity([], error ? { erro: 'Timeout ao acessar Apps Script' } : {});
}

test('timeout marks historical count unknown and total partial, leaving 2026 visible', () => {
  const rendered = render(true);
  assert.match(rendered, /tower-kpi-value">—<\/div><div class="tower-kpi-label">PROJETOS <= 2025/);
  assert.match(rendered, /tower-kpi-value">215<\/div><div class="tower-kpi-label">PROJETOS 2026/);
  assert.match(rendered, /tower-kpi-value">≥215<\/div><div class="tower-kpi-label">TOTAL CLICKUP \(PARCIAL\)/);
  assert.match(rendered, /<b>Projetos<\/b> ≥215 \(parcial\)/);
  assert.match(rendered, /Histórico indisponível; dados deste consultor são parciais/);
  assert.doesNotMatch(rendered, /nenhum projeto histórico foi encontrado/);
});

test('a valid zero-result response still displays zero rather than unavailable', () => {
  const rendered = render(false);
  assert.match(rendered, /tower-kpi-value">0<\/div><div class="tower-kpi-label">PROJETOS <= 2025/);
  assert.match(rendered, /nenhum projeto histórico foi encontrado/);
  assert.doesNotMatch(rendered, /TOTAL CLICKUP \(PARCIAL\)/);
});

function inventoryLoader(response, cachedProjects) {
  const saved = cachedProjects ? JSON.stringify({ saved_at: Date.now(), projetos: cachedProjects }) : 'null';
  const context = {
    AUTH_STATE: { user: { username: 'consultor' } },
    GESTAO_CLICKUP_INVENTORY: null,
    GESTAO_CLICKUP_INVENTORY_LOADING: null,
    localStorage: { getItem: () => saved, setItem: () => {} },
    chamarAppsScriptJsonp: () => Promise.resolve(response),
    authTokenParam: () => 'session',
    gestaoNormalizeInventoryProject: (item) => item,
    setTimeout: (callback) => callback()
  };
  vm.runInNewContext(functionBody('gestaoLoadClickUpInventory', 'function gestaoBucketLabel('), context);
  return context.gestaoLoadClickUpInventory();
}

test('incomplete inventory response does not poison the view or an empty cache', async () => {
  await assert.rejects(inventoryLoader({ ok: true, projetos: [], has_more: false }, []), /abaixo do piso 201/);
});

test('incomplete inventory response retains a previously valid historical cache', async () => {
  const projects = Array.from({ length: 337 }, (_, index) => ({ cliente: `Cliente ${index + 1}` }));
  const result = await inventoryLoader({ ok: true, projetos: [], has_more: false }, projects);
  assert.equal(result.length, 337);
  assert.match(result._inventoryWarning, /última leitura válida/);
});

test('historical audit links report unavailable rather than an empty project slice', () => {
  let modal = '';
  let filteredLegacy = false;
  const context = {
    GESTAO_CLICKUP_HISTORY_UNAVAILABLE: true,
    dashOpenAuditModal: (_title, content) => { modal = content; },
    gestaoFilterProjects: (_mode, bucket) => { if (bucket === 'legacy') filteredLegacy = true; return []; },
    gestaoConsultorLabel: (_best, value) => value,
    gestaoProjectRow: () => '',
    gestaoIsProjectLate: () => false,
    gestaoBucketLabel: () => '',
    escapeHtml: (value) => String(value)
  };
  vm.runInNewContext(functionBody('gestaoOpenAudit', 'function gestaoOpenConsultorAudit(')
    + functionBody('gestaoOpenConsultorAudit', 'function initControlTower('), context);
  assert.equal(context.gestaoOpenAudit('todos', 'legacy'), false);
  assert.match(modal, /inventário histórico não foi carregado/);
  assert.equal(context.gestaoOpenConsultorAudit('Consultor'), false);
  assert.match(modal, /Histórico indisponível/);
  assert.doesNotMatch(modal, /0 histórico|Histórico <= 2025 • 0 projeto/);
  assert.equal(filteredLegacy, false);
});
