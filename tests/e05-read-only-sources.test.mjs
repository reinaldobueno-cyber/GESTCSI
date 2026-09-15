import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const contract = JSON.parse(await readFile(new URL('../governance/e05-http-method-contract.json', import.meta.url), 'utf8'));

function body(name) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf('\nfunction ', start + 10);
  assert.ok(start >= 0 && end > start, name);
  return source.slice(start, end);
}

test('inventory, milestone, decisions and financial GETs do not call sheet creators', () => {
  const names = ['getClickUpInventory_', 'getClickUpMilestoneClosing_', 'getProjectClosingDecisions_',
    'getConsultantCompensationData_', 'getBonusSalesIndications_'];
  for (const name of names) {
    const code = body(name);
    assert.match(code, /getSheetByName\(/, name);
    assert.doesNotMatch(code, /insertSheet\(|ensureHeaders_\(|getClickUpInventorySheet_|getClickUpMilestoneClosingSheet_|getProjectClosingDecisionSheet_|getConsultantCompensationSheet_|getUsersSheet_|getBonusSalesIndicationsSheet_/);
  }
  for (const action of ['getClickUpInventory', 'getClickUpMilestoneClosing',
    'getProjectClosingDecisions', 'getConsultantCompensation', 'getBonusSalesIndications']) {
    assert.ok(contract.read_verified.includes(action), action);
  }
});

test('missing inventory, milestone, decisions and bonus sheets are unavailable, not successful zeroes', () => {
  const names = ['getClickUpInventory_', 'getClickUpMilestoneClosing_', 'getProjectClosingDecisions_', 'getBonusSalesIndications_'];
  let inserts = 0;
  const context = {
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => null, insertSheet: () => { inserts += 1; } }) },
    getScriptProperty_: (_name, fallback) => fallback || 'fake-sheet-id',
    requireUser_: () => ({ role: 'user' }),
    requireAdmin_: () => ({ role: 'admin' }),
    sanitizeMonth_: () => '',
    BONUS_SALES_INDICATIONS_SHEET: 'BONUS_INDICACOES'
  };
  vm.runInNewContext(names.map(body).join('\n'), context);
  for (const name of names) {
    const result = context[name]({ auth_token: 'test' });
    assert.equal(result.ok, false, name);
    assert.equal(result.source_unavailable, true, name);
  }
  assert.equal(inserts, 0);
});

test('compensation read does not create remuneration or users sheets', () => {
  let inserts = 0;
  const context = {
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => null, insertSheet: () => { inserts += 1; } }) },
    getScriptProperty_: (_name, fallback) => fallback || 'fake-sheet-id',
    CONSULTANT_COMPENSATION_SHEET: 'CONSULTORES_REMUNERACAO',
    CONSULTANT_SENIORITY_RATES: { master: 220 }
  };
  vm.runInNewContext(body('getConsultantCompensationData_'), context);
  assert.equal(context.getConsultantCompensationData_().ok, true);
  assert.equal(inserts, 0);
});

test('CMAX GET does not create view sheets or write cache entries', () => {
  const cmax = body('getCmaxDailyEvents_');
  assert.doesNotMatch(cmax, /getOrCreateSheet_|insertSheet\(|ensureHeaders_|writeCompressedScriptCache_|scheduleCmaxDailyViewBuild_/);
  assert.match(cmax, /getSheetByName\(CMAX_DAILY_VIEW_SHEET\)/);
});

test('remaining local data GETs use read-only source helpers', () => {
  for (const action of ['getClickUpUserActivity', 'getPanelUpdateHistory', 'listUsers',
    'getProjectFollowups', 'getMonthlyProjects', 'getProjectClosingCandidates', 'getCmaxDailyEvents']) {
    assert.ok(contract.read_verified.includes(action), action);
  }
  for (const name of ['getClickUpUserActivity_', 'getPanelUpdateHistory_', 'listUsers_',
    'getSharedProjectFollowups_', 'getProjectKanbanStates_', 'getProjectClosingCandidateRows_']) {
    const code = body(name);
    assert.doesNotMatch(code, /insertSheet\(|ensureHeaders_\(|\.setValues\(|\.appendRow\(|\.setProperty\(/, name);
  }
  const monthly = body('getMonthlyProjectsPayload_');
  const candidates = body('getProjectClosingCandidates_');
  const cmax = body('getCmaxDailyEvents_');
  assert.doesNotMatch(monthly, /scheduleMonthlyPortfolioSnapshotRefresh_|\.setProperty\(|insertSheet\(/);
  assert.doesNotMatch(candidates, /refreshProjectClosingCandidates_|writeProjectClosingCandidateRows_|\.setProperty\(/);
  assert.doesNotMatch(cmax, /scheduleCmaxDailyViewBuild_|writeCompressedScriptCache_|getOrCreateSheet_|\.setProperty\(/);
});

test('empty followup responses preserve the kanban-state field', () => {
  const code = body('getSharedProjectFollowups_');
  assert.equal((code.match(/followups: \[\], total: 0, kanban_states: getProjectKanbanStates_\(\)/g) || []).length, 2);
});
