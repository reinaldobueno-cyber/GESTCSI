import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const contract = JSON.parse(await readFile(new URL('../governance/e05-http-method-contract.json', import.meta.url), 'utf8'));

function functionBody(name, nextName) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf(`function ${nextName}(`, start);
  assert.ok(start >= 0 && end > start, `${name}/${nextName}`);
  return source.slice(start, end);
}

test('moves known refresh-on-read effects to explicit POST dispatch', () => {
  assert.deepEqual(contract.conditional_effect_get, []);
  for (const action of ['getMonthlyProjects', 'getProjectClosingCandidates', 'getCmaxDailyEvents']) {
    assert.ok(contract.read_candidate.includes(action), `${action} still needs complete read audit`);
  }
  const router = source.slice(source.indexOf('function doGet(e)'), source.indexOf('function legacyActionPolicy_'));
  assert.match(router, /legacyRefreshAction_\(action\) && String\(params\.refresh/);
  const post = functionBody('dispatchLegacyPostCommand_', 'getHealthPayload_');
  assert.match(post, /scheduleMonthlyPortfolioSnapshotRefresh_\(\)/);
  assert.match(post, /refreshProjectClosingCandidates_\(params\)/);
  assert.match(post, /scheduleCmaxDailyViewBuild_\(\)/);
});

test('monthly and CMAX GET payload builders do not schedule refresh', () => {
  const monthly = functionBody('getMonthlyProjectsPayload_', 'monthlyPortfolioSnapshotProject_');
  const cmax = functionBody('getCmaxDailyEvents_', 'cmaxDailyMonthSnapshotSheetName_');
  assert.doesNotMatch(monthly, /scheduleMonthlyPortfolioSnapshotRefresh_\(/);
  assert.doesNotMatch(cmax, /scheduleCmaxDailyViewBuild_\(/);
  assert.doesNotMatch(cmax, /writeCompressedScriptCache_\(/);
});

test('candidate GET reads saved rows without rebuilding them', () => {
  const read = functionBody('getProjectClosingCandidates_', 'refreshProjectClosingCandidates_');
  const rows = functionBody('getProjectClosingCandidateRows_', 'writeProjectClosingCandidateRows_');
  assert.doesNotMatch(read, /writeProjectClosingCandidateRows_|refreshProjectClosingCandidates_|fetchClickUpProjectClosingApprovalTasks_/);
  assert.match(rows, /getSheetByName\('CLICKUP_PROJECT_CLOSING_CANDIDATES'\)/);
  assert.doesNotMatch(rows, /getProjectClosingCandidateSheet_|ensureHeaders_/);
});
