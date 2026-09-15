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

test('status GET helpers have no setters or trigger scheduling', () => {
  for (const action of ['health', 'getProjectSyncBackgroundStatus', 'getProjectClosingSyncStatus',
    'getClickUpUserActivityStatus', 'getCmaxDailyHistoryStatus', 'me']) {
    assert.ok(contract.read_verified.includes(action), action);
  }
  for (const name of ['getProjectSyncBackgroundStatus_', 'getProjectClosingSyncBackgroundStatus_',
    'getClickUpUserActivityBackgroundStatus_', 'getCmaxDailyHistoryStatus_']) {
    const code = body(name);
    assert.doesNotMatch(code, /\.setProperty\(|\.deleteProperty\(|\.setProperties\(|schedule[A-Z]|ScriptApp\.newTrigger/);
  }
  const activity = body('getClickUpUserActivityBackgroundStatus_');
  assert.doesNotMatch(activity, /preservePreQueueProjectSyncRequest_|clearClickUpUserActivityBackgroundTriggers_/);
});

test('session lookup does not extend TTL during GET; explicit POST does', () => {
  const cache = { writes: 0, get() { return JSON.stringify({ role: 'consultant', enabled: true }); }, put() { this.writes += 1; } };
  const context = {
    CacheService: { getScriptCache: () => cache },
    sanitizeText_: (value) => String(value || '').trim(),
    makeAuthToken_: () => 'new-session',
    publicUser_: (user) => user,
    JSON
  };
  vm.runInNewContext(['sessionCache_', 'storeSession_', 'requireUser_', 'refreshUserSession_']
    .map(body).join('\n'), context);
  assert.equal(context.requireUser_({ auth_token: 'valid' }).role, 'consultant');
  assert.equal(cache.writes, 0);
  assert.equal(context.refreshUserSession_({ auth_token: 'valid' }).refreshed, true);
  assert.equal(cache.writes, 1);
});

test('sheet-backed status reads no longer create missing sheets', () => {
  for (const name of ['getPanelUpdateHistory_', 'listUsers_', 'getSharedProjectFollowups_',
    'getProjectKanbanStates_', 'getProjectClosingCandidateRows_']) {
    const code = body(name);
    assert.match(code, /getSheetByName\(/);
    assert.doesNotMatch(code, /insertSheet\(|ensureHeaders_\(|getPanelUpdateHistorySheet_|getUsersSheet_|getProjectFollowupSheet_|getProjectKanbanStateSheet_|getProjectClosingCandidateSheet_/);
  }
});
