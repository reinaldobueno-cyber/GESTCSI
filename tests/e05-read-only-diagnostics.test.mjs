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

test('both remaining diagnostics and their transitive helpers contain no external writes', () => {
  for (const action of ['diagnoseClickUpMilestoneTask', 'diagnoseProjectClosingCandidateCounts']) {
    assert.ok(contract.read_verified.includes(action), action);
    assert.ok(!contract.read_candidate.includes(action), action);
  }
  const helpers = [
    'diagnoseClickUpMilestoneTask_', 'diagnoseProjectClosingCandidateCounts_',
    'loadClickUpMilestoneClosingMappings_', 'findProjectMappingForTask_',
    'fallbackProjectMappingForTask_', 'buildNormalizedMilestoneCoverageProject_',
    'loadProjectClosingCandidateMappings_', 'projectClosingCandidateFromTask_',
    'projectClosingSavedBreakOffCandidates_', 'refreshProjectClosingCandidatesByTaskId_',
    'fetchClickUpProjectClosingApprovalTaskScan_', 'fetchClickUpTasksByIds_',
    'getMonthlyProjectsFromSheet_', 'loadProjectMappings_'
  ];
  for (const name of helpers) {
    const code = body(name);
    assert.doesNotMatch(code, /\.setProperty\(|\.setProperties\(|\.put\(|\.appendRow\(|\.setValues\(|insertSheet\(|ensureHeaders_\(|ScriptApp\.newTrigger|clickupRequest_\('(post|put|delete)'/i, name);
  }
  assert.match(body('diagnoseClickUpMilestoneTask_'), /clickupRequest_\('get'/);
  assert.match(body('fetchClickUpProjectClosingApprovalTaskScan_'), /clickupRequest_\('get'/);
  assert.match(body('fetchClickUpTasksByIds_'), /method: 'get'/);
});

test('isolated diagnostics perform reads and return counts without scheduling or writing', () => {
  let clickupReads = 0;
  let writes = 0;
  const context = {
    normalizeClickUpId_: (value) => String(value || '').trim(),
    sanitizeText_: (value) => String(value || '').trim(),
    clickupRequest_: (method) => { assert.equal(method, 'get'); clickupReads += 1; return { id: 'task-1', name: 'Marco', status: { status: 'Closed' } }; },
    clickUpTaskStatusText_: (task) => task.status.status,
    loadClickUpMilestoneClosingMappings_: () => [],
    findProjectMappingForTask_: () => ({ project_key: 'P1', cliente: 'Projeto 1' }),
    fallbackProjectMappingForTask_: () => ({ project_key: 'P1', cliente: 'Projeto 1' }),
    buildNormalizedMilestoneCoverageProject_: () => ({ marcos: [{ id: 'task-1' }] }),
    clickUpMilestoneSituation_: () => 'aguardando',
    isMilestoneTask_: () => true,
    isProjectDeliveryTask_: () => false,
    isClosingTrackedTask_: () => true,
    clickUpTaskCustomItemName_: () => '',
    fromMillisIso_: () => '',
    loadProjectClosingCandidateMappings_: () => [],
    fetchClickUpProjectClosingApprovalTaskScan_: () => ({ tasks: [{ id: 'task-1' }], raw_tasks: 1, status_tasks: 1, delivery_tasks: 1, closing_tasks: 1 }),
    projectClosingCandidateFromTask_: () => ({ item_id: 'task-1', item_status: 'APROVAR' }),
    projectClosingSavedBreakOffCandidates_: () => [],
    refreshProjectClosingCandidatesByTaskId_: (items) => items,
    isProjectClosingApprovalStatus_: (value) => value === 'APROVAR',
    simplifyErrorMessage_: (error) => String(error.message),
    CLICKUP_PROJECT_CLOSING_RULE_VERSION: 'test',
    PropertiesService: { getScriptProperties: () => ({ setProperty: () => { writes += 1; } }) },
    ScriptApp: { newTrigger: () => { writes += 1; } }
  };
  vm.runInNewContext(['diagnoseClickUpMilestoneTask_', 'diagnoseProjectClosingCandidateCounts_'].map(body).join('\n'), context);
  assert.equal(context.diagnoseClickUpMilestoneTask_({ task_id: 'task-1' }).ok, true);
  assert.equal(context.diagnoseProjectClosingCandidateCounts_().eligible_after_status_refresh, 1);
  assert.equal(clickupReads, 1);
  assert.equal(writes, 0);
});
