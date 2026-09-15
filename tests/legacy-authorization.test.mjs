import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const contract = JSON.parse(await readFile(new URL('../governance/data-contract.json', import.meta.url), 'utf8'));
const policyStart = source.indexOf('function legacyActionPolicy_(action)');
const policyEnd = source.indexOf('\n/**\n * Lightweight deployment probe.', policyStart);
assert.ok(policyStart >= 0 && policyEnd > policyStart);
const context = {
  requireAdmin_: ({ auth_token: token }) => token === 'admin' ? { role: 'admin' } : (() => { throw new Error('admin required'); })(),
  requireUser_: ({ auth_token: token }) => ['admin', 'manager', 'consultant'].includes(token)
    ? { role: token } : (() => { throw new Error('login required'); })(),
  loginUser_: () => ({ ok: true, token: 'new-session' }),
  syncProjectByKey: (key) => ({ project_key: key }),
  processDirtyQueue: ({ limit }) => ({ ok: true, limit }),
  validarClickUpConfig: () => ({ ok: true, sheet: 'diagnostic' }),
  toInt_: (value, fallback) => Number.parseInt(value, 10) || fallback
};
vm.runInNewContext(source.slice(policyStart, policyEnd), context);
const actions = contract.interfaces.apps_script_actions;

test('assigns a role to every cataloged legacy route and denies unknown actions', () => {
  assert.equal(actions.length, 56);
  assert.deepEqual(actions.filter((action) => context.legacyActionPolicy_(action) === 'unknown'), []);
  assert.equal(context.legacyActionPolicy_('notAnAction'), 'unknown');
  assert.equal(context.legacyActionPolicy_('', 'JAN'), 'user');
  assert.equal(context.legacyActionPolicy_(''), 'public');
});

test('enforces anonymous, consultant, manager and admin access at the router', () => {
  for (const action of actions) {
    const policy = context.legacyActionPolicy_(action);
    for (const token of ['', 'consultant', 'manager', 'admin']) {
      const allowed = policy === 'public' ||
        (policy === 'user' && token !== '') ||
        (policy === 'admin' && token === 'admin');
      if (allowed) assert.doesNotThrow(() => context.authorizeLegacyAction_(action, { auth_token: token }), `${action}/${token}`);
      else assert.throws(() => context.authorizeLegacyAction_(action, { auth_token: token }), `${action}/${token}`);
    }
  }
});

test('does not let query aliases bypass the history write guard', () => {
  assert.equal(context.legacyActionPolicy_('logPanelUpdate'), 'user');
  assert.equal(context.legacyActionPolicy_('syncProject'), 'admin');
  assert.equal(context.legacyActionPolicy_('processDirty'), 'admin');
  assert.throws(() => context.authorizeLegacyAction_('', { log_update: '1' }));
  assert.throws(() => context.authorizeLegacyAction_('', { history: '1' }));
});

test('keeps project-closing status polling read-only', () => {
  const routeStart = source.indexOf("if (action === 'getProjectClosingSyncStatus')");
  const routeEnd = source.indexOf("if (action === 'stopProjectClosingSync')", routeStart);
  const route = source.slice(routeStart, routeEnd);
  assert.match(route, /getProjectClosingSyncBackgroundStatus_\(\)/);
  assert.doesNotMatch(route, /advanceProjectClosingSyncBackgroundFromStatus_\(\)/);
});

test('passes the administrative session into the protected milestone readback', () => {
  const start = source.indexOf('function syncClickUpMilestoneTask_(params)');
  const end = source.indexOf('\nfunction confirmClickUpMilestoneStatuses_', start);
  const route = source.slice(start, end);
  assert.match(route, /diagnosis\.after = getClickUpMilestoneClosing_\(\{\s*auth_token: params\.auth_token/);
});

test('routes login and the first three administrative commands through POST only', () => {
  const getRouter = source.slice(source.indexOf('function doGet(e)'), source.indexOf('function legacyActionPolicy_'));
  for (const action of ['login', 'syncProject', 'processDirty', 'validateConfig']) {
    assert.equal(context.legacyPostOnlyAction_(action), true);
    assert.doesNotMatch(getRouter, new RegExp(`if \\(action === '${action}'\\)`));
  }
  assert.match(getRouter, /if \(legacyPostOnlyAction_\(action\)\)/);
  assert.equal(context.dispatchLegacyPostCommand_('login', {}).token, 'new-session');
  assert.equal(context.dispatchLegacyPostCommand_('syncProject', { auth_token: 'admin', project_key: 'A' }).project_key, 'A');
  assert.equal(context.dispatchLegacyPostCommand_('processDirty', { auth_token: 'admin', limit: '5' }).limit, 5);
  assert.equal(context.dispatchLegacyPostCommand_('validateConfig', { auth_token: 'admin' }).sheet, 'diagnostic');
  assert.throws(() => context.dispatchLegacyPostCommand_('syncProject', { auth_token: 'consultant' }));
  assert.throws(() => context.dispatchLegacyPostCommand_('unknown', { auth_token: 'admin' }));
});
