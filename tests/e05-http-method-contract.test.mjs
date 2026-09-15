import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const methodContract = JSON.parse(await readFile(new URL('../governance/e05-http-method-contract.json', import.meta.url), 'utf8'));
const dataContract = JSON.parse(await readFile(new URL('../governance/data-contract.json', import.meta.url), 'utf8'));
const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const plan = await readFile(new URL('../PLANO_EVOLUCAO_BACKEND_GEST_CSI.md', import.meta.url), 'utf8');
const catalog = dataContract.interfaces.apps_script_actions;
const classes = ['read_verified', 'read_candidate', 'conditional_effect_get', 'post_required'];

test('classifies each of the 56 legacy actions exactly once', () => {
  const assigned = classes.flatMap((name) => methodContract[name]);
  assert.equal(catalog.length, 56);
  assert.equal(new Set(assigned).size, assigned.length);
  assert.deepEqual(assigned.slice().sort(), catalog.slice().sort());
});

test('does not silently accept E05 while mutable GET routes remain', () => {
  const getRouter = source.slice(source.indexOf('function doGet(e)'), source.indexOf('function legacyActionPolicy_'));
  const outstanding = methodContract.post_required.filter((action) =>
    getRouter.includes(`action === '${action}'`));
  const e05 = plan.slice(plan.indexOf('#### E05'), plan.indexOf('#### E06'));
  if (/\*\*Status:\*\* \[x\] OK/.test(e05)) {
    assert.deepEqual(outstanding, []);
    assert.deepEqual(methodContract.conditional_effect_get, []);
    assert.deepEqual(methodContract.read_candidate, []);
  } else {
    assert.ok(outstanding.length + methodContract.conditional_effect_get.length + methodContract.read_candidate.length > 0);
    assert.ok(!outstanding.includes('login'));
    assert.ok(!outstanding.includes('syncProject'));
    assert.ok(outstanding.includes('logPanelUpdate'));
  }
});
