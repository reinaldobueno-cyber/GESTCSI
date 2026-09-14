import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { sha256 } from '../src/baseline-metrics.mjs';

const baseline = JSON.parse(await readFile(new URL('../baselines/e02-2026-09-14.json', import.meta.url), 'utf8'));

test('locks the baseline to the verified production version', () => {
  assert.equal(baseline.schema_version, 1);
  assert.equal(baseline.production_base.git_sha, '11e7c2fdc9af2b1f36f1217ae90a061773f19490');
  assert.equal(baseline.production_base.app_version, '2026-09-11-clickup-activity-v4');
});

test('keeps source evidence internally consistent', () => {
  for (const source of Object.values(baseline.sources)) {
    assert.equal(source.status, 200);
    assert.match(source.sha256, /^[a-f0-9]{64}$/);
  }
  assert.equal(
    baseline.aggregate_sha256,
    sha256(JSON.stringify({ sources: baseline.sources, competences: baseline.competences }))
  );
});

test('reconciles both reference competences without missing rates', () => {
  assert.deepEqual(Object.keys(baseline.competences), ['2026-07', '2026-08']);
  for (const competence of Object.values(baseline.competences)) {
    assert.equal(competence.reconciliation.component_sum_matches_total, true);
    assert.equal(competence.reconciliation.portfolio_declared_total_matches_rows, true);
    assert.equal(competence.reconciliation.cmax_has_no_missing_rates, true);
  }
});

test('contains only aggregate baseline fields', () => {
  const prohibitedKeys = new Set([
    'cliente', 'client', 'consultor', 'consultant', 'consultant_name',
    'project_name', 'created_by', 'decided_by', 'notes', 'auth_token'
  ]);
  function inspect(value) {
    if (!value || typeof value !== 'object') return;
    for (const [key, child] of Object.entries(value)) {
      assert.equal(prohibitedKeys.has(key), false, `campo identificável encontrado: ${key}`);
      inspect(child);
    }
  }
  inspect(baseline);
});
