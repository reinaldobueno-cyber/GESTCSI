import test from 'node:test';
import assert from 'node:assert/strict';
import '../src/portfolio-integrity.js';

const integrity = globalThis.PortfolioIntegrity;

function projects(count, month = 'JAN') {
  return Array.from({ length: count }, (_, index) => ({
    project_key: `${month}|CLIENTE ${index + 1}`,
    mes: month,
    cliente: `Cliente ${index + 1}`
  }));
}

test('accepts the currently verified 201-project portfolio', () => {
  const list = projects(201);
  const result = integrity.validateSnapshot({
    projetos: list,
    total: 201,
    projetos_por_mes: { JAN: 201 }
  }, { minimumTotal: 201, requireDeclaredTotal: true });

  assert.equal(result.ok, true);
  assert.equal(result.manifest.total, 201);
});

test('automatically accepts a newly added project without changing the configured floor', () => {
  const candidate = projects(202);
  const result = integrity.shouldReplaceSnapshot({
    projetos: candidate,
    total: 202,
    projetos_por_mes: { JAN: 202 }
  }, projects(201), { minimumTotal: 201, requireDeclaredTotal: true });

  assert.equal(result.replace, true);
  assert.equal(result.manifest.total, 202);
});

test('rejects the stale 199-project snapshot', () => {
  const result = integrity.validateSnapshot({
    projetos: projects(199),
    total: 199,
    projetos_por_mes: { JAN: 199 }
  }, { minimumTotal: 201, requireDeclaredTotal: true });

  assert.equal(result.ok, false);
  assert.ok(result.errors.includes('below_minimum_total'));
});

test('accepts a portfolio whose declared totals match the payload', () => {
  const list = projects(197);
  const result = integrity.validateSnapshot({
    projetos: list,
    total: 197,
    projetos_por_mes: { JAN: 197, FEV: 0 }
  }, { minimumTotal: 197, requireDeclaredTotal: true });

  assert.equal(result.ok, true);
  assert.equal(result.manifest.total, 197);
});

test('rejects a partial list even when the server declares the complete total', () => {
  const result = integrity.validateSnapshot({
    projetos: projects(194),
    total: 197,
    projetos_por_mes: { JAN: 197 }
  }, { minimumTotal: 197, requireDeclaredTotal: true });

  assert.equal(result.ok, false);
  assert.ok(result.errors.includes('below_minimum_total'));
  assert.ok(result.errors.includes('declared_total_mismatch'));
  assert.ok(result.errors.includes('month_total_mismatch:JAN'));
});

test('rejects duplicate stable project keys', () => {
  const list = projects(2);
  list[1].project_key = list[0].project_key;
  const result = integrity.validateSnapshot({ projetos: list, total: 2 }, {
    minimumTotal: 1,
    requireDeclaredTotal: true
  });

  assert.equal(result.ok, false);
  assert.ok(result.errors.includes('duplicate_project_keys'));
});

test('does not replace a larger valid snapshot with a smaller candidate', () => {
  const result = integrity.shouldReplaceSnapshot(
    { projetos: projects(196), total: 196 },
    projects(197),
    { minimumTotal: 1, requireDeclaredTotal: true }
  );

  assert.equal(result.replace, false);
  assert.equal(result.reason, 'candidate_smaller_than_current');
});

test('builds stable identities only when a persistent row or key exists', () => {
  assert.equal(integrity.projectIdentity({ mes: 'JAN', cliente: 'Fazenda Águas', _sheet_row: 12 }), 'JAN|FAZENDA AGUAS|12');
  assert.equal(integrity.projectIdentity({ mes: 'JAN', cliente: 'Fazenda Águas' }), '');
});
