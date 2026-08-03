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
