import test from 'node:test';
import assert from 'node:assert/strict';
import {
  cmaxMetrics,
  indicationMetrics,
  milestoneMetrics,
  parseCsv,
  portfolioMetrics,
  projectClosingMetrics,
  reconciliationFor,
  sha256
} from '../src/baseline-metrics.mjs';

test('parses quoted CSV without leaking structure across rows', () => {
  const rows = parseCsv('id,name,note\n1,"A, B","line 1\nline 2"\n2,C,ok\n');
  assert.deepEqual(rows, [
    { id: '1', name: 'A, B', note: 'line 1\nline 2' },
    { id: '2', name: 'C', note: 'ok' }
  ]);
});

test('creates deterministic SHA-256 evidence', () => {
  assert.equal(sha256('GEST CSI'), '1abd6e90ba7db01187c6b79574d43581fb86ce3ab31619d398661ed359640fe3');
});

test('reconciles portfolio rows against the declared snapshot total', () => {
  const rows = [
    { mes: 'JUL', snapshot_total: '2', source_sheet_row: '1' },
    { mes: 'AGO', snapshot_total: '2', source_sheet_row: '2' }
  ];
  assert.deepEqual(portfolioMetrics(rows, 'JUL'), {
    total_rows: 2,
    declared_snapshot_totals: [2],
    competence_rows: 1,
    distinct_source_row_numbers: 2
  });
});

test('calculates CMAX caps, proportional duration and special-day multiplier', () => {
  const events = [
    { event_key: 1, data: '2026-08-01', consultor: 'Pessoa A', resultado: 'Positivo', modalidade: 'Treinamento On Line', hora_inicio: '08:00', hora_fim: '10:00' },
    { event_key: 2, data: '2026-08-01', consultor: 'Pessoa A', resultado: 'Positivo', modalidade: 'Treinamento On Line', hora_inicio: '10:00', hora_fim: '14:00' },
    { event_key: 3, data: '2026-08-02', consultor: 'Pessoa B', resultado: 'Negativo', modalidade: 'Treinamento In Loco', hora_inicio: '08:00', hora_fim: '17:00' }
  ];
  const compensation = [{ consultant_name: 'Pessoa A', seniority: 'pleno', daily_value: '85' }];
  const metrics = cmaxMetrics(events, compensation, '2026-08');
  assert.equal(metrics.payable_equivalent, 0.8333);
  assert.equal(metrics.special_equivalent, 0.8333);
  assert.equal(metrics.daily_value, 106.2457);
  assert.equal(metrics.consultants_without_rate, 0);
});

test('separates milestone, project-closing and indication values', () => {
  const milestones = [
    { task_id: 'm1', situacao: 'aprovado', mes_fechamento: '2026-07', validation_at: '2026-08-01', valor_bonus: '30' },
    { task_id: 'm2', situacao: 'reprovado', mes_fechamento: '2026-07', mes_validacao: '2026-07', valor_bonus: '0' },
    { task_id: 'p1', situacao: 'aprovado', mes_fechamento: '2026-07', item_tipo: 'Fechamento de projeto', valor_bonus: '80' }
  ];
  assert.deepEqual(milestoneMetrics(milestones, '2026-07'), {
    closed_rows: 2,
    closed_status: { aprovado: 1, reprovado: 1 },
    approved_for_payment: 1,
    rejected_for_payment: 1,
    milestone_value: 30
  });
  assert.equal(projectClosingMetrics([{ month: '2026-07', decision: 'approved', bonus_value: '80' }], '2026-07').project_value, 80);
  assert.equal(indicationMetrics([{ sale_date: '2026-07-20', value: '50' }], '2026-07').indication_value, 50);
});

test('detects reconciliation failures instead of accepting them silently', () => {
  const result = reconciliationFor({
    portfolio: { total_rows: 2, declared_snapshot_totals: [3] },
    cmax: { consultants_without_rate: 1 },
    financial: { daily_value: 10, milestone_value: 20, project_value: 30, indication_value: 40, total_value: 99 }
  });
  assert.deepEqual(result, {
    component_sum_matches_total: false,
    component_sum: 100,
    portfolio_declared_total_matches_rows: false,
    cmax_has_no_missing_rates: false
  });
});
