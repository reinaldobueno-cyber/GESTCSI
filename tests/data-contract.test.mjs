import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const contract = JSON.parse(await readFile(new URL('../governance/data-contract.json', import.meta.url), 'utf8'));
const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');

const ownerIds = new Set(Object.keys(contract.owners));
const sourceIds = new Set(contract.sources.map((source) => source.id));
const requiredViews = new Set([
  'dashboard', 'executivo', 'gestao', 'acompanhamento', 'cmax', 'fechamento',
  'projetos', 'consultores', 'vendedores', 'avulso_novo', 'logistica', 'prazos',
  'alertas', 'bonificacao', 'administracao'
]);

test('keeps unique, fully owned data sources with SLA and fallback', () => {
  assert.equal(contract.sources.length, sourceIds.size);
  assert.ok(contract.sources.length >= 20);
  for (const source of contract.sources) {
    for (const field of ['id', 'label', 'system', 'classification', 'cadence_current', 'freshness_slo', 'fallback', 'target_backend']) {
      assert.ok(String(source[field] || '').trim(), `${source.id}: ${field} ausente`);
    }
    assert.ok(ownerIds.has(source.owner_operational), `${source.id}: owner operacional inválido`);
    assert.ok(ownerIds.has(source.owner_technical), `${source.id}: owner técnico inválido`);
    assert.ok(Array.isArray(source.consumers) && source.consumers.length, `${source.id}: consumidor ausente`);
  }
});

test('maps every metric to valid sources, owners, SLA, fallback and views', () => {
  const metricIds = new Set(contract.metrics.map((metric) => metric.id));
  assert.equal(contract.metrics.length, metricIds.size);
  assert.ok(contract.metrics.length >= 50);
  for (const metric of contract.metrics) {
    assert.ok(metric.label && metric.area, `${metric.id}: descrição ausente`);
    assert.ok(ownerIds.has(metric.owner_operational), `${metric.id}: owner operacional inválido`);
    assert.ok(ownerIds.has(metric.owner_technical), `${metric.id}: owner técnico inválido`);
    assert.ok(metric.freshness_slo && metric.fallback, `${metric.id}: SLA/fallback ausente`);
    assert.ok(Array.isArray(metric.source_ids) && metric.source_ids.length, `${metric.id}: fonte ausente`);
    metric.source_ids.forEach((id) => assert.ok(sourceIds.has(id), `${metric.id}: fonte desconhecida ${id}`));
    assert.ok(Array.isArray(metric.consumer_views) && metric.consumer_views.length, `${metric.id}: tela consumidora ausente`);
  }
});

test('covers all current navigation areas with cataloged metrics', () => {
  const covered = new Set(contract.metrics.flatMap((metric) => metric.consumer_views));
  requiredViews.forEach((view) => assert.ok(covered.has(view), `tela sem métrica catalogada: ${view}`));
});

test('inventories every Apps Script action exposed by the current router', () => {
  const discovered = [...appsScript.matchAll(/action === '([^']+)'/g)].map((match) => match[1]);
  const expected = [...new Set(discovered)].sort();
  const cataloged = [...new Set(contract.interfaces.apps_script_actions)].sort();
  assert.deepEqual(cataloged, expected);
});

test('catalogs configuration classes without persisting secret values', () => {
  assert.ok(contract.configuration.length >= 14);
  for (const item of contract.configuration) {
    assert.ok(item.name && item.class && item.owner && item.rotation && item.target);
    assert.ok(ownerIds.has(item.owner), `${item.name}: owner inválido`);
    for (const forbidden of ['value', 'secret', 'token', 'password', 'credential']) {
      assert.equal(Object.hasOwn(item, forbidden), false, `${item.name}: valor sensível não pode ser catalogado`);
    }
  }
  assert.equal(contract.policy.no_secret_values, true);
});
