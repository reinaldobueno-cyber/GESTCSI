import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const html = await readFile(new URL('../index.html', import.meta.url), 'utf8');

test('loads the portfolio integrity module before the application script', () => {
  const moduleIndex = html.indexOf('<script src="src/portfolio-integrity.js"></script>');
  const appIndex = html.indexOf('var DATA = JSON.parse');
  assert.ok(moduleIndex > 0);
  assert.ok(moduleIndex < appIndex);
});

test('keeps critical production invariants in the staging build', () => {
  assert.match(html, /PANEL_MIN_2026_PROJECTS = 197/);
  assert.match(html, /@page\{size:A4 portrait/);
  assert.match(html, /gestcsi_map_geocode_cache_v2_exact/);
  assert.doesNotMatch(html, /Fallback por UF/);
});

test('validates the Apps Script ALL response before accepting it', () => {
  assert.match(html, /PortfolioIntegrity\.validateSnapshot/);
  assert.match(html, /apps_script_all_integrity_failed/);
  assert.match(html, /apps_script_all_validated/);
  assert.match(html, /apps_script_monthly_validated/);
  assert.doesNotMatch(html, /apps_script_all_merged/);
  assert.doesNotMatch(html, /apps_script_monthly_merged/);
});

test('keeps approved project-closing bonuses readable from the immutable sheet', () => {
  assert.match(html, /function projectClosingLoadDecisionsFromSheet\(\)/);
  assert.match(html, /gid=533338656/);
  assert.match(html, /projectClosingMergeDecisionSources/);
  assert.match(html, /item\.decision === 'approved' \|\| item\.decision === 'rejected'/);
});
