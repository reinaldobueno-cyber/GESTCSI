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
  assert.match(html, /PANEL_MIN_2026_PROJECTS = 201/);
  assert.match(html, /PANEL_APP_VERSION = '2026-08-09-map-country-validation-v5'/);
  assert.match(html, /@page\{size:A4 portrait/);
  assert.match(html, /gestcsi_map_geocode_cache_v3_country_validated/);
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

test('validates map geocoding by country and state', () => {
  assert.match(html, /MAP_COUNTRY_PROFILES/);
  assert.match(html, /countryCode: 'br'/);
  assert.match(html, /countryCode: 'py'/);
  assert.match(html, /countrycodes=' \+ profile\.countryCode/);
  assert.match(html, /countryCode !== profile\.countryCode/);
  assert.match(html, /profile\.countryCode === 'br'.+iso\.slice\(-2\) !== info\.uf/);
  assert.match(html, /CUMARU\|PA/);
  assert.match(html, /GOGO DA ONCA\|PA/);
  assert.match(html, /jobs\.push\(info\)/);
  assert.match(html, /setTimeout\(resolve, 1100\)/);
  assert.doesNotMatch(html, /\|\| CIDADES_COORDS\[info\.cidadeNorm\]/);
});

test('uses one authoritative portfolio source and never manufactures a hybrid snapshot', () => {
  assert.match(html, /Aguardando manifesto oficial completo da carteira/);
  assert.match(html, /Fonte oficial indisponível; nenhuma base parcial será promovida/);
  assert.match(html, /Planilha validada:[^\n]+DATA\.projetos\.length[^\n]+projetos/);
  assert.doesNotMatch(html, /_finalizarLoad\(projetosDiretos/);
  assert.doesNotMatch(html, /painelMesclarBaseParcialComPreservada/);
  assert.doesNotMatch(html, /Carga parcial complementada/);
});

test('discovers new spreadsheet rows automatically without overlapping full loads', () => {
  assert.match(html, /PANEL_AUTO_REFRESH_INTERVAL_MS = 5 \* 60 \* 1000/);
  assert.match(html, /function reconciliarCarteiraAutomatica\(gatilho\)/);
  assert.match(html, /apps_script_auto_validated/);
  assert.match(html, /_finalizarLoad\(response\.projetos, true, 'apps_script_auto_validated', 'auto_monitor'\)/);
  assert.match(html, /visibilitychange/);
  assert.match(html, /if \(PANEL_LOAD_IN_FLIGHT\)/);
  assert.doesNotMatch(html, /source: 'google_sheets_client_column'/);
});

test('keeps approved project-closing bonuses readable from the immutable sheet', () => {
  assert.match(html, /function projectClosingLoadDecisionsFromSheet\(\)/);
  assert.match(html, /gid=533338656/);
  assert.match(html, /projectClosingMergeDecisionSources/);
  assert.match(html, /item\.decision === 'approved' \|\| item\.decision === 'rejected'/);
});

test('invalidates expired browser sessions without replacing the portfolio', () => {
  assert.match(html, /function authInvalidateExpiredSession\(\)/);
  assert.match(html, /a carteira íntegra foi preservada/);
  assert.match(html, /Carga parcial bloqueada sem base íntegra de substituição/);
  assert.doesNotMatch(html, /Carga parcial exibida\. A planilha ainda não retornou a base completa/);
});

test('keeps operational CMAX statuses visible while paying only positive events', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');

  assert.match(html, /function cmaxIsPositiveResult\(item\)/);
  assert.match(html, /if\(!cmaxIsPositiveResult\(item\)\)return false;/);
  assert.match(html, /Status da agenda/);
  assert.match(html, /Registros exibidos/);
  assert.match(html, /cmaxResultLabel\(item\)/);
  assert.match(appsScript, /A agenda operacional precisa manter todos os resultados do CMAX/);
  assert.match(appsScript, /return item && item\.mes === month;/);
  assert.match(appsScript, /Eventos previstos ainda não possuem resultado no CMAX/);
  assert.match(appsScript, /if \(eventDate && eventSignal !== ''\)/);
  assert.doesNotMatch(appsScript, /item && item\.mes === month && isCmaxPositiveResult_\(item\.resultado\)/);
  assert.doesNotMatch(appsScript, /eventDate && result !== '' && eventSignal !== ''/);
});

test('uses one canonical name for Anita in milestone closing', () => {
  assert.match(html, /key==='ANITA'\|\|key==='ANITA CRISTINA RODRIGUES TAVARES'/);
  assert.match(html, /return 'Anita Cristina Rodrigues Tavares'/);
});
