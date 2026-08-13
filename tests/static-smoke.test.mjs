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
  assert.match(html, /PANEL_SNAPSHOT_SCHEMA = 2/);
  assert.match(html, /PANEL_APP_VERSION = '2026-08-12-materialized-portfolio-v1'/);
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

test('validates the materialized portfolio sheet and never manufactures a hybrid snapshot', () => {
  assert.match(html, /Lendo aba técnica consolidada da carteira/);
  assert.match(html, /MONTHLY_PORTFOLIO_BROWSER_GID = '1667939534'/);
  assert.match(html, /snapshot_generation/);
  assert.match(html, /snapshot_month_counts/);
  assert.match(html, /sheet_materialized_validated/);
  assert.match(html, /\{mes:'ALL', lean:1, compressed:1\}/);
  assert.match(html, /new DecompressionStream\('gzip'\)/);
  assert.match(html, /\{mes:'ALL', lean:1\}/);
  assert.match(html, /Fonte oficial indisponível; nenhuma base parcial será promovida/);
  assert.match(html, /Planilha validada:[^\n]+DATA\.projetos\.length[^\n]+projetos/);
  assert.match(html, /_finalizarLoad\(resp\.projetos, isRefresh, 'apps_script_all_validated'/);
  assert.doesNotMatch(html, /painelMesclarBaseParcialComPreservada/);
  assert.doesNotMatch(html, /Carga parcial complementada/);
});

test('allows slow monthly tabs without letting the watchdog erase a valid snapshot', () => {
  assert.match(html, /loadGvizTableJsonp\(url, 35000\)/);
  assert.match(html, /Watchdog preservou o snapshot íntegro já renderizado/);
  assert.match(html, /\}, 150000\);/);
});

test('discovers new spreadsheet rows automatically without overlapping full loads', () => {
  assert.match(html, /PANEL_AUTO_REFRESH_INTERVAL_MS = 5 \* 60 \* 1000/);
  assert.match(html, /function reconciliarCarteiraAutomatica\(gatilho\)/);
  assert.match(html, /sheet_materialized_auto_validated/);
  assert.match(html, /_finalizarLoad\(response, true, 'sheet_materialized_auto_validated', 'auto_monitor'\)/);
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

test('shows the same consultant follow-up notifications to every authenticated viewer', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  const start = appsScript.indexOf('function getProjectFollowups_(');
  const end = appsScript.indexOf('\nfunction findProjectFollowupRow_', start);
  const source = appsScript.slice(start, end);
  assert.match(source, /requireUser_\(params\)/);
  assert.match(source, /return getSharedProjectFollowups_\(limit\)/);
  assert.doesNotMatch(source, /canUserAccessProjectItem_/);
  assert.match(appsScript, /function getSharedProjectFollowups_\(limit\)/);
  assert.match(appsScript, /kanban_states: getProjectKanbanStates_\(\)/);
});

test('invalidates expired browser sessions without replacing the portfolio', () => {
  assert.match(html, /function authInvalidateExpiredSession\(\)/);
  assert.match(html, /a carteira íntegra foi preservada/);
  assert.match(html, /Carga parcial bloqueada sem base íntegra de substituição/);
  assert.doesNotMatch(html, /Carga parcial exibida\. A planilha ainda não retornou a base completa/);
});

test('stops ClickUp actions immediately when the panel session expires', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  assert.match(appsScript, /if \(action === 'syncAll'\) \{\s+requireAdmin_\(params\)/);
  assert.match(appsScript, /sessionCache_\(\)\.put\('session:' \+ token, raw, 21600\)/);
  const syncStart = html.indexOf('window.sincronizarTodosProjetosClickup = function');
  const syncEnd = html.indexOf('\nfunction sleep(', syncStart);
  const syncSource = html.slice(syncStart, syncEnd);
  assert.match(syncSource, /Sessão expirada\. Entre novamente para iniciar o Sync ClickUp/);
  assert.doesNotMatch(syncSource, /sincronizarTodosProjetosClickupEmLotes\(/);
});

test('keeps ClickUp background status reads lightweight and server driven', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  assert.match(appsScript, /action === 'getProjectSyncBackgroundStatus'[\s\S]{0,180}getProjectSyncBackgroundStatus_\(\)/);
  assert.doesNotMatch(appsScript, /action === 'getProjectSyncBackgroundStatus'[\s\S]{0,180}advanceProjectSyncBackgroundFromStatus_/);
  assert.match(appsScript, /CLICKUP_PROJECT_SYNC_BATCH_SIZE', '20'/);
  assert.match(appsScript, /execution_deadline_ms: new Date\(\)\.getTime\(\) \+ 160000/);
  assert.match(appsScript, /batch_total: attempted/);
});

test('resumes the adoption estimate with a lightweight ClickUp reader', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  assert.match(appsScript, /function fetchProjectTasksForActivity_\(/);
  assert.match(appsScript, /fetchAllListTasksForActivity_/);
  assert.match(appsScript, /var payload = fetchProjectTasksForActivity_\(mapping/);
  assert.match(appsScript, /CLICKUP_ACTIVITY_BACKGROUND_BATCH_SIZE', '30'/);
  assert.match(appsScript, /projects_attempted: projectsAttempted/);
  assert.match(appsScript, /scanOffset \+ attemptedInBatch/);
  assert.match(appsScript, /recentComplete && !forceRestart/);
  assert.match(appsScript, /function readClickUpUserActivityProgress_\(/);
  assert.match(appsScript, /getRange\(1, 1, 2, lastColumn\)\.getValues\(\)/);
  assert.match(html, /chamarAppsScriptJsonp\('getClickUpUserActivityStatus'/);
  assert.match(html, /force_restart:'1'/);
  assert.match(html, /if\(resp\.already_complete\)/);
  assert.doesNotMatch(html, /force_estimated:'1',[\s\S]{0,100}scan_batch_size:'1'/);
});

test('does not restart a recent completed estimate unless force is explicit', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  const start = appsScript.indexOf('function startClickUpUserActivityBackground_(');
  const end = appsScript.indexOf('\nfunction continueClickUpUserActivityBackgroundTrigger(', start);
  const source = appsScript.slice(start, end);
  function run(params) {
    const state = new Map([['CLICKUP_ACTIVITY_BACKGROUND_ACTIVE', '0']]);
    const props = {
      getProperty: key => state.get(key) || '',
      setProperty: (key, value) => state.set(key, value),
      deleteProperty: key => state.delete(key)
    };
    let writes = 0;
    let schedules = 0;
    const factory = new Function(
      'PropertiesService', 'readClickUpUserActivityRows_', 'clearClickUpUserActivityBackgroundTriggers_',
      'writeClickUpUserActivitySummary_', 'scheduleClickUpUserActivityBackground_',
      `return (${source.replace(/^function startClickUpUserActivityBackground_/, 'function')});`
    );
    const fn = factory(
      { getScriptProperties: () => props },
      () => [{ sincronizacao_completa_controle: 'sim', sincronizado_em: new Date().toISOString() }],
      () => {},
      () => { writes += 1; },
      () => { schedules += 1; }
    );
    return { result: fn(params), writes, schedules };
  }
  const normal = run({});
  assert.equal(normal.result.already_complete, true);
  assert.equal(normal.writes, 0);
  assert.equal(normal.schedules, 0);
  const forced = run({ force_restart: '1' });
  assert.equal(forced.result.scheduled, true);
  assert.equal(forced.writes, 1);
  assert.equal(forced.schedules, 1);
});

test('queues ClickUp sync and estimate instead of reporting mutual exclusion as an error', async () => {
  const appsScript = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  assert.match(appsScript, /CLICKUP_PROJECT_SYNC_PENDING/);
  assert.match(appsScript, /CLICKUP_ACTIVITY_BACKGROUND_PENDING/);
  assert.match(appsScript, /function startPendingProjectSyncIfAny_\(/);
  assert.match(appsScript, /function startPendingClickUpUserActivityIfAny_\(/);
  assert.match(appsScript, /startPendingProjectSyncIfAny_\(props\)/);
  assert.match(appsScript, /startPendingClickUpUserActivityIfAny_\(props\)/);
  assert.match(appsScript, /function preservePreQueueProjectSyncRequest_\(/);
  assert.match(appsScript, /CLICKUP_QUEUE_MIGRATION_V269/);
  assert.match(appsScript, /preservePreQueueProjectSyncRequest_\(props, complete\)/);
  assert.doesNotMatch(appsScript, /error: 'A estimativa de adoção ClickUp está em andamento/);
  assert.doesNotMatch(appsScript, /error: 'O Sync ClickUp está em andamento/);
  assert.match(html, /Sync ClickUp adicionado à fila/);
  assert.match(html, /Estimativa adicionada à fila/);
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
