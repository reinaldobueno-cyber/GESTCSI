/**
 * ClickUp -> Google Sheets sync for the PMO panel.
 *
 * What it does:
 * 1. Reads project mappings from CLICKUP_CONFIG.
 * 2. Fetches tasks from ClickUp using list_id (preferred) or view_id.
 * 3. Builds a normalized project summary JSON.
 * 4. Writes summary columns back to the monthly sheet row.
 * 5. Supports webhook + time trigger driven refresh.
 *
 * Required Script Properties:
 * - CLICKUP_TOKEN
 * - SHEET_ID
 *
 * Optional Script Properties:
 * - CLICKUP_CONFIG_SHEET        default: CLICKUP_CONFIG
 * - CLICKUP_WEBHOOKS_SHEET      default: CLICKUP_WEBHOOKS
 * - CLICKUP_DIRTY_SHEET         default: CLICKUP_DIRTY_QUEUE
 * - CLICKUP_SYNC_BATCH_SIZE     default: 10
 * - CLICKUP_WEBHOOK_ENDPOINT    full deployed Apps Script web app URL
 * - CLICKUP_TEAM_ID             required for webhook registration and user activity audit
 * - CLICKUP_WEBHOOK_TOKEN       shared token appended to webhook endpoint
 * - CLICKUP_USER_ACTIVITY_SHEET default: CLICKUP_USER_ACTIVITY
 * - CLICKUP_AUDIT_LOG_SHEET     default: CLICKUP_AUDIT_LOGS
 * - CLICKUP_ACTIVITY_DAYS       default: 90
 */

var CLICKUP_API_BASE = 'https://api.clickup.com/api/v2';
var CLICKUP_DEFAULT_WORKSPACE_ID = '9007083069';
var CLICKUP_MILESTONE_BONUS_VALUE = 30;
var CLICKUP_PROJECT_CLOSING_BONUS_VALUE = 80;
var CLICKUP_PROJECT_CLOSING_BONUS_START = '2026-06-15';
var CLICKUP_PROJECT_CLOSING_RULE_VERSION = 'breakoff-entrega-id-aprovar-v7';
var CLICKUP_PROJECT_DELIVERY_CUSTOM_ITEM_IDS = ['1001'];
var CLICKUP_MILESTONE_AUDIT_TASK_IDS = [];
var CLICKUP_MILESTONE_CLOSING_SCHEMA_VERSION = 'strict-milestones-v2';
var MONTHS = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
var HISTORICAL_CLICKUP_SPACES = [
  { name: 'CSI-PROJETOS-ENGORDA', space_id: '90130063112', url: 'https://app.clickup.com/9007083069/v/s/90130063112' },
  { name: 'CSI-PROJETOS-CICLO COMPLETO', space_id: '90130064659', url: 'https://app.clickup.com/9007083069/v/s/90130064659' },
  { name: 'CSI-PROJETOS-ELITE', space_id: '90130063158', url: 'https://app.clickup.com/9007083069/v/s/90130063158' },
  { name: 'CSI-PROJETOS-CRIA', space_id: '90130063122', url: 'https://app.clickup.com/9007083069/v/s/90130063122' },
  { name: 'CSI-AVULSOS', space_id: '90139026911', url: 'https://app.clickup.com/9007083069/v/s/90139026911' }
];
var OUTPUT_COLUMNS = [
  'tasks_concluidas',
  'tasks_pendentes',
  'marcos_concluidos',
  'marcos_pendentes',
  'fases_total',
  'progresso',
  'data_ultima_atualizacao',
  'dias_sem_atualizacao',
  'link_projeto',
  'view_id',
  'list_id',
  'clickup_json',
  'ultima_sync_clickup',
  'sync_status_clickup',
  'sync_error_clickup'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ClickUp Sync')
    .addItem('Validar CLICKUP_CONFIG', 'validarClickUpConfig')
    .addItem('Importar projetos dos spaces históricos', 'inserirSpacesHistoricosClickUp')
    .addItem('Sincronizar inventário ClickUp', 'sincronizarInventarioClickUp')
    .addItem('Atualizar fechamento de marcos', 'sincronizarFechamentoMarcosClickUp')
    .addItem('Restaurar fechamento pelo histórico mensal', 'restaurarFechamentoMarcosHistoricoMensal')
    .addItem('Sincronizar atividade dos usuários ClickUp', 'sincronizarAtividadeUsuariosClickUp')
    .addItem('Sincronizar diárias CMAX do mês atual', 'sincronizarDiariasCmaxMesAtual')
    .addItem('Sincronizar histórico de diárias CMAX', 'sincronizarHistoricoDiariasCmax')
    .addItem('Reconstruir visão rápida de diárias CMAX', 'reconstruirVisaoDiariasCmax')
    .addItem('Diagnosticar falhas de sync ClickUp', 'diagnosticarFalhasSyncClickUp')
    .addItem('Sincronizar todos habilitados', 'syncAllProjectsTrigger')
    .addItem('Sincronizar primeiro configurado', 'syncPrimeiroProjetoConfigurado')
    .addToUi();
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = String(params.action || '').trim();

  try {
    if (!action && params.mes) {
      return jsonOutput_(getMonthlyProjectsPayload_(params), params.callback);
    }
    if (action === 'getMonthlyProjects') {
      return jsonOutput_(getMonthlyProjectsPayload_(params), params.callback);
    }
    if (action === 'syncProject') {
      var projectKey = String(params.project_key || '').trim();
      var result = syncProjectByKey(projectKey);
      result.ok = true;
      return jsonOutput_(result, params.callback);
    }
    if (action === 'syncAll') {
      var allResult = syncAllProjects({
        force: String(params.force || '') === '1',
        limit: toInt_(params.limit, null),
        offset: toInt_(params.offset, 0)
      });
      return jsonOutput_(allResult, params.callback);
    }
    if (action === 'startProjectSyncBackground') {
      requireAdmin_(params);
      return jsonOutput_(startProjectSyncBackground_(params), params.callback);
    }
    if (action === 'getProjectSyncBackgroundStatus') {
      requireAdmin_(params);
      return jsonOutput_(advanceProjectSyncBackgroundFromStatus_(), params.callback);
    }
    if (action === 'startProjectClosingSync') {
      requireAdmin_(params);
      return jsonOutput_(startProjectClosingSyncBackground_(), params.callback);
    }
    if (action === 'getProjectClosingSyncStatus') {
      requireUser_(params);
      return jsonOutput_(advanceProjectClosingSyncBackgroundFromStatus_(), params.callback);
    }
    if (action === 'stopProjectClosingSync') {
      requireAdmin_(params);
      return jsonOutput_(stopProjectClosingSyncBackground_(), params.callback);
    }
    if (action === 'processDirty') {
      var dirtyResult = processDirtyQueue({
        limit: toInt_(params.limit, null)
      });
      return jsonOutput_(dirtyResult, params.callback);
    }
    if (action === 'validateConfig') {
      return jsonOutput_(validarClickUpConfig(), params.callback);
    }
    if (action === 'getClickUpInventory') {
      return jsonOutput_(getClickUpInventory_(params), params.callback);
    }
    if (action === 'getClickUpMilestoneClosing') {
      return jsonOutput_(getClickUpMilestoneClosing_(params), params.callback);
    }
    if (action === 'diagnoseProjectClosing') {
      return jsonOutput_(diagnoseProjectClosing_(params), params.callback);
    }
    if (action === 'diagnoseClickUpMilestoneTask') {
      requireAdmin_(params);
      return jsonOutput_(diagnoseClickUpMilestoneTask_(params), params.callback);
    }
    if (action === 'syncClickUpMilestoneTask') {
      requireAdmin_(params);
      return jsonOutput_(syncClickUpMilestoneTask_(params), params.callback);
    }
    if (action === 'getProjectClosingDecisions') {
      return jsonOutput_(getProjectClosingDecisions_(params), params.callback);
    }
    if (action === 'getProjectClosingCandidates') {
      return jsonOutput_(getProjectClosingCandidates_(params), params.callback);
    }
    if (action === 'setProjectClosingDecision') {
      return jsonOutput_(setProjectClosingDecision_(params), params.callback);
    }
    if (action === 'startClickUpMilestoneClosingBackground') {
      requireAdmin_(params);
      return jsonOutput_(startClickUpMilestoneClosingBackground_(params), params.callback);
    }
    if (action === 'restoreMilestoneClosingFromMonthlyHistory') {
      requireAdmin_(params);
      return jsonOutput_(restoreClickUpMilestoneClosingFromMonthlyHistory_(params), params.callback);
    }
    if (action === 'syncClickUpMilestoneRecent') {
      requireAdmin_(params);
      return jsonOutput_(syncClickUpRecentMilestoneAndGetClosing_(params), params.callback);
    }
    if (action === 'confirmClickUpMilestoneStatuses') {
      requireAdmin_(params);
      return jsonOutput_(confirmClickUpMilestoneStatuses_(params), params.callback);
    }
    if (action === 'syncClickUpClosedMilestones') {
      requireAdmin_(params);
      return jsonOutput_(syncClickUpClosedMilestones_(params), params.callback);
    }
    if (action === 'syncClickUpApprovedMilestones') {
      requireAdmin_(params);
      return jsonOutput_(syncClickUpValidationSituation_(params, 'aprovado'), params.callback);
    }
    if (action === 'syncClickUpRejectedMilestones') {
      requireAdmin_(params);
      return jsonOutput_(syncClickUpValidationSituation_(params, 'reprovado'), params.callback);
    }
    if (action === 'stopLegacyClickUpMilestoneAudit') {
      requireAdmin_(params);
      return jsonOutput_(stopLegacyClickUpMilestoneAudit_(), params.callback);
    }
    if (action === 'syncClickUpUserActivity') {
      requireAdmin_(params);
      return jsonOutput_(syncClickUpUserActivity_(params), params.callback);
    }
    if (action === 'startClickUpUserActivityBackground') {
      requireAdmin_(params);
      return jsonOutput_(startClickUpUserActivityBackground_(params), params.callback);
    }
    if (action === 'getClickUpUserActivity') {
      requireAdmin_(params);
      return jsonOutput_(getClickUpUserActivity_(params), params.callback);
    }
    if (action === 'getCmaxDailyEvents') {
      return jsonOutput_(getCmaxDailyEvents_(params), params.callback);
    }
    if (action === 'getConsultantCompensation') {
      return jsonOutput_(getConsultantCompensation_(params), params.callback);
    }
    if (action === 'setConsultantSeniority') {
      return jsonOutput_(setConsultantSeniority_(params), params.callback);
    }
    if (action === 'getBonusSalesIndications') {
      requireAdmin_(params);
      return jsonOutput_(getBonusSalesIndications_(params), params.callback);
    }
    if (action === 'saveBonusSalesIndication') {
      requireAdmin_(params);
      return jsonOutput_(saveBonusSalesIndication_(params), params.callback);
    }
    if (action === 'deleteBonusSalesIndication') {
      requireAdmin_(params);
      return jsonOutput_(deleteBonusSalesIndication_(params), params.callback);
    }
    if (action === 'getCmaxDailyHistoryStatus') {
      return jsonOutput_(getCmaxDailyHistoryStatus_(), params.callback);
    }
    if (action === 'syncCmaxDailyEvents') {
      requireAdmin_(params);
      return jsonOutput_(syncCmaxDailyEvents_(params), params.callback);
    }
    if (action === 'startCmaxDailyHistoryBackground') {
      requireAdmin_(params);
      return jsonOutput_(startCmaxDailyHistoryBackground_(params), params.callback);
    }
    if (action === 'continueCmaxDailyHistoryBatch') {
      requireAdmin_(params);
      return jsonOutput_(continueCmaxDailyHistoryBatch_(params), params.callback);
    }
    if (action === 'logPanelUpdate' || String(params.log_update || '') === '1') {
      var logResult = logPanelUpdate_(params);
      return jsonOutput_(logResult, params.callback);
    }
    if (action === 'getPanelUpdateHistory' || String(params.history || '') === '1') {
      var historyResult = getPanelUpdateHistory_(toInt_(params.limit, 20));
      return jsonOutput_(historyResult, params.callback);
    }
    if (action === 'login') {
      return jsonOutput_(loginUser_(params), params.callback);
    }
    if (action === 'me') {
      return jsonOutput_(getCurrentUser_(params), params.callback);
    }
    if (action === 'listUsers') {
      return jsonOutput_(listUsers_(params), params.callback);
    }
    if (action === 'createUser') {
      return jsonOutput_(createUser_(params), params.callback);
    }
    if (action === 'setUserEnabled') {
      return jsonOutput_(setUserEnabled_(params), params.callback);
    }
    if (action === 'setUserSeniority') {
      return jsonOutput_(setUserSeniority_(params), params.callback);
    }
    if (action === 'logProjectFollowup') {
      var followupResult = logProjectFollowup_(params);
      return jsonOutput_(followupResult, params.callback);
    }
    if (action === 'getProjectFollowups') {
      var followupsResult = getProjectFollowups_(params, toInt_(params.limit, 1000));
      return jsonOutput_(followupsResult, params.callback);
    }
    if (action === 'setProjectFollowupStatus') {
      var followupStatusResult = setProjectFollowupStatus_(params);
      return jsonOutput_(followupStatusResult, params.callback);
    }
    if (action === 'setProjectKanbanStage') {
      var kanbanResult = setProjectKanbanStage_(params);
      return jsonOutput_(kanbanResult, params.callback);
    }
    if (action === 'deleteProjectFollowup') {
      var deleteFollowupResult = deleteProjectFollowup_(params);
      return jsonOutput_(deleteFollowupResult, params.callback);
    }

    var payload = {
      ok: true,
      service: 'clickup-sync',
      message: 'Use action=getMonthlyProjects|syncProject|syncAll|processDirty|validateConfig|getClickUpInventory|getClickUpMilestoneClosing|startClickUpMilestoneClosingBackground|syncClickUpUserActivity|startClickUpUserActivityBackground|getClickUpUserActivity|getCmaxDailyEvents|getCmaxDailyHistoryStatus|syncCmaxDailyEvents|startCmaxDailyHistoryBackground|logPanelUpdate|getPanelUpdateHistory|login|me|listUsers|createUser|setUserEnabled|logProjectFollowup|getProjectFollowups|setProjectFollowupStatus|setProjectKanbanStage|deleteProjectFollowup'
    };
    return jsonOutput_(payload, params.callback);
  } catch (error) {
    return jsonOutput_({
      ok: false,
      action: action,
      error: simplifyErrorMessage_(error),
      raw_error: error && error.message ? error.message : String(error || ''),
      at: new Date().toISOString()
    }, params.callback);
  }
}

function doPost(e) {
  var rawBody = ((e || {}).postData || {}).contents || '';
  var event = rawBody ? JSON.parse(rawBody) : {};
  if (!verifyWebhookRequest_(e, rawBody)) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'invalid_signature' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var dirty = enqueueDirtyEvent_(event);
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    queued: dirty.queued,
    event: event.event || ''
  })).setMimeType(ContentService.MimeType.JSON);
}

function syncAllProjects(options) {
  options = options || {};
  var mappings = loadProjectSyncMappings_();
  var offset = Math.max(0, toInt_(options.offset, 0) || 0);
  var limit = options.limit || mappings.length;
  var force = !!options.force;
  var processed = [];
  var errors = [];

  mappings.slice(offset, offset + limit).forEach(function(mapping) {
    try {
      processed.push(syncProjectMapping_(mapping, {
        force: force,
        deadline_ms: new Date().getTime() + Math.max(15000, Math.min(toInt_(options.project_timeout_ms, 45000), 90000))
      }));
    } catch (error) {
      errors.push({
        project_key: mapping.project_key,
        error: error.message
      });
      writeSyncStatus_(mapping, 'error', error.message);
    }
  });

  return {
    ok: errors.length === 0,
    offset: offset,
    total: mappings.length,
    batch_total: Math.min(limit, Math.max(0, mappings.length - offset)),
    total_available: mappings.length,
    next_offset: Math.min(mappings.length, offset + limit),
    has_more: offset + limit < mappings.length,
    done: offset + limit >= mappings.length,
    processed: processed,
    errors: errors
  };
}

function startProjectSyncBackground_(params) {
  params = params || {};
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_PROJECT_SYNC_ACTIVE') === '1') {
    props.setProperty('CLICKUP_PROJECT_SYNC_UPDATED_AT', new Date().toISOString());
    scheduleProjectSyncBackground_(1000);
    var activeStatus = getProjectSyncBackgroundStatus_();
    activeStatus.already_active = true;
    return activeStatus;
  }
  props.setProperty('CLICKUP_PROJECT_SYNC_ACTIVE', '1');
  props.setProperty('CLICKUP_PROJECT_SYNC_OFFSET', '0');
  props.setProperty('CLICKUP_PROJECT_SYNC_TOTAL', '0');
  props.setProperty('CLICKUP_PROJECT_SYNC_PROCESSED', '0');
  props.setProperty('CLICKUP_PROJECT_SYNC_ERRORS', '0');
  props.setProperty('CLICKUP_PROJECT_SYNC_STARTED_AT', new Date().toISOString());
  props.setProperty('CLICKUP_PROJECT_SYNC_UPDATED_AT', new Date().toISOString());
  props.deleteProperty('CLICKUP_PROJECT_SYNC_ERROR');
  props.deleteProperty('CLICKUP_PROJECT_SYNC_COMPLETED_AT');
  scheduleProjectSyncBackground_(1000);
  var status = getProjectSyncBackgroundStatus_();
  status.initializing = true;
  return status;
}

function continueProjectSyncBackgroundTrigger() {
  continueProjectSyncBackgroundStepWithLock_();
}

function advanceProjectSyncBackgroundFromStatus_() {
  return continueProjectSyncBackgroundStepWithLock_({ status_poll: true });
}

function continueProjectSyncBackgroundStepWithLock_(options) {
  options = options || {};
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_PROJECT_SYNC_ACTIVE') !== '1') {
    clearProjectSyncBackgroundTriggers_();
    return getProjectSyncBackgroundStatus_();
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(options.status_poll ? 1000 : 5000)) {
    var busy = getProjectSyncBackgroundStatus_();
    busy.busy = true;
    return busy;
  }
  try {
    normalizeProjectSyncBackgroundQueue_(props);
    return continueProjectSyncBackgroundStep_();
  } finally {
    lock.releaseLock();
  }
}

function continueProjectSyncBackgroundStep_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_PROJECT_SYNC_ACTIVE') !== '1') {
    clearProjectSyncBackgroundTriggers_();
    return getProjectSyncBackgroundStatus_();
  }
  try {
    var offset = toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_OFFSET'), 0);
    var batchSize = Math.max(1, Math.min(toInt_(getScriptProperty_('CLICKUP_PROJECT_SYNC_BATCH_SIZE', '2'), 2), 2));
    var result = syncAllProjects({
      force: true,
      offset: offset,
      limit: batchSize,
      project_timeout_ms: toInt_(getScriptProperty_('CLICKUP_PROJECT_SYNC_PROJECT_TIMEOUT_MS', '30000'), 30000)
    });
    var attempted = Number(result.batch_total || 0) || (result.processed.length + result.errors.length);
    var processed = toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_PROCESSED'), 0) + attempted;
    var errors = toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_ERRORS'), 0) + result.errors.length;
    props.setProperty('CLICKUP_PROJECT_SYNC_TOTAL', String(result.total || 0));
    props.setProperty('CLICKUP_PROJECT_SYNC_OFFSET', String(result.next_offset));
    props.setProperty('CLICKUP_PROJECT_SYNC_PROCESSED', String(processed));
    props.setProperty('CLICKUP_PROJECT_SYNC_ERRORS', String(errors));
    props.setProperty('CLICKUP_PROJECT_SYNC_UPDATED_AT', new Date().toISOString());
    props.deleteProperty('CLICKUP_PROJECT_SYNC_ERROR');
    if (result.done) {
      props.setProperty('CLICKUP_PROJECT_SYNC_ACTIVE', '0');
      props.setProperty('CLICKUP_PROJECT_SYNC_COMPLETED_AT', new Date().toISOString());
      clearProjectSyncBackgroundTriggers_();
      return getProjectSyncBackgroundStatus_();
    }
    scheduleProjectSyncBackground_(8000);
    return getProjectSyncBackgroundStatus_();
  } catch (error) {
    props.setProperty('CLICKUP_PROJECT_SYNC_ERROR', simplifyErrorMessage_(error));
    props.setProperty('CLICKUP_PROJECT_SYNC_ACTIVE', '0');
    clearProjectSyncBackgroundTriggers_();
    return getProjectSyncBackgroundStatus_();
  }
}

function normalizeProjectSyncBackgroundQueue_(props) {
  props = props || PropertiesService.getScriptProperties();
  var expectedTotal = loadProjectSyncMappings_().length;
  var currentTotal = toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_TOTAL'), 0);
  if (currentTotal === expectedTotal) return;
  props.setProperty('CLICKUP_PROJECT_SYNC_TOTAL', String(expectedTotal));
  props.setProperty('CLICKUP_PROJECT_SYNC_OFFSET', '0');
  props.setProperty('CLICKUP_PROJECT_SYNC_PROCESSED', '0');
  props.setProperty('CLICKUP_PROJECT_SYNC_ERRORS', '0');
  props.deleteProperty('CLICKUP_PROJECT_SYNC_ERROR');
  props.deleteProperty('CLICKUP_PROJECT_SYNC_COMPLETED_AT');
  props.setProperty('CLICKUP_PROJECT_SYNC_UPDATED_AT', new Date().toISOString());
}

function getProjectSyncBackgroundStatus_() {
  var props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    active: props.getProperty('CLICKUP_PROJECT_SYNC_ACTIVE') === '1',
    total: toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_TOTAL'), 0),
    processed: toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_PROCESSED'), 0),
    offset: toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_OFFSET'), 0),
    errors: toInt_(props.getProperty('CLICKUP_PROJECT_SYNC_ERRORS'), 0),
    started_at: props.getProperty('CLICKUP_PROJECT_SYNC_STARTED_AT') || '',
    updated_at: props.getProperty('CLICKUP_PROJECT_SYNC_UPDATED_AT') || '',
    completed_at: props.getProperty('CLICKUP_PROJECT_SYNC_COMPLETED_AT') || '',
    error: props.getProperty('CLICKUP_PROJECT_SYNC_ERROR') || ''
  };
}

function scheduleProjectSyncBackground_(delayMs) {
  clearProjectSyncBackgroundTriggers_();
  ScriptApp.newTrigger('continueProjectSyncBackgroundTrigger')
    .timeBased()
    .after(Math.max(1000, Number(delayMs || 15000)))
    .create();
}

function clearProjectSyncBackgroundTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'continueProjectSyncBackgroundTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function syncProjectByKey(projectKey) {
  if (!projectKey) throw new Error('project_key is required');
  var mapping = findProjectMapping_(projectKey);
  if (!mapping) throw new Error('Project mapping not found: ' + projectKey);
  return syncProjectMapping_(mapping, { force: true });
}

function diagnosticarFalhasSyncClickUp() {
  var mappings = loadProjectMappings_();
  var rows = [];
  MONTHS.forEach(function(month) {
    var sheet;
    try {
      sheet = getMonthSheet_(month);
    } catch (e) {
      return;
    }
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return;
    var header = normalizeMonthlyHeader_(values[0]);
    values.slice(1).forEach(function(row, index) {
      var project = monthlyProjectFromRow_(month, header, row, index + 2);
      if (!project) return;
      var syncStatus = sanitizeText_(project.sync_status_clickup);
      var syncError = sanitizeText_(project.sync_error_clickup);
      var hasExplicitError = /error|erro|falha/i.test(syncStatus) || !!syncError;
      var expectsClickUp = monthlyProjectExpectsClickUpSummary_(project);
      var hasSummary = monthlyProjectHasClickUpSummary_(project);
      if (!hasExplicitError && (!expectsClickUp || hasSummary)) return;
      rows.push(buildClickUpSyncDiagnosticRow_(project, mappings, hasExplicitError ? 'Falha explicita' : 'Sem resumo apos sync'));
    });
  });

  var headers = [
    'mes', 'linha', 'cliente', 'consultor', 'status_projeto', 'tipo',
    'situacao', 'categoria', 'acao_recomendada',
    'sync_status', 'sync_error',
    'clickup', 'link_linha', 'list_id_linha', 'view_id_linha',
    'config_encontrada', 'config_enabled', 'config_project_key', 'config_cliente',
    'config_list_id', 'config_view_id', 'config_folder_id', 'config_space_id', 'config_url'
  ];
  var sheet = getOrCreateSheet_('CLICKUP_SYNC_DIAGNOSTICO');
  writeObjectsToSheet_(sheet, rows, headers);
  return { ok: true, total: rows.length, sheet: sheet.getName() };
}

function monthlyProjectExpectsClickUpSummary_(project) {
  var clickup = normalizeKey_(project && project.clickup);
  var status = normalizeKey_(project && project.status);
  if (clickup === 'NAO' || status.indexOf('CANCEL') >= 0) return false;
  return clickup === 'SIM' || !!sanitizeText_(project && (project.link_projeto || project.projeto_link));
}

function monthlyProjectHasClickUpSummary_(project) {
  return [
    'tasks_concluidas',
    'tasks_pendentes',
    'marcos_concluidos',
    'marcos_pendentes',
    'fases_total',
    'progresso',
    'data_ultima_atualizacao',
    'clickup_json',
    'ultima_sync_clickup'
  ].some(function(field) {
    return String(project && project[field] !== null && project[field] !== undefined ? project[field] : '').trim() !== '';
  });
}

function buildClickUpSyncDiagnosticRow_(project, mappings, situacao) {
  var rowUrl = sanitizeText_(project.link_projeto || project.projeto_link);
  var rowListId = normalizeClickUpId_(project.list_id) || extractClickUpIdFromUrl_(rowUrl, 'list');
  var rowViewId = normalizeClickUpId_(project.view_id) || extractClickUpIdFromUrl_(rowUrl, 'view');
  var match = findDiagnosticMappingForProject_(project, mappings);
  var detail = sanitizeText_(project.sync_error_clickup || project.sync_status_clickup);
  var categoryAction = classifyClickUpSyncProblem_(project, match && match.mapping, detail);
  var mapping = match && match.mapping || null;
  return {
    mes: project.mes,
    linha: project._sheet_row,
    cliente: project.cliente,
    consultor: project.consultor,
    status_projeto: project.status,
    tipo: project.tipo,
    situacao: situacao,
    categoria: categoryAction.categoria,
    acao_recomendada: categoryAction.acao,
    sync_status: project.sync_status_clickup,
    sync_error: project.sync_error_clickup,
    clickup: project.clickup,
    link_linha: rowUrl,
    list_id_linha: rowListId,
    view_id_linha: rowViewId,
    config_encontrada: mapping ? (match && match.by || 'sim') : 'nao',
    config_enabled: mapping ? (mapping.enabled ? 'sim' : 'nao') : '',
    config_project_key: mapping && mapping.project_key || '',
    config_cliente: mapping && mapping.cliente || '',
    config_list_id: mapping && mapping.list_id || '',
    config_view_id: mapping && mapping.view_id || '',
    config_folder_id: mapping && mapping.folder_id || '',
    config_space_id: mapping && mapping.space_id || '',
    config_url: mapping && mapping.project_url || ''
  };
}

function findDiagnosticMappingForProject_(project, mappings) {
  var rowUrl = sanitizeText_(project && (project.link_projeto || project.projeto_link));
  var rowListId = normalizeClickUpId_(project && project.list_id) || extractClickUpIdFromUrl_(rowUrl, 'list');
  var rowViewId = normalizeClickUpId_(project && project.view_id) || extractClickUpIdFromUrl_(rowUrl, 'view');
  var projectClient = normalizeKey_(project && project.cliente);
  var projectMonth = sanitizeMonth_(project && project.mes);
  var exactClient = null;
  var best = null;

  (mappings || []).some(function(mapping) {
    if (sanitizeMonth_(mapping && mapping.mes) !== projectMonth) return false;
    if (projectClient && normalizeKey_(mapping && mapping.cliente) === projectClient) {
      exactClient = { mapping: mapping, by: 'cliente' };
      return true;
    }
    return false;
  });
  if (exactClient) return exactClient;

  (mappings || []).some(function(mapping) {
    var mapUrl = sanitizeText_(mapping && mapping.project_url);
    var mapListId = normalizeClickUpId_(mapping && mapping.list_id) || extractClickUpIdFromUrl_(mapUrl, 'list');
    var mapViewId = normalizeClickUpId_(mapping && mapping.view_id) || extractClickUpIdFromUrl_(mapUrl, 'view');
    if (rowListId && mapListId && rowListId === mapListId) {
      best = { mapping: mapping, by: 'list_id' };
      return true;
    }
    if (rowViewId && mapViewId && rowViewId === mapViewId) {
      best = { mapping: mapping, by: 'view_id' };
      return true;
    }
    if (rowUrl && mapUrl && rowUrl === mapUrl) {
      best = { mapping: mapping, by: 'link' };
      return true;
    }
    return false;
  });
  return best;
}

function classifyClickUpSyncProblem_(project, mapping, detail) {
  var text = normalizeKey_(detail);
  if (!mapping) {
    return {
      categoria: 'Sem configuracao no CLICKUP_CONFIG',
      acao: 'Incluir ou corrigir o projeto no CLICKUP_CONFIG com mes, cliente, enabled=TRUE e link/list_id/view_id correto.'
    };
  }
  if (!mapping.enabled) {
    return {
      categoria: 'Configuracao desabilitada',
      acao: 'Alterar enabled para TRUE no CLICKUP_CONFIG e rodar o Sync ClickUp novamente.'
    };
  }
  if (!(mapping.list_id || mapping.view_id || mapping.folder_id || mapping.space_id || mapping.project_url)) {
    return {
      categoria: 'Configuracao sem ID/link',
      acao: 'Preencher project_url, list_id ou view_id no CLICKUP_CONFIG.'
    };
  }
  if (text.indexOf('50000') >= 0 || text.indexOf('LIMITE MAXIMO') >= 0) {
    return {
      categoria: 'Limite de celula da planilha',
      acao: 'Publicar o Apps Script atualizado com JSON compacto e rodar o Sync ClickUp novamente.'
    };
  }
  if (text.indexOf('FOLDER ID INVALID') >= 0 || text.indexOf('INPUT 011') >= 0) {
    return {
      categoria: 'Folder/lista invalido no cadastro',
      acao: 'Trocar folder_id/list_id/view_id ou project_url por um link valido da lista/view do projeto no ClickUp.'
    };
  }
  if (text.indexOf('LIST ID INVALID') >= 0 || text.indexOf('VIEW ID INVALID') >= 0) {
    return {
      categoria: 'ID ClickUp invalido',
      acao: 'Copiar novamente o link da lista/view correta e atualizar o CLICKUP_CONFIG.'
    };
  }
  if (text.indexOf('CLIENT ROW NOT FOUND') >= 0) {
    return {
      categoria: 'Linha mensal nao localizada',
      acao: 'Publicar o Apps Script atualizado com busca por link/list_id/view_id e rodar o Sync novamente.'
    };
  }
  if (text.indexOf('TOKEN') >= 0 || text.indexOf('401') >= 0) {
    return {
      categoria: 'Token ClickUp',
      acao: 'Renovar CLICKUP_TOKEN nas propriedades do Apps Script.'
    };
  }
  return {
    categoria: detail ? 'Erro ClickUp/sync nao classificado' : 'Sem resumo gravado',
    acao: 'Conferir CLICKUP_CONFIG, publicar o Apps Script atualizado e rodar Sync ClickUp novamente. Se persistir, abrir o erro bruto.'
  };
}

function processDirtyQueue(options) {
  options = options || {};
  var sheet = getDirtyQueueSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { ok: true, processed: [], total: 0 };
  }

  var header = values[0];
  var rows = values.slice(1);
  var limit = options.limit || Number(getScriptProperty_('CLICKUP_SYNC_BATCH_SIZE', '10'));
  var processed = [];
  var keepRows = [header];

  rows.forEach(function(row) {
    if (processed.length >= limit) {
      keepRows.push(row);
      return;
    }
    var item = rowToObject_(header, row);
    try {
      var mapping = null;
      if (item.project_key) mapping = findProjectMapping_(item.project_key);
      if (!mapping && item.list_id) mapping = findProjectMappingByListId_(item.list_id);
      if (!mapping && item.task_id) mapping = findProjectMappingByTaskId_(item.task_id);

      if (!mapping) {
        keepRows.push(row);
        return;
      }

      processed.push(syncProjectMapping_(mapping, { force: true }));
    } catch (error) {
      keepRows.push(row);
    }
  });

  sheet.clearContents();
  sheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
  return {
    ok: true,
    total: processed.length,
    processed: processed
  };
}

function registerAllWebhooks() {
  var endpoint = getScriptProperty_('CLICKUP_WEBHOOK_ENDPOINT');
  var teamId = getScriptProperty_('CLICKUP_TEAM_ID');
  var webhookToken = getScriptProperty_('CLICKUP_WEBHOOK_TOKEN', '');
  if (!endpoint) throw new Error('CLICKUP_WEBHOOK_ENDPOINT is required');
  if (!teamId) throw new Error('CLICKUP_TEAM_ID is required');

  var mappings = loadProjectMappings_().filter(function(item) {
    return item.enabled && item.list_id;
  });
  var webhookSheet = getWebhookSheet_();
  var registered = [];

  mappings.forEach(function(mapping) {
    var body = {
      endpoint: appendQueryParam_(endpoint, 'webhook_token', webhookToken),
      events: [
        'taskCreated',
        'taskUpdated',
        'taskDeleted',
        'taskMoved',
        'taskStatusUpdated'
      ],
      list_id: Number(mapping.list_id)
    };
    var response = clickupRequest_('post', '/team/' + teamId + '/webhook', body);
    registered.push({
      project_key: mapping.project_key,
      list_id: mapping.list_id,
      webhook_id: response.id || response.webhook && response.webhook.id || '',
      secret: response.secret || response.webhook && response.webhook.secret || ''
    });
  });

  writeObjectsToSheet_(webhookSheet, registered, ['project_key', 'list_id', 'webhook_id', 'secret']);
  return { ok: true, total: registered.length, registered: registered };
}

function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('processDirtyQueueTrigger').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('syncAllProjectsTrigger').timeBased().everyHours(6).create();
  installClickUpClosedMilestonesTrigger();
}

function processDirtyQueueTrigger() {
  processDirtyQueue({});
}

function syncAllProjectsTrigger() {
  syncAllProjects({ force: true });
}

function syncClickUpClosedMilestonesTrigger() {
  syncClickUpClosedMilestones_({ skip_result: true });
}

function installClickUpClosedMilestonesTrigger() {
  stopLegacyClickUpMilestoneAudit_();
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncClickUpClosedMilestonesTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('syncClickUpClosedMilestonesTrigger').timeBased().everyHours(12).create();
  return { ok: true, handler: 'syncClickUpClosedMilestonesTrigger', every_hours: 12 };
}

function stopLegacyClickUpMilestoneAudit_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_MILESTONE_CLOSING_PHASE') === 'disabled' &&
      props.getProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE') !== '1') {
    return { ok: true, legacy_audit_disabled: true, already_disabled: true };
  }
  props.setProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE', '0');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'disabled');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === 'syncClickUpMilestoneClosingTrigger' ||
        handler === 'continueClickUpMilestoneClosingBackgroundTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  return { ok: true, legacy_audit_disabled: true };
}

function diagnosticarSetup() {
  var sheetId = getScriptProperty_('SHEET_ID', '');
  var token = getScriptProperty_('CLICKUP_TOKEN', '');
  var ss = sheetId ? SpreadsheetApp.openById(sheetId) : null;
  var mappings = loadProjectMappings_();
  var first = mappings.length ? mappings[0] : null;
  var monthSheet = first ? ss.getSheetByName(first.mes) : null;
  var row = first && monthSheet ? findClientRow_(monthSheet, first.cliente) : null;

  var result = {
    ok: true,
    propriedades: {
      SHEET_ID_preenchido: !!sheetId,
      CLICKUP_TOKEN_preenchido: !!token
    },
    planilha: {
      encontrada: !!ss,
      abas: ss ? ss.getSheets().map(function(s) { return s.getName(); }) : []
    },
    config: {
      total_projetos: mappings.length,
      primeiro_projeto: first || null
    },
    linha_projeto: {
      aba_encontrada: !!monthSheet,
      linha_encontrada: !!row,
      numero_linha: row || null
    }
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function syncPrimeiroProjetoConfigurado() {
  var mappings = loadProjectMappings_().filter(function(item) { return item.enabled; });
  if (!mappings.length) throw new Error('Nenhum projeto habilitado em CLICKUP_CONFIG');
  return syncProjectMapping_(mappings[0], { force: true });
}

function garantirColunasPrimeiroProjeto() {
  var mappings = loadProjectMappings_().filter(function(item) { return item.enabled; });
  if (!mappings.length) throw new Error('Nenhum projeto habilitado em CLICKUP_CONFIG');
  var mapping = mappings[0];
  var sheet = getMonthSheet_(mapping.mes);
  var cols = ensureMonthlyOutputColumns_(sheet);
  Logger.log(JSON.stringify({
    ok: true,
    mes: mapping.mes,
    cliente: mapping.cliente,
    colunas: cols
  }, null, 2));
  return cols;
}

function diagnosticarPrimeiroProjetoClickup() {
  var mappings = loadProjectMappings_().filter(function(item) { return item.enabled; });
  if (!mappings.length) throw new Error('Nenhum projeto habilitado em CLICKUP_CONFIG');
  var mapping = mappings[0];
  var payload = fetchProjectTasks_(mapping);
  var tasks = payload.tasks || [];
  var sample = tasks.slice(0, 15).map(function(task) {
    return {
      id: String(task.id || ''),
      name: sanitizeText_(task.name),
      parent: String(task.parent || ''),
      status: task.status && (task.status.status || task.status.type || task.status.label) || '',
      list_name: task.list && task.list.name || '',
      custom_item_id: String(task.custom_item_id || ''),
      custom_item_name: task.custom_item && task.custom_item.name || ''
    };
  });
  var result = {
    ok: true,
    project_key: mapping.project_key,
    cliente: mapping.cliente,
    source: payload.source,
    view_id: mapping.view_id,
    list_id: mapping.list_id,
    total_tasks: tasks.length,
    sample: sample
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function getMonthlyProjectsPayload_(params) {
  params = params || {};
  var requested = sanitizeMonth_(params.mes || 'ALL');
  var months = requested && requested !== 'ALL' ? [requested] : MONTHS.slice();
  var projetos = [];
  var byMonth = {};
  months.forEach(function(month) {
    if (MONTHS.indexOf(month) < 0) return;
    var rows = getMonthlyProjectsFromSheet_(month);
    byMonth[month] = rows.length;
    projetos = projetos.concat(rows);
  });
  return {
    ok: true,
    mes: requested || 'ALL',
    total: projetos.length,
    projetos_por_mes: byMonth,
    projetos: projetos
  };
}

function getMonthlyProjectsFromSheet_(month) {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(month);
  if (!sheet) return [];
  var values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return [];
  var header = normalizeMonthlyHeader_(values[0] || []);
  var out = [];
  var ultimaDataVenda = '';
  values.slice(1).forEach(function(row, index) {
    var item = monthlyProjectFromRow_(month, header, row, index + 2);
    if (item) {
      if (item.data_venda) ultimaDataVenda = item.data_venda;
      else if (ultimaDataVenda) item.data_venda = ultimaDataVenda;
      out.push(item);
    }
  });
  return out;
}

function monthlyProjectFromRow_(month, header, row, rowNumber) {
  var obj = rowToObject_(header, row);
  function pick(names, fallbackIndex) {
    for (var i = 0; i < names.length; i++) {
      var value = obj[names[i]];
      if (value !== undefined && value !== null && String(value).trim() !== '') return value;
    }
    return fallbackIndex !== undefined && fallbackIndex !== null ? row[fallbackIndex] : '';
  }

  var cliente = sanitizeText_(pick(['cliente', 'nome_cliente', 'nome_do_cliente'], 1));
  if (!isValidMonthlyClient_(cliente)) return null;

  return {
    mes: month,
    mes_origem: month,
    data_venda: pick(['data_venda', 'data_da_venda'], 0),
    cliente: cliente,
    pacote: pick(['pacote'], 2),
    adicionais: pick(['adicionais'], 3),
    tipo: pick(['tipo', 'tipo_projeto'], 4),
    vendedor: pick(['vendedor'], 5),
    consultor: pick(['consultor', 'consultora', 'responsavel'], 7),
    formato: pick(['formato', 'modalidade'], 8),
    cidade: pick(['cidade'], 9),
    data_estimada: pick(['data_estimada', 'previsao', 'mes_estimado'], 10),
    kickoff: pick(['kickoff', 'kick_off'], 11),
    data_kick: pick(['data_kick', 'data_kickoff'], 12),
    clickup: pick(['clickup'], 13),
    data_inicio: pick(['data_inicio', 'primeiro_treinamento'], 14),
    diarias_cont: pick(['diarias_cont', 'diarias_contratadas', 'diarias_total'], 15),
    diarias_real: pick(['diarias_real', 'diarias_realizadas', 'diarias_consumidas'], 16),
    diarias_rest: pick(['diarias_rest', 'diarias_restantes', 'saldo_diarias'], 17),
    acompanhamento: pick(['acompanhamento', 'observacao_acompanhamento'], 18),
    data_enc: pick(['data_enc', 'data_encerramento'], 19),
    avaliacao_consultor: pick(['avaliacao_consultor'], 20),
    status: pick(['status', 'status_projeto'], 21),
    projeto_link: pick(['projeto_link', 'link_projeto', 'link_do_projeto', 'url_projeto'], 22),
    link_projeto: pick(['link_projeto', 'projeto_link', 'link_do_projeto', 'url_projeto'], 22),
    tasks_concluidas: pick(['tasks_concluidas'], null),
    tasks_pendentes: pick(['tasks_pendentes'], null),
    marcos_concluidos: pick(['marcos_concluidos'], null),
    marcos_pendentes: pick(['marcos_pendentes'], null),
    fases_total: pick(['fases_total'], null),
    progresso: pick(['progresso'], null),
    data_ultima_atualizacao: pick(['data_ultima_atualizacao'], null),
    dias_sem_atualizacao: pick(['dias_sem_atualizacao'], null),
    view_id: pick(['view_id'], null),
    list_id: pick(['list_id'], null),
    clickup_json: pick(['clickup_json'], null),
    ultima_sync_clickup: pick(['ultima_sync_clickup'], null),
    sync_status_clickup: pick(['sync_status_clickup'], null),
    sync_error_clickup: pick(['sync_error_clickup'], null),
    _sheet_row: rowNumber
  };
}

function normalizeMonthlyHeader_(header) {
  return (header || []).map(function(name, index) {
    var key = canonicalMonthlyHeader_(name);
    return key || ('col_' + index);
  });
}

function canonicalMonthlyHeader_(name) {
  var key = normalizeKey_(name);
  var aliases = {
    data_venda: ['DATA VENDA', 'DATA DA VENDA'],
    cliente: ['CLIENTE', 'NOME CLIENTE', 'NOME DO CLIENTE'],
    pacote: ['PACOTE'],
    adicionais: ['ADICIONAIS'],
    tipo: ['TIPO', 'TIPO PROJETO'],
    vendedor: ['VENDEDOR'],
    consultor: ['CONSULTOR', 'CONSULTORA', 'RESPONSAVEL'],
    formato: ['FORMATO', 'MODALIDADE'],
    cidade: ['CIDADE'],
    data_estimada: ['DATA ESTIMADA', 'PREVISAO', 'MES ESTIMADO'],
    kickoff: ['KICKOFF', 'KICK OFF'],
    data_kick: ['DATA KICK', 'DATA KICKOFF'],
    clickup: ['CLICKUP'],
    data_inicio: ['DATA INICIO', 'PRIMEIRO TREINAMENTO'],
    diarias_cont: ['DIARIAS CONT', 'DIARIAS CONTRATADAS', 'DIARIAS TOTAL'],
    diarias_real: ['DIARIAS REAL', 'DIARIAS REALIZADAS', 'DIARIAS CONSUMIDAS'],
    diarias_rest: ['DIARIAS REST', 'DIARIAS RESTANTES', 'SALDO DIARIAS'],
    acompanhamento: ['ACOMPANHAMENTO', 'OBSERVACAO ACOMPANHAMENTO'],
    data_enc: ['DATA ENC', 'DATA ENCERRAMENTO'],
    avaliacao_consultor: ['AVALIACAO CONSULTOR'],
    status: ['STATUS', 'STATUS PROJETO'],
    projeto_link: ['PROJETO LINK', 'LINK PROJETO', 'LINK DO PROJETO', 'URL PROJETO'],
    link_projeto: ['LINK PROJETO', 'LINK DO PROJETO', 'PROJETO LINK', 'URL PROJETO'],
    tasks_concluidas: ['TASKS CONCLUIDAS'],
    tasks_pendentes: ['TASKS PENDENTES'],
    marcos_concluidos: ['MARCOS CONCLUIDOS'],
    marcos_pendentes: ['MARCOS PENDENTES'],
    fases_total: ['FASES TOTAL'],
    progresso: ['PROGRESSO'],
    data_ultima_atualizacao: ['DATA ULTIMA ATUALIZACAO'],
    dias_sem_atualizacao: ['DIAS SEM ATUALIZACAO'],
    view_id: ['VIEW ID'],
    list_id: ['LIST ID'],
    clickup_json: ['CLICKUP JSON'],
    ultima_sync_clickup: ['ULTIMA SYNC CLICKUP'],
    sync_status_clickup: ['SYNC STATUS CLICKUP'],
    sync_error_clickup: ['SYNC ERROR CLICKUP']
  };
  var found = '';
  Object.keys(aliases).some(function(canonical) {
    if (aliases[canonical].indexOf(key) >= 0) {
      found = canonical;
      return true;
    }
    return false;
  });
  return found;
}

function isValidMonthlyClient_(cliente) {
  var key = normalizeKey_(cliente);
  if (!key) return false;
  return ['CLIENTE', 'NOME CLIENTE', 'NOME DO CLIENTE', 'TOTAL DE PROJETOS', 'CONSULTOR', 'VENDEDOR', 'STATUS', 'FORMATO', 'TIPO', 'PERIODO', 'MES', 'INDICADORES GERAIS'].indexOf(key) < 0;
}

function syncProjectMapping_(mapping, options) {
  options = options || {};
  mapping = reconcileProjectMappingWithMonthlyLink_(mapping);
  var normalized = buildNormalizedProjectFromClickUp_(mapping, options);
  var hasMonthlySheet = MONTHS.indexOf(sanitizeMonth_(mapping.mes)) >= 0;
  if (hasMonthlySheet) writeProjectSummaryToMonthlySheet_(mapping, normalized);
  if (!options.skip_milestone_closing) upsertClickUpMilestoneClosing_(mapping, normalized);
  reconcileProjectClosingDecisionFromNormalized_(mapping, normalized);
  if (hasMonthlySheet) writeSyncStatus_(mapping, 'ok', '');
  return {
    project_key: mapping.project_key,
    cliente: mapping.cliente,
    mes: mapping.mes,
    tasks_total: normalized.resumo.tasks_total,
    marcos_total: normalized.resumo.marcos_total,
    synced_at: normalized.synced_at,
    project_closing_rule_version: CLICKUP_PROJECT_CLOSING_RULE_VERSION,
    source: normalized.clickup_payload && normalized.clickup_payload.source || '',
    warning: normalized.clickup_payload && normalized.clickup_payload.warning || ''
  };
}

function reconcileProjectMappingWithMonthlyLink_(mapping) {
  if (!mapping || MONTHS.indexOf(sanitizeMonth_(mapping.mes)) < 0) return mapping;
  var rows = getMonthlyProjectsFromSheet_(sanitizeMonth_(mapping.mes));
  var target = normalizeKey_(mapping.cliente);
  var matches = rows.filter(function(item) {
    return normalizeKey_(item && item.cliente) === target;
  });
  if (!matches.length) return mapping;

  var mappingUrl = sanitizeText_(mapping.project_url || mapping.link_projeto || mapping.projeto_link);
  var targetListId = normalizeClickUpId_(mapping.list_id) || extractClickUpIdFromUrl_(mappingUrl, 'list');
  var targetViewId = normalizeClickUpId_(mapping.view_id) || extractClickUpIdFromUrl_(mappingUrl, 'view');
  var targetFolderId = normalizeClickUpNumericId_(mapping.folder_id) || extractClickUpIdFromUrl_(mappingUrl, 'folder');
  var targetSpaceId = normalizeClickUpNumericId_(mapping.space_id) || extractClickUpIdFromUrl_(mappingUrl, 'space');
  var monthly = matches.filter(function(item) {
    var rowUrl = sanitizeText_(item && (item.link_projeto || item.projeto_link));
    return (targetListId && targetListId === extractClickUpIdFromUrl_(rowUrl, 'list')) ||
      (targetViewId && targetViewId === extractClickUpIdFromUrl_(rowUrl, 'view')) ||
      (targetFolderId && targetFolderId === extractClickUpIdFromUrl_(rowUrl, 'folder')) ||
      (targetSpaceId && targetSpaceId === extractClickUpIdFromUrl_(rowUrl, 'space')) ||
      (mappingUrl && rowUrl && sanitizeText_(mappingUrl) === rowUrl);
  })[0] || matches.filter(function(item) {
    return normalizeProjectKey_(item && item.project_key) === normalizeProjectKey_(mapping.project_key);
  })[0] || matches[0];

  var monthlyUrl = sanitizeText_(monthly && (monthly.link_projeto || monthly.projeto_link));
  if (!monthlyUrl || !isRecognizedClickUpSourceUrl_(monthlyUrl)) return mapping;
  return projectMappingFromConfigItem_(Object.assign({}, mapping, {
    project_url: monthlyUrl,
    link_projeto: monthlyUrl
  })) || mapping;
}

function buildNormalizedProjectFromClickUp_(mapping, options) {
  options = options || {};
  var payload = fetchProjectTasks_(mapping, options);
  var tasks = payload.tasks || [];
  var phaseMap = {};
  var byId = {};
  var ignoredNestedItems = 0;

  tasks.forEach(function(task) {
    byId[String(task.id)] = task;
  });

  tasks.forEach(function(task) {
    if (isPhaseTask_(task, byId)) {
      phaseMap[String(task.id)] = {
        tipo: 'fase',
        id: String(task.id),
        nome: sanitizeText_(task.name),
        ordem: extractLeadingNumber_(task.name),
        status_original: task.status && (task.status.status || task.status.type || task.status.label) || '',
        custom_item_id: String(task.custom_item_id || ''),
        custom_item_name: clickUpTaskCustomItemName_(task),
        marcador_entrega: isProjectDeliveryTask_(task) ? 'sim' : '',
        marcador_marco: isMilestoneTask_(task) ? 'sim' : '',
        updated_at: task.date_updated ? fromMillisIso_(task.date_updated) : '',
        task_url: sanitizeText_(task.url || task.permalink || task.link || task.html_url) ||
          ('https://app.clickup.com/t/' + String(task.id || '')),
        tasks_concluidas: 0,
        tasks_pendentes: 0,
        marcos_concluidos: 0,
        marcos_pendentes: 0
      };
    }
  });

  var normalizedTasks = [];
  var normalizedMilestones = [];

  tasks.forEach(function(task) {
    if (isPhaseTask_(task, byId)) {
      var phase = phaseMap[String(task.id)];
      if (phase && isProjectClosingDeliveryItem_(task, phase.nome, phase)) {
        var phaseMilestone = {
          id: String(task.id),
          tipo: 'marco',
          nome: sanitizeText_(task.name),
          fase_nome: phase.nome,
          parent_id: String(task.parent || ''),
          status_original: task.status && (task.status.status || task.status.type || task.status.label) || '',
          fase_status_original: phase.status_original || '',
          custom_item_id: String(task.custom_item_id || ''),
          custom_item_name: clickUpTaskCustomItemName_(task),
          marcador_entrega: isProjectDeliveryTask_(task) ? 'sim' : '',
          responsaveis: (task.assignees || []).map(function(user) {
            return sanitizeText_(user && (user.username || user.name || user.email));
          }).filter(function(name) { return !!name; }).join(', '),
          task_url: sanitizeText_(task.url || task.permalink || task.link || task.html_url) ||
            ('https://app.clickup.com/t/' + String(task.id || '')),
          date_closed: task.date_closed ? fromMillisIso_(task.date_closed) : '',
          updated_at: task.date_updated ? fromMillisIso_(task.date_updated) : '',
          due_date: task.due_date ? fromMillisIso_(task.due_date) : ''
        };
        if (hasProjectClosingApprovalSignalByTaskOrPhase_(task, phaseMilestone.status_original, phase)) {
          phaseMilestone.status_original = 'APROVAR';
        }
        phaseMilestone.concluido = isClosedStatus_(phaseMilestone.status_original);
        normalizedMilestones.push(phaseMilestone);
        if (phaseMilestone.concluido) phase.marcos_concluidos += 1;
        else phase.marcos_pendentes += 1;
      }
      return;
    }
    var phaseInfo = resolvePhaseForTask_(task, byId, phaseMap);
    var phase = phaseInfo.phase;
    var countInSummary = shouldCountTaskInSummary_(task, phaseInfo, byId);
    var projectClosingTask = isProjectClosingDeliveryItem_(task, phase ? phase.nome : '', phase);
    var milestoneTask = isMilestoneTask_(task) || projectClosingTask;

    // Marcos podem ter subtarefas de evidencia/validacao. Eles ainda precisam
    // entrar no fechamento, mesmo quando nao contam como task folha no resumo.
    if (!countInSummary && !milestoneTask) {
      ignoredNestedItems += 1;
      return;
    }

    var item = {
      id: String(task.id),
      tipo: 'task',
      nome: sanitizeText_(task.name),
      fase_nome: phase ? phase.nome : '',
      parent_id: String(task.parent || ''),
      status_original: task.status && (task.status.status || task.status.type || task.status.label) || '',
      fase_status_original: phase && phase.status_original || '',
      custom_item_id: String(task.custom_item_id || ''),
      custom_item_name: clickUpTaskCustomItemName_(task),
      marcador_entrega: isProjectDeliveryTask_(task) ? 'sim' : '',
      responsaveis: (task.assignees || []).map(function(user) {
        return sanitizeText_(user && (user.username || user.name || user.email));
      }).filter(function(name) { return !!name; }).join(', '),
      task_url: sanitizeText_(task.url || task.permalink || task.link || task.html_url) ||
        ('https://app.clickup.com/t/' + String(task.id || '')),
      date_closed: task.date_closed ? fromMillisIso_(task.date_closed) : '',
      updated_at: task.date_updated ? fromMillisIso_(task.date_updated) : '',
      due_date: task.due_date ? fromMillisIso_(task.due_date) : ''
    };
    if (projectClosingTask && hasProjectClosingApprovalSignalByTaskOrPhase_(task, item.status_original, phase)) {
      item.status_original = 'APROVAR';
    }

    if (milestoneTask) {
      item.tipo = 'marco';
      item.concluido = isClosedStatus_(item.status_original);
      normalizedMilestones.push(item);
      if (phase) {
        if (item.concluido) phase.marcos_concluidos += 1;
        else phase.marcos_pendentes += 1;
      }
      return;
    }

    item.concluida = isClosedStatus_(item.status_original);
    normalizedTasks.push(item);
    if (phase) {
      if (item.concluida) phase.tasks_concluidas += 1;
      else phase.tasks_pendentes += 1;
    }
  });

  var phases = Object.keys(phaseMap).map(function(id) { return phaseMap[id]; });
  phases.sort(function(a, b) {
    if (a.ordem !== b.ordem) return a.ordem - b.ordem;
    return a.nome.localeCompare(b.nome);
  });

  var tasksDone = normalizedTasks.filter(function(item) { return item.concluida; }).length;
  var tasksPending = normalizedTasks.length - tasksDone;
  var milestonesDone = normalizedMilestones.filter(function(item) { return item.concluido; }).length;
  var milestonesPending = normalizedMilestones.length - milestonesDone;
  var totalItems = normalizedTasks.length + normalizedMilestones.length;
  var progress = totalItems ? Math.round(((tasksDone + milestonesDone) / totalItems) * 100) : 0;
  var latestUpdateItem = getLatestUpdateItemFromRawTasks_(tasks, byId, phaseMap);
  var latestUpdate = latestUpdateItem && latestUpdateItem.updated_at ? latestUpdateItem.updated_at : getLatestUpdate_(tasks);
  var projectUrl = mapping.project_url || buildProjectUrl_(mapping, tasks);
  var consultorInferido = mapping.consultor || inferConsultorFromTasks_(tasks);

  return {
    project_key: mapping.project_key,
    cliente: mapping.cliente,
    mes: mapping.mes,
    consultor: consultorInferido,
    project_url: projectUrl,
    view_id: mapping.view_id,
    list_id: mapping.list_id,
    synced_at: new Date().toISOString(),
    fases: phases.map(function(phase) {
      var tasksTotal = phase.tasks_concluidas + phase.tasks_pendentes;
      var marcosTotal = phase.marcos_concluidos + phase.marcos_pendentes;
      var totalItens = tasksTotal + marcosTotal;
      var totalConcluidos = phase.tasks_concluidas + phase.marcos_concluidos;
      return {
        tipo: 'fase',
        id: phase.id,
        nome: phase.nome,
        ordem: phase.ordem,
        status_original: phase.status_original || '',
        custom_item_id: phase.custom_item_id || '',
        custom_item_name: phase.custom_item_name || '',
        marcador_entrega: phase.marcador_entrega || '',
        marcador_marco: phase.marcador_marco || '',
        updated_at: phase.updated_at || '',
        task_url: phase.task_url || '',
        tasks_concluidas: phase.tasks_concluidas,
        tasks_pendentes: phase.tasks_pendentes,
        marcos_concluidos: phase.marcos_concluidos,
        marcos_pendentes: phase.marcos_pendentes,
        tasks_total: tasksTotal,
        marcos_total: marcosTotal,
        total_itens: totalItens,
        progresso: totalItens ? Math.round((totalConcluidos / totalItens) * 100) : 0
      };
    }),
    tasks: normalizedTasks,
    marcos: normalizedMilestones,
    resumo: {
      fases_total: phases.length,
      tasks_total: normalizedTasks.length,
      tasks_concluidas: tasksDone,
      tasks_pendentes: tasksPending,
      marcos_total: normalizedMilestones.length,
      marcos_concluidos: milestonesDone,
      marcos_pendentes: milestonesPending,
      progresso: progress,
      data_ultima_atualizacao: latestUpdate || '',
      ultima_alteracao_item: latestUpdateItem || null
    },
    clickup_payload: {
      source: payload.source,
      warning: payload.warning || '',
      fetched_task_count: tasks.length,
      ignored_nested_items: ignoredNestedItems,
      list_id: mapping.list_id || '',
      view_id: mapping.view_id || ''
    }
  };
}

function assertClickUpActivityDeadline_(options) {
  if (options && options.deadline_ms && new Date().getTime() >= Number(options.deadline_ms)) {
    throw new Error('Projeto excedeu o tempo seguro de leitura e foi marcado para nova tentativa.');
  }
}

function fetchProjectTasks_(mapping, options) {
  options = options || {};
  var listError = null;
  if (mapping.list_id) {
    try {
      return {
        source: 'list',
        tasks: fetchAllListTasks_(mapping.list_id, options)
      };
    } catch (error) {
      listError = error;
      if (!mapping.view_id || !isClickUpRecoverableSyncError_(error)) {
        throw error;
      }
    }
  }
  if (mapping.view_id) {
    try {
      return {
        source: listError ? 'view_fallback' : 'view',
        tasks: fetchAllViewTasks_(mapping.view_id, options),
        warning: listError ? simplifyErrorMessage_(listError) : ''
      };
    } catch (viewError) {
      if (listError) {
        throw new Error(
          'List sync failed: ' + simplifyErrorMessage_(listError) +
          ' | View fallback failed: ' + simplifyErrorMessage_(viewError)
        );
      }
      throw viewError;
    }
  }
  if (mapping.folder_id) {
    return {
      source: 'folder',
      tasks: fetchAllFolderTasks_(mapping.folder_id, options)
    };
  }
  if (mapping.space_id) {
    return {
      source: 'space',
      tasks: fetchAllSpaceTasks_(mapping.space_id, options)
    };
  }
  throw new Error('Project mapping must have list_id, view_id, folder_id or space_id: ' + mapping.project_key);
}

function fetchAllListTasks_(listId, options) {
  listId = normalizeClickUpId_(listId);
  if (!listId) throw new Error('CLICKUP_CONFIG com list_id invalido ou vazio.');
  var all = [];
  var customItemTypes = fetchClickUpCustomItemTypes_();
  var queryVariants = [[], ['custom_items[]=1']];
  Object.keys(customItemTypes).forEach(function(id) {
    if (normalizeKey_(customItemTypes[id]) !== 'ENTREGA') return;
    var param = 'custom_items[]=' + encodeURIComponent(id);
    if (!queryVariants.some(function(parts) { return parts.indexOf(param) >= 0; })) queryVariants.push([param]);
  });
  queryVariants.forEach(function(extraParams) {
    var page = 0;
    while (true) {
      assertClickUpActivityDeadline_(options);
      var query = [
        'include_closed=true',
        'subtasks=true',
        'include_timl=true',
        'page=' + page
      ].concat(extraParams).join('&');
      var response = clickupRequest_('get', '/list/' + listId + '/task?' + query);
      var batch = response.tasks || [];
      all = all.concat(batch);
      if (batch.length < 100) break;
      page += 1;
    }
  });
  return enrichClickUpCustomItemNames_(dedupeTasks_(all), customItemTypes);
}

function startProjectClosingSyncBackground_() {
  var props = PropertiesService.getScriptProperties();
  var alreadyActive = props.getProperty('PROJECT_CLOSING_SYNC_ACTIVE') === '1';
  if (!alreadyActive) {
    var total = loadProjectClosingSyncMappings_().length;
    props.setProperty('PROJECT_CLOSING_SYNC_ACTIVE', '1');
    props.setProperty('PROJECT_CLOSING_SYNC_CURSOR', '0');
    props.setProperty('PROJECT_CLOSING_SYNC_TOTAL', String(total));
    props.setProperty('PROJECT_CLOSING_SYNC_PROCESSED', '0');
    props.setProperty('PROJECT_CLOSING_SYNC_ERRORS', '0');
    props.setProperty('PROJECT_CLOSING_SYNC_STARTED_AT', new Date().toISOString());
    props.deleteProperty('PROJECT_CLOSING_SYNC_COMPLETED_AT');
    props.deleteProperty('PROJECT_CLOSING_SYNC_ERROR');
  }
  scheduleProjectClosingSyncBackground_(1000);
  var status = getProjectClosingSyncBackgroundStatus_();
  status.already_active = alreadyActive;
  return status;
}

function continueProjectClosingSyncBackgroundTrigger() {
  continueProjectClosingSyncBackgroundStepWithLock_();
}

function advanceProjectClosingSyncBackgroundFromStatus_() {
  return continueProjectClosingSyncBackgroundStepWithLock_({ status_poll: true });
}

function continueProjectClosingSyncBackgroundStepWithLock_(options) {
  options = options || {};
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('PROJECT_CLOSING_SYNC_ACTIVE') !== '1') {
    clearProjectClosingSyncBackgroundTriggers_();
    return getProjectClosingSyncBackgroundStatus_();
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(options.status_poll ? 1000 : 5000)) {
    var busy = getProjectClosingSyncBackgroundStatus_();
    busy.busy = true;
    return busy;
  }
  try {
    return continueProjectClosingSyncBackgroundStep_();
  } finally {
    lock.releaseLock();
  }
}

function continueProjectClosingSyncBackgroundStep_() {
  var props = PropertiesService.getScriptProperties();
  var mappings = loadProjectClosingSyncMappings_();
  try {
    var cursor = Math.max(0, toInt_(props.getProperty('PROJECT_CLOSING_SYNC_CURSOR'), 0));
    var batchSize = Math.max(1, Math.min(toInt_(getScriptProperty_('PROJECT_CLOSING_SYNC_BATCH_SIZE', '2'), 2), 3));
    var batch = mappings.slice(cursor, cursor + batchSize);
    var processed = toInt_(props.getProperty('PROJECT_CLOSING_SYNC_PROCESSED'), 0);
    var errors = toInt_(props.getProperty('PROJECT_CLOSING_SYNC_ERRORS'), 0);
    batch.forEach(function(mapping) {
      try {
        syncProjectMapping_(mapping, {
          force: true,
          skip_milestone_closing: true,
          deadline_ms: new Date().getTime() + Math.max(30000, Math.min(
            toInt_(getScriptProperty_('PROJECT_CLOSING_SYNC_PROJECT_TIMEOUT_MS', '45000'), 45000),
            90000
          ))
        });
        processed += 1;
      } catch (error) {
        errors += 1;
        writeSyncStatus_(mapping, 'error', error.message);
        props.setProperty('PROJECT_CLOSING_SYNC_ERROR', simplifyErrorMessage_(error));
      }
    });
    cursor += batch.length;
    props.setProperty('PROJECT_CLOSING_SYNC_CURSOR', String(cursor));
    props.setProperty('PROJECT_CLOSING_SYNC_PROCESSED', String(processed));
    props.setProperty('PROJECT_CLOSING_SYNC_ERRORS', String(errors));
    props.setProperty('PROJECT_CLOSING_SYNC_TOTAL', String(mappings.length));
    props.setProperty('PROJECT_CLOSING_SYNC_UPDATED_AT', new Date().toISOString());
    if (!batch.length || cursor >= mappings.length) {
      props.setProperty('PROJECT_CLOSING_SYNC_ACTIVE', '0');
      props.setProperty('PROJECT_CLOSING_SYNC_COMPLETED_AT', new Date().toISOString());
      clearProjectClosingSyncBackgroundTriggers_();
      return getProjectClosingSyncBackgroundStatus_();
    }
    scheduleProjectClosingSyncBackground_(3000);
    return getProjectClosingSyncBackgroundStatus_();
  } catch (error) {
    props.setProperty('PROJECT_CLOSING_SYNC_ERROR', simplifyErrorMessage_(error));
    props.setProperty('PROJECT_CLOSING_SYNC_ACTIVE', '0');
    clearProjectClosingSyncBackgroundTriggers_();
    return getProjectClosingSyncBackgroundStatus_();
  }
}

function loadProjectClosingSyncMappings_() {
  var out = [];
  var seen = {};

  function sourceKey(mapping) {
    if (!mapping) return '';
    if (mapping.list_id) return 'list|' + mapping.list_id;
    if (mapping.view_id) return 'view|' + mapping.view_id;
    if (mapping.folder_id) return 'folder|' + mapping.folder_id;
    if (mapping.space_id) return 'space|' + mapping.space_id + '|' + normalizeKey_(mapping.cliente);
    return normalizeProjectKey_(mapping.project_key || buildProjectKey_(mapping.mes, mapping.cliente));
  }

  function add(mapping) {
    if (!mapping || !mapping.enabled) return;
    if (!(mapping.list_id || mapping.view_id || mapping.folder_id || mapping.space_id || mapping.project_url)) return;
    var key = sourceKey(mapping);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(mapping);
  }

  loadProjectSyncMappings_().forEach(add);
  loadClickUpMilestoneClosingMappings_().forEach(add);
  return out;
}

function getProjectClosingSyncBackgroundStatus_() {
  var props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    active: props.getProperty('PROJECT_CLOSING_SYNC_ACTIVE') === '1',
    processed: toInt_(props.getProperty('PROJECT_CLOSING_SYNC_PROCESSED'), 0),
    total: toInt_(props.getProperty('PROJECT_CLOSING_SYNC_TOTAL'), 0),
    errors: toInt_(props.getProperty('PROJECT_CLOSING_SYNC_ERRORS'), 0),
    started_at: props.getProperty('PROJECT_CLOSING_SYNC_STARTED_AT') || '',
    updated_at: props.getProperty('PROJECT_CLOSING_SYNC_UPDATED_AT') || '',
    completed_at: props.getProperty('PROJECT_CLOSING_SYNC_COMPLETED_AT') || '',
    error: props.getProperty('PROJECT_CLOSING_SYNC_ERROR') || ''
  };
}

function stopProjectClosingSyncBackground_() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PROJECT_CLOSING_SYNC_ACTIVE', '0');
  props.setProperty('PROJECT_CLOSING_SYNC_UPDATED_AT', new Date().toISOString());
  clearProjectClosingSyncBackgroundTriggers_();
  var status = getProjectClosingSyncBackgroundStatus_();
  status.stopped = true;
  return status;
}

function scheduleProjectClosingSyncBackground_(delayMs) {
  clearProjectClosingSyncBackgroundTriggers_();
  ScriptApp.newTrigger('continueProjectClosingSyncBackgroundTrigger')
    .timeBased()
    .after(Math.max(1000, Number(delayMs || 3000)))
    .create();
}

function clearProjectClosingSyncBackgroundTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'continueProjectClosingSyncBackgroundTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function fetchAllViewTasks_(viewId, options) {
  viewId = normalizeClickUpId_(viewId);
  if (!viewId) throw new Error('CLICKUP_CONFIG com view_id invalido ou vazio.');
  var page = 0;
  var all = [];
  while (true) {
    assertClickUpActivityDeadline_(options);
    var response = clickupRequest_('get', '/view/' + viewId + '/task?page=' + page);
    var batch = response.tasks || [];
    all = all.concat(batch);
    if (batch.length < 100) break;
    page += 1;
  }
  return dedupeTasks_(all);
}

function fetchClickUpMilestoneCoverageTasks_() {
  var workspaceId = getClickUpWorkspaceId_();
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado para varredura geral de marcos.');
  var statuses = [
    'Closed', 'closed',
    'Aprovado Gestao', 'Aprovado Gestão', 'aprovado gestão',
    'Reprovado Gestao', 'Reprovado Gestão', 'reprovado gestão'
  ];
  var all = [];
  var successfulQueries = 0;
  statuses.forEach(function(status) {
    try {
      var page = 0;
      while (true) {
        var query = [
          'include_closed=true',
          'subtasks=true',
          'custom_items[]=1',
          'page=' + page,
          'statuses[]=' + encodeURIComponent(status)
        ].join('&');
        var response = clickupRequest_('get', '/team/' + workspaceId + '/task?' + query);
        var batch = response.tasks || [];
        batch.forEach(function(task) { task._confirmed_milestone = true; });
        all = all.concat(batch.filter(isMilestoneTask_));
        if (batch.length < 100) break;
        page += 1;
      }
      successfulQueries += 1;
    } catch (error) {
      // Status customizados podem ter grafias diferentes entre os spaces.
    }
  });
  statuses.forEach(function(status) {
    try {
      var page = 0;
      while (true) {
        var query = [
          'include_closed=true',
          'subtasks=true',
          'page=' + page,
          'statuses[]=' + encodeURIComponent(status)
        ].join('&');
        var response = clickupRequest_('get', '/team/' + workspaceId + '/task?' + query);
        var batch = response.tasks || [];
        all = all.concat(batch.filter(isProjectDeliveryTask_));
        if (batch.length < 100) break;
        page += 1;
      }
    } catch (error) {
      // Mantem a varredura de marcos mesmo se a consulta ampla falhar.
    }
  });
  if (!successfulQueries) throw new Error('Nenhum status de marco pôde ser consultado na varredura geral.');
  return dedupeTasks_(all);
}

function fetchClickUpRecentMilestoneCoverageTasks_(sinceMillis, options) {
  options = options || {};
  var workspaceId = getClickUpWorkspaceId_();
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado para atualizacao recente de marcos.');
  var since = Math.max(0, Number(sinceMillis || 0));
  var page = 0;
  var maxPages = Math.max(1, Number(options.max_pages || 10));
  var all = [];
  while (page < maxPages) {
    var query = [
      'include_closed=true',
      'subtasks=true',
      'custom_items[]=1',
      'date_updated_gt=' + since,
      'order_by=updated',
      'reverse=true',
      'page=' + page
    ].join('&');
    var response = clickupRequest_('get', '/team/' + workspaceId + '/task?' + query);
    var batch = response.tasks || [];
    batch.forEach(function(task) { task._confirmed_milestone = true; });
    // Também precisamos receber marcos que saíram do fluxo para remover
    // aprovações/reprovações antigas da base de fechamento.
    all = all.concat(batch.filter(isMilestoneTask_));
    all = all.concat(batch.filter(isProjectDeliveryTask_));
    if (batch.length < 100) break;
    page += 1;
  }
  page = 0;
  while (page < maxPages) {
    var deliveryQuery = [
      'include_closed=true',
      'subtasks=true',
      'date_updated_gt=' + since,
      'order_by=updated',
      'reverse=true',
      'page=' + page
    ].join('&');
    var deliveryResponse = clickupRequest_('get', '/team/' + workspaceId + '/task?' + deliveryQuery);
    var deliveryBatch = deliveryResponse.tasks || [];
    all = all.concat(deliveryBatch.filter(isProjectDeliveryTask_));
    if (deliveryBatch.length < 100) break;
    page += 1;
  }
  return dedupeTasks_(all);
}

function fetchClickUpTasksByIds_(taskIds) {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  var seen = {};
  var ids = (taskIds || []).map(normalizeClickUpId_).filter(function(id) {
    if (!id || seen[id]) return false;
    seen[id] = true;
    return true;
  }).slice(0, 150);
  var tasks = [];
  for (var offset = 0; offset < ids.length; offset += 50) {
    var batch = ids.slice(offset, offset + 50);
    var responses = UrlFetchApp.fetchAll(batch.map(function(id) {
      return {
        url: CLICKUP_API_BASE + '/task/' + id,
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          Authorization: token,
          Accept: 'application/json',
          'Content-Type': 'application/json'
        }
      };
    }));
    responses.forEach(function(response) {
      if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return;
      try {
        var task = JSON.parse(response.getContentText() || '{}');
        task._confirmed_milestone = true;
        tasks.push(task);
      } catch (error) {}
    });
  }
  return tasks;
}

function fetchClickUpValidationMilestoneTasks_() {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  var workspaceId = getClickUpWorkspaceId_();
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado.');
  var statuses = [
    'Aprovado Gestao', 'Aprovado Gestão', 'aprovado gestão',
    'Reprovado Gestao', 'Reprovado Gestão', 'reprovado gestão'
  ];
  var responses = UrlFetchApp.fetchAll(statuses.map(function(status) {
    return {
      url: CLICKUP_API_BASE + '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'custom_items[]=1',
        'page=0',
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'),
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Authorization: token,
        Accept: 'application/json',
        'Content-Type': 'application/json'
      }
    };
  }));
  var tasks = [];
  responses.forEach(function(response) {
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return;
    try {
      (JSON.parse(response.getContentText() || '{}').tasks || []).forEach(function(task) {
        task._confirmed_milestone = true;
        tasks.push(task);
      });
    } catch (error) {}
  });
  statuses.forEach(function(status) {
    try {
      var page = 0;
      while (page < 20) {
        var query = [
          'include_closed=true',
          'subtasks=true',
          'page=' + page,
          'statuses[]=' + encodeURIComponent(status)
        ].join('&');
        var rawBatch = clickupRequest_('get', '/team/' + workspaceId + '/task?' + query).tasks || [];
        tasks = tasks.concat(rawBatch.filter(isProjectDeliveryTask_));
        if (rawBatch.length < 100) break;
        page += 1;
      }
    } catch (error) {}
  });
  return dedupeTasks_(tasks);
}

function fetchClickUpValidationAndCurrentTasks_(currentValidationIds) {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  var workspaceId = getClickUpWorkspaceId_();
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado.');
  var statuses = [
    'Aprovado Gestao', 'Aprovado Gestão', 'aprovado gestão',
    'Reprovado Gestao', 'Reprovado Gestão', 'reprovado gestão'
  ];
  var requests = statuses.map(function(status) {
    return {
      url: CLICKUP_API_BASE + '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'custom_items[]=1',
        'page=0',
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'),
      kind: 'status'
    };
  });
  statuses.forEach(function(status) {
    requests.push({
      url: CLICKUP_API_BASE + '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'page=0',
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'),
      kind: 'status_delivery'
    });
  });
  (currentValidationIds || []).map(normalizeClickUpId_).filter(function(id, index, all) {
    return !!id && all.indexOf(id) === index;
  }).forEach(function(id) {
    requests.push({ url: CLICKUP_API_BASE + '/task/' + id, kind: 'task' });
  });
  var responses = UrlFetchApp.fetchAll(requests.map(function(request) {
    return {
      url: request.url,
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Authorization: token,
        Accept: 'application/json',
        'Content-Type': 'application/json'
      }
    };
  }));
  var tasks = [];
  responses.forEach(function(response, index) {
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return;
    try {
      var payload = JSON.parse(response.getContentText() || '{}');
      var found = requests[index].kind === 'status' || requests[index].kind === 'status_delivery' ? (payload.tasks || []) : [payload];
      if (requests[index].kind === 'status_delivery') found = found.filter(isProjectDeliveryTask_);
      found.forEach(function(task) {
        if (requests[index].kind !== 'status_delivery') task._confirmed_milestone = true;
        tasks.push(task);
      });
      var rawFoundLength = (requests[index].kind === 'status' || requests[index].kind === 'status_delivery')
        ? ((payload.tasks || []).length)
        : found.length;
      if ((requests[index].kind === 'status' || requests[index].kind === 'status_delivery') && rawFoundLength === 100) {
        var page = 1;
        while (page < 20) {
          var next = clickupRequestAbsolute_('get', requests[index].url.replace('page=0', 'page=' + page));
          var rawNextTasks = next.tasks || [];
          var nextTasks = requests[index].kind === 'status_delivery' ? rawNextTasks.filter(isProjectDeliveryTask_) : rawNextTasks;
          nextTasks.forEach(function(task) {
            if (requests[index].kind !== 'status_delivery') task._confirmed_milestone = true;
            tasks.push(task);
          });
          if (rawNextTasks.length < 100) break;
          page += 1;
        }
      }
    } catch (error) {}
  });
  return dedupeTasks_(tasks);
}

function clickUpMilestoneStatusAliases_(situation) {
  if (situation === 'aprovado') {
    return ['APROVADO GESTÃO', 'Aprovado Gestão', 'Aprovado Gestao', 'aprovado gestão'];
  }
  if (situation === 'reprovado') {
    return ['REPROVADO GESTÃO', 'Reprovado Gestão', 'Reprovado Gestao', 'reprovado gestão'];
  }
  return ['Closed', 'closed'];
}

function clickUpTaskStatusText_(task) {
  return sanitizeText_(task && task.status && (task.status.status || task.status.type || task.status.label));
}

function clickUpMilestoneStatusMatchesSituation_(status, situation) {
  var key = normalizeKey_(status);
  if (situation === 'aprovado') return key === 'APROVADO GESTAO';
  if (situation === 'reprovado') return key === 'REPROVADO GESTAO';
  if (situation === 'aguardando') return clickUpMilestoneSituation_(status) === 'aguardando';
  return clickUpMilestoneSituation_(status) === situation;
}

function fetchClickUpMilestonesBySituation_(situation, currentIds) {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  var workspaceId = getClickUpWorkspaceId_();
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado.');
  var requests = clickUpMilestoneStatusAliases_(situation).map(function(status) {
    return {
      url: CLICKUP_API_BASE + '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'custom_items[]=1',
        'page=0',
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'),
      kind: 'status'
    };
  });
  clickUpMilestoneStatusAliases_(situation).forEach(function(status) {
    requests.push({
      url: CLICKUP_API_BASE + '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'page=0',
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'),
      kind: 'status_delivery'
    });
  });
  (currentIds || []).map(normalizeClickUpId_).filter(function(id, index, all) {
    return !!id && all.indexOf(id) === index;
  }).forEach(function(id) {
    requests.push({ url: CLICKUP_API_BASE + '/task/' + id, kind: 'task' });
  });
  var responses = UrlFetchApp.fetchAll(requests.map(function(request) {
    return {
      url: request.url,
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Authorization: token,
        Accept: 'application/json',
        'Content-Type': 'application/json'
      }
    };
  }));
  var tasks = [];
  responses.forEach(function(response, index) {
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return;
    try {
      var payload = JSON.parse(response.getContentText() || '{}');
      var found = requests[index].kind === 'status' || requests[index].kind === 'status_delivery' ? (payload.tasks || []) : [payload];
      if (requests[index].kind === 'status_delivery') found = found.filter(isProjectDeliveryTask_);
      found = found.filter(function(task) {
        return clickUpMilestoneStatusMatchesSituation_(clickUpTaskStatusText_(task), situation);
      });
      found.forEach(function(task) {
        if (requests[index].kind !== 'status_delivery') task._confirmed_milestone = true;
        tasks.push(task);
      });
      var rawFoundLength = (requests[index].kind === 'status' || requests[index].kind === 'status_delivery')
        ? ((payload.tasks || []).length)
        : found.length;
      if ((requests[index].kind === 'status' || requests[index].kind === 'status_delivery') && rawFoundLength === 100) {
        var page = 1;
        while (page < 20) {
          var next = clickupRequestAbsolute_('get', requests[index].url.replace('page=0', 'page=' + page));
          var rawNextTasks = next.tasks || [];
          var nextTasks = requests[index].kind === 'status_delivery' ? rawNextTasks.filter(isProjectDeliveryTask_) : rawNextTasks;
          nextTasks = nextTasks.filter(function(task) {
            return clickUpMilestoneStatusMatchesSituation_(clickUpTaskStatusText_(task), situation);
          });
          nextTasks.forEach(function(task) {
            if (requests[index].kind !== 'status_delivery') task._confirmed_milestone = true;
            tasks.push(task);
          });
          if (rawNextTasks.length < 100) break;
          page += 1;
        }
      }
    } catch (error) {}
  });
  return dedupeTasks_(tasks);
}

function fetchAllFolderTasks_(folderId, options) {
  folderId = normalizeClickUpId_(folderId);
  if (!folderId) throw new Error('CLICKUP_CONFIG com folder_id invalido ou vazio.');
  var response = clickupRequest_('get', '/folder/' + folderId + '/list?archived=false');
  var lists = response.lists || [];
  var all = [];
  lists.forEach(function(list) {
    assertClickUpActivityDeadline_(options);
    if (list && list.id) {
      all = all.concat(fetchAllListTasks_(list.id, options));
    }
  });
  return dedupeTasks_(all);
}

function fetchAllSpaceTasks_(spaceId, options) {
  spaceId = normalizeClickUpId_(spaceId);
  if (!spaceId) throw new Error('CLICKUP_CONFIG com space_id invalido ou vazio.');
  var all = [];
  var folderResponse = clickupRequest_('get', '/space/' + spaceId + '/folder?archived=false');
  var folders = folderResponse.folders || [];
  folders.forEach(function(folder) {
    assertClickUpActivityDeadline_(options);
    if (folder && folder.id) {
      all = all.concat(fetchAllFolderTasks_(folder.id, options));
    }
  });
  var listResponse = clickupRequest_('get', '/space/' + spaceId + '/list?archived=false');
  var lists = listResponse.lists || [];
  lists.forEach(function(list) {
    assertClickUpActivityDeadline_(options);
    if (list && list.id) {
      all = all.concat(fetchAllListTasks_(list.id, options));
    }
  });
  return dedupeTasks_(all);
}

function inserirSpacesHistoricosClickUp() {
  var result = importClickUpSpaceListsToConfig_(HISTORICAL_CLICKUP_SPACES);
  try {
    SpreadsheetApp.getUi().alert(
      'Importacao ClickUp concluida',
      'Projetos/pastas historicos inseridos: ' + result.inserted +
        '\nLinhas historicas antigas removidas: ' + result.removed +
        '\nJa existiam no CLICKUP_CONFIG: ' + result.skipped +
        '\nSpaces lidos: ' + result.spaces +
        '\n\nDepois rode "Sincronizar inventario ClickUp".',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {}
  return result;
}

function importClickUpSpaceListsToConfig_(spaces) {
  var sheet = getConfigSheet_();
  var headers = ['enabled', 'mes', 'cliente', 'project_key', 'project_url', 'view_id', 'list_id', 'folder_id', 'space_id', 'sync_mode', 'notes'];
  ensureHeaders_(sheet, headers);

  var values = sheet.getDataRange().getValues();
  var header = normalizeConfigHeader_(values[0] || headers);
  var keptRows = [];
  var removed = 0;
  var existingListIds = {};
  var existingFolderIds = {};
  values.slice(1).forEach(function(row) {
    var item = rowToObject_(header, row);
    if (isAutoHistoricalConfigRow_(item)) {
      removed += 1;
      return;
    }
    keptRows.push(item);
    var listId = normalizeClickUpId_(item.list_id) || extractClickUpIdFromUrl_(item.project_url, 'list');
    var folderId = normalizeClickUpId_(item.folder_id) || extractClickUpIdFromUrl_(item.project_url, 'folder');
    if (listId) existingListIds[listId] = true;
    if (folderId) existingFolderIds[folderId] = true;
  });

  var rows = [];
  var skipped = 0;
  (spaces || []).forEach(function(space) {
    var projects = fetchClickUpImplementationProjectsFromSpace_(space);
    projects.forEach(function(project) {
      var listId = normalizeClickUpId_(project.list_id);
      var folderId = normalizeClickUpId_(project.folder_id);
      if (!listId && !folderId) return;
      if ((folderId && existingFolderIds[folderId]) || (listId && existingListIds[listId])) {
        skipped += 1;
        return;
      }
      if (listId) existingListIds[listId] = true;
      if (folderId) existingFolderIds[folderId] = true;
      rows.push({
        enabled: true,
        mes: 'HIST',
        cliente: sanitizeText_(project.name),
        project_key: buildHistoricalProjectKey_(space.name, project),
        project_url: listId ? buildClickUpListUrl_(listId) : buildClickUpFolderUrl_(folderId),
        view_id: '',
        list_id: listId,
        folder_id: folderId,
        space_id: normalizeClickUpId_(space.space_id),
        sync_mode: 'list',
        notes: 'Importado automaticamente de ' + sanitizeText_(space.name) + ' / Cronograma de Implantacao'
      });
    });
  });

  sheet.clearContents();
  ensureHeaders_(sheet, headers);
  var outputObjects = keptRows.concat(rows);
  if (outputObjects.length) {
    var output = outputObjects.map(function(item) {
      return headers.map(function(headerName) { return item[headerName] !== undefined ? item[headerName] : ''; });
    });
    sheet.getRange(2, 1, output.length, headers.length).setValues(output);
  }

  return {
    ok: true,
    spaces: (spaces || []).length,
    inserted: rows.length,
    skipped: skipped,
    removed: removed,
    sheet: sheet.getName()
  };
}

function isAutoHistoricalConfigRow_(item) {
  var mes = sanitizeMonth_(item && item.mes);
  var notes = normalizeKey_(item && item.notes);
  var projectKey = normalizeKey_(item && item.project_key);
  return mes === 'HIST' && (
    notes.indexOf('IMPORTADO AUTOMATICAMENTE') >= 0 ||
    projectKey.indexOf('HIST CSI PROJETOS') === 0 ||
    projectKey.indexOf('HIST CSI AVULSOS') === 0
  );
}

function fetchClickUpImplementationProjectsFromSpace_(space) {
  var spaceId = normalizeClickUpId_(space && space.space_id);
  if (!spaceId) throw new Error('Space historico sem space_id valido: ' + sanitizeText_(space && space.name));
  var all = [];
  var seen = {};

  function addProjectFromList(list, folder) {
    var listId = normalizeClickUpId_(list && list.id);
    if (!listId || seen['list:' + listId]) return;
    if (!isImplementationScheduleList_(list && list.name)) return;
    seen['list:' + listId] = true;
    all.push({
      type: 'list',
      id: listId,
      list_id: listId,
      folder_id: normalizeClickUpId_(folder && folder.id),
      name: sanitizeText_((folder && folder.name) || list.name),
      space_id: spaceId,
      space_name: sanitizeText_(space.name)
    });
  }

  var rootResponse = clickupRequest_('get', '/space/' + spaceId + '/list?archived=false');
  (rootResponse.lists || []).forEach(function(list) {
    addProjectFromList(list, null);
  });

  var folderResponse = clickupRequest_('get', '/space/' + spaceId + '/folder?archived=false');
  (folderResponse.folders || []).forEach(function(folder) {
    var folderLists = folder.lists || [];
    if (!folderLists.length && folder && folder.id) {
      var listResponse = clickupRequest_('get', '/folder/' + folder.id + '/list?archived=false');
      folderLists = listResponse.lists || [];
    }
    folderLists.forEach(function(list) {
      addProjectFromList(list, folder);
    });
  });

  return all;
}

function isImplementationScheduleList_(name) {
  var key = normalizeKey_(name);
  return key === 'CRONOGRAMA DE IMPLANTACAO' ||
    key.indexOf('CRONOGRAMA DE IMPLANTACAO') >= 0;
}

function buildHistoricalProjectKey_(spaceName, project) {
  return 'HIST|' + normalizeKey_(spaceName) + '|' + normalizeKey_(project.type || 'project') + '|' + normalizeClickUpId_(project.id);
}

function buildClickUpListUrl_(listId) {
  return 'https://app.clickup.com/9007083069/v/li/' + normalizeClickUpId_(listId);
}

function buildClickUpFolderUrl_(folderId) {
  return 'https://app.clickup.com/9007083069/v/f/' + normalizeClickUpId_(folderId);
}

function sincronizarInventarioClickUp() {
  var result = syncClickUpInventoryFromConfig_();
  try {
    SpreadsheetApp.getUi().alert(
      'Inventario ClickUp sincronizado',
      'Projetos/listas historicos gravados: ' + result.total +
        '\nErros: ' + result.errors.length +
        '\nAba: ' + result.sheet,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {}
  return result;
}

function syncClickUpInventoryFromConfig_() {
  var mappings = loadProjectMappings_().filter(function(item) {
    return item.enabled &&
      sanitizeMonth_(item.mes) === 'HIST' &&
      (item.list_id || item.view_id || item.folder_id || item.space_id);
  });
  var rows = [];
  var errors = [];
  mappings.forEach(function(mapping) {
    try {
      var normalized = buildNormalizedProjectFromClickUp_(mapping);
      rows.push(inventoryRowFromNormalized_(mapping, normalized, 'ok', ''));
    } catch (error) {
      errors.push({ project_key: mapping.project_key, cliente: mapping.cliente, error: simplifyErrorMessage_(error) });
      rows.push(inventoryRowFromNormalized_(mapping, null, 'error', simplifyErrorMessage_(error)));
    }
  });

  var sheet = getClickUpInventorySheet_();
  var headers = getClickUpInventoryHeaders_();
  writeObjectsToSheet_(sheet, rows, headers);
  return { ok: errors.length === 0, total: rows.length, errors: errors, sheet: sheet.getName() };
}

function inventoryRowFromNormalized_(mapping, normalized, status, errorMessage) {
  var resumo = normalized && normalized.resumo || {};
  return {
    mes: mapping.mes || '',
    cliente: mapping.cliente || '',
    consultor: normalized && normalized.consultor || mapping.consultor || '',
    status: status === 'ok' ? inferProjectStatusFromSummary_(resumo) : '',
    project_key: mapping.project_key || '',
    project_url: normalized && normalized.project_url || mapping.project_url || '',
    view_id: mapping.view_id || '',
    list_id: mapping.list_id || '',
    folder_id: mapping.folder_id || '',
    space_id: mapping.space_id || '',
    tasks_concluidas: resumo.tasks_concluidas || 0,
    tasks_pendentes: resumo.tasks_pendentes || 0,
    marcos_concluidos: resumo.marcos_concluidos || 0,
    marcos_pendentes: resumo.marcos_pendentes || 0,
    fases_total: resumo.fases_total || 0,
    progresso: resumo.progresso || 0,
    data_ultima_atualizacao: resumo.data_ultima_atualizacao || '',
    dias_sem_atualizacao: diffDaysFromIso_(resumo.data_ultima_atualizacao),
    clickup_json: normalized ? buildSheetSafeClickUpJson_(normalized) : '',
    ultima_sync_clickup: new Date(),
    sync_status_clickup: status || '',
    sync_error_clickup: errorMessage || ''
  };
}

function getClickUpInventory_(params) {
  var user = requireUser_(params || {});
  var sheet = getClickUpInventorySheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, projetos: [], total: 0 };
  var header = values[0];
  var projetos = values.slice(1).map(function(row) {
    var item = rowToObject_(header, row);
    if (item.ultima_sync_clickup instanceof Date) item.ultima_sync_clickup = item.ultima_sync_clickup.toISOString();
    return item;
  }).filter(function(item) {
    return !!sanitizeText_(item.cliente) && canUserAccessProjectItem_(user, item);
  });
  return { ok: true, projetos: projetos, total: projetos.length };
}

function getClickUpMilestoneClosing_(params) {
  requireUser_(params || {});
  var sheet = getClickUpMilestoneClosingSheet_();
  var values = sheet.getDataRange().getDisplayValues();
  var rows = values.length > 1 ? values.slice(1).map(function(row) {
    var item = rowToObject_(values[0], row);
    item.mes_fechamento = normalizeClickUpMonthReference_(item.mes_fechamento, item.closed_at);
    item.mes_validacao = normalizeClickUpMonthReference_(item.mes_validacao, item.validation_at);
    item.consultor = clickUpMilestoneConsultant_(item.consultor || item.responsaveis);
    return item;
  }).filter(function(item) {
    return !!sanitizeText_(item.task_id) && /^(MARCO|FECHAMENTO DE PROJETO|ENTREGA DE PROJETO)$/.test(normalizeKey_(item.item_tipo));
  }) : [];
  rows = rows.filter(function(item) {
    return String(item.mes_fechamento || '').slice(0, 7) >= '2024-01';
  });
  var month = sanitizeText_((params || {}).month || (params || {}).mes).slice(0, 7);
  if (month) {
    rows = rows.filter(function(item) {
      return String(item.mes_fechamento || '').slice(0, 7) === month;
    });
  }
  rows.sort(function(a, b) {
    return String(b.closed_at || b.validation_at || b.updated_at || '').localeCompare(
      String(a.closed_at || a.validation_at || a.updated_at || '')
    );
  });
  return {
    ok: true,
    marcos: rows,
    total: rows.length,
    bonus_por_marco: CLICKUP_MILESTONE_BONUS_VALUE,
    background_sync: getClickUpMilestoneClosingBackgroundStatus_()
  };
}

function diagnoseProjectClosing_(params) {
  requireUser_(params || {});
  var raw = sanitizeText_(params.list_id || params.query || params.link || params.url);
  var ids = raw.match(/\d{6,}/g) || [];
  var listId = normalizeClickUpId_(params.list_id || ids[ids.length - 1] || raw);
  if (!listId) throw new Error('Informe um list_id ou link do ClickUp.');
  var tasks = fetchAllListTasks_(listId, { deadline: new Date().getTime() + 25000 });
  var phaseMap = {};
  var byId = {};
  tasks.forEach(function(task) { byId[String(task.id)] = task; });
  tasks.forEach(function(task) {
    if (isPhaseTask_(task, byId)) {
      phaseMap[String(task.id)] = {
        tipo: 'fase',
        id: String(task.id),
        nome: sanitizeText_(task.name),
        ordem: extractLeadingNumber_(task.name),
        status_original: task.status && (task.status.status || task.status.type || task.status.label) || '',
        custom_item_name: clickUpTaskCustomItemName_(task)
      };
    }
  });
  var breakOffItems = [];
  tasks.forEach(function(task) {
    var phaseTask = isPhaseTask_(task, byId);
    var phaseInfo = phaseTask
      ? { phase: phaseMap[String(task.id)] || null }
      : resolvePhaseForTask_(task, byId, phaseMap);
    var phaseName = phaseInfo.phase ? phaseInfo.phase.nome : '';
    if (!isProjectBreakOffText_(phaseName) && !isProjectBreakOffText_(task && task.name)) return;
    var status = sanitizeText_(task.status && (task.status.status || task.status.type || task.status.label));
    var marker = clickUpTaskCustomItemName_(task);
    var phaseStatus = sanitizeText_(phaseInfo.phase && phaseInfo.phase.status_original);
    var approvalSignal = hasProjectClosingApprovalSignalByTaskOrPhase_(task, status, phaseInfo.phase);
    breakOffItems.push({
      id: String(task.id || ''),
      nome: sanitizeText_(task.name),
      fase_nome: phaseName,
      status_original: status,
      fase_status_original: phaseStatus,
      custom_item_id: String(task.custom_item_id || ''),
      custom_item_name: marker,
      item_tipo_clickup: marker,
      marcador_marco: isMilestoneTask_(task) ? 'sim' : '',
      aprovar: approvalSignal,
      entra_regra: isProjectClosingDeliveryItem_(task, phaseName, phaseInfo.phase),
      sinal_aprovacao_bruto: approvalSignal && !isProjectClosingApprovalStatus_(status) ? 'sim' : '',
      debug_tipo: summarizeClickUpProjectClosingDebug_(task),
      link: sanitizeText_(task.url || task.permalink || task.link || task.html_url) || ('https://app.clickup.com/t/' + String(task.id || ''))
    });
  });
  return {
    ok: true,
    project_closing_rule_version: CLICKUP_PROJECT_CLOSING_RULE_VERSION,
    list_id: listId,
    tasks_total: tasks.length,
    breakoff_total: breakOffItems.length,
    aprovados_total: breakOffItems.filter(function(item) { return item.entra_regra; }).length,
    items: breakOffItems
  };
}

function summarizeClickUpProjectClosingDebug_(task) {
  var parts = [];
  function push(label, value) {
    var text = sanitizeText_(value);
    if (text) parts.push(label + ': ' + text);
  }
  push('custom_item_id', task && task.custom_item_id);
  push('custom_item', task && task.custom_item && (task.custom_item.name || task.custom_item.label || task.custom_item.type || JSON.stringify(task.custom_item)));
  push('custom_item_type', task && task.custom_item_type && (task.custom_item_type.name || task.custom_item_type.label || JSON.stringify(task.custom_item_type)));
  push('custom_type', task && task.custom_type && (task.custom_type.name || task.custom_type.label || JSON.stringify(task.custom_type)));
  push('custom_task_type', task && task.custom_task_type && (task.custom_task_type.name || task.custom_task_type.label || JSON.stringify(task.custom_task_type)));
  if (task && Array.isArray(task.tags) && task.tags.length) {
    push('tags', task.tags.map(function(tag) { return sanitizeText_(tag && (tag.name || tag.tag || tag.label)); }).filter(Boolean).join(', '));
  }
  if (task && Array.isArray(task.custom_fields) && task.custom_fields.length) {
    var fields = task.custom_fields.map(function(field) {
      var name = sanitizeText_(field && field.name);
      var value = sanitizeText_(field && (field.value || field.display_value || field.text || field.type_config && JSON.stringify(field.type_config)));
      return name && value ? name + '=' + value : '';
    }).filter(Boolean).slice(0, 6).join(' | ');
    push('campos', fields);
  }
  var result = parts.join(' · ');
  return result.length > 600 ? result.slice(0, 600) + '...' : result;
}

function clickUpMilestoneClosingValidatedRecovery_(taskId) {
  var recoveries = {
    '86aguzfdc': { validation_at: '2026-06-30T13:19:00.000Z', justificativa: 'Validado no fechamento enviado em JUN/2026; preservado contra retorno para Closed.' },
    '86aguzf0z': { validation_at: '2026-06-30T13:20:00.000Z', justificativa: 'Validado no fechamento enviado em JUN/2026; preservado contra retorno para Closed.' },
    '86ag9bqbt': { validation_at: '2026-06-30T20:31:00.000Z', justificativa: 'Validado no fechamento enviado em JUN/2026; preservado contra retorno para Closed.' }
  };
  var recovery = recoveries[String(taskId || '').trim()];
  if (!recovery) return null;
  return Object.assign({
    status_atual: 'aprovado gestão',
    situacao: 'aprovado',
    mes_fechamento: '2026-06',
    mes_validacao: '2026-06',
    valor_bonus: CLICKUP_MILESTONE_BONUS_VALUE
  }, recovery);
}

function upsertClickUpMilestoneClosing_(mapping, normalized, options) {
  options = options || {};
  if (options.preserve_closed_history === undefined) options.preserve_closed_history = true;
  var sheet = getClickUpMilestoneClosingSheet_();
  var headers = getClickUpMilestoneClosingHeaders_();
  var current = options.current || loadClickUpMilestoneClosingCurrent_(sheet);
  var now = new Date().toISOString();
  (normalized.marcos || []).forEach(function(milestone) {
    var taskId = String(milestone.id || '');
    if (!taskId) return;
    var previous = current[taskId] || null;
    var status = sanitizeText_(milestone.status_original);
    var hasClosedDate = !!sanitizeText_(milestone.date_closed);
    var situation = clickUpMilestoneSituation_(status);
    var hasDeliveryMarker = isProjectClosingMilestone_(milestone);
    var isProjectClosing = hasDeliveryMarker && isProjectClosingApprovalStatus_(status);
    if (isProjectClosing) situation = 'aprovado';
    var validatedRecovery = clickUpMilestoneClosingValidatedRecovery_(taskId);
    if (validatedRecovery) {
      situation = validatedRecovery.situacao;
      status = validatedRecovery.status_atual;
    }
    var previousSituation = sanitizeText_(previous && previous.situacao);
    var preserveValidatedHistory = previous &&
      (previousSituation === 'aprovado' || previousSituation === 'reprovado') &&
      situation === 'aguardando' &&
      options.preserve_closed_history !== false &&
      options.allow_validation_downgrade !== true;
    if (preserveValidatedHistory) {
      situation = previousSituation;
      status = sanitizeText_(previous && previous.status_atual) || status;
    }
    var returnedToClosed = !preserveValidatedHistory && ((situation === 'aguardando' &&
      (previousSituation === 'aprovado' || previousSituation === 'reprovado')) ||
      (situation === 'outro' && hasClosedDate && !previous));
    var recoveredClosedHistory = false;
    if (situation === 'outro' && hasClosedDate && options.preserve_closed_history) {
      situation = 'aguardando';
      recoveredClosedHistory = true;
    }
    if (!previous && situation === 'outro') return;
    if (previous && situation === 'outro' && options.authoritative) return;
    var rowStatus = recoveredClosedHistory ? 'Closed historico' : status;
    var statusChanged = !previous || sanitizeText_(previous.status_atual) !== rowStatus;
    var closedAt = (returnedToClosed ? sanitizeText_(milestone.updated_at) : '') ||
      sanitizeText_(previous && previous.closed_at) ||
      sanitizeText_(milestone.date_closed) ||
      (situation !== 'outro' ? sanitizeText_(milestone.updated_at) || now : '');
    var validationAt = situation === 'aprovado' || situation === 'reprovado'
      ? sanitizeText_(previous && previous.validation_at)
      : '';
    if ((situation === 'aprovado' || situation === 'reprovado') && (!validationAt || statusChanged)) {
      validationAt = sanitizeText_(milestone.updated_at) || now;
    }
    var isValidation = situation === 'aprovado' || situation === 'reprovado';
    var clearComment = options.fetch_comments !== false &&
      options.authoritative &&
      (!options.validation_comments_only || isValidation);
    var justification = (clearComment || returnedToClosed) ? '' : sanitizeText_(previous && previous.justificativa);
    var justificationBy = (clearComment || returnedToClosed) ? '' : sanitizeText_(previous && previous.justificativa_por);
    var shouldFetchComment = options.fetch_comments !== false &&
      situation !== 'outro' &&
      (!options.validation_comments_only || isValidation) &&
      (options.authoritative || statusChanged || !justification);
    if (shouldFetchComment) {
      var comment = fetchLatestClickUpTaskComment_(taskId);
      justification = comment.text || '';
      justificationBy = comment.user || '';
    }
    var history = parseJsonArray_(previous && previous.status_history_json);
    if (statusChanged) {
      history.push({ status: status, situacao: situation, at: sanitizeText_(milestone.updated_at) || now });
    }
    var milestoneConsultant = clickUpMilestoneConsultant_(milestone.responsaveis);
    var portfolioConsultant = clickUpMilestoneConsultant_(normalized.consultor || mapping.consultor);
    if (!milestoneConsultant || /ADMINISTRATIVO|ADMINISTRADOR|MULTSOFT|SUPORTE/.test(normalizeKey_(milestoneConsultant))) {
      milestoneConsultant = portfolioConsultant || milestoneConsultant;
    }
    var row = {
      task_id: taskId,
      item_tipo: 'Marco',
      project_key: mapping.project_key || '',
      projeto: mapping.cliente || normalized.cliente || '',
      consultor: milestoneConsultant || 'Sem consultor',
      marco: milestone.nome || '',
      fase: milestone.fase_nome || '',
      status_atual: rowStatus,
      situacao: situation,
      closed_at: closedAt,
      mes_fechamento: clickUpMonthReference_(closedAt),
      validation_at: validationAt,
      mes_validacao: clickUpMonthReference_(validationAt),
      justificativa: justification,
      justificativa_por: justificationBy,
      valor_bonus: situation === 'aprovado'
        ? (isProjectClosing && String(validationAt || '').slice(0, 10) >= CLICKUP_PROJECT_CLOSING_BONUS_START
          ? CLICKUP_PROJECT_CLOSING_BONUS_VALUE
          : (hasDeliveryMarker ? 0 : CLICKUP_MILESTONE_BONUS_VALUE))
        : 0,
      link: milestone.task_url || ('https://app.clickup.com/t/' + taskId),
      responsaveis: milestone.responsaveis || '',
      updated_at: milestone.updated_at || now,
      sincronizado_em: now,
      status_history_json: JSON.stringify(history),
      marcador_entrega: hasDeliveryMarker ? 'sim' : '',
      fechamento_projeto: isProjectClosing ? 'sim' : ''
    };
    if (validatedRecovery) {
      Object.keys(validatedRecovery).forEach(function(key) {
        if (key === 'justificativa' && row.justificativa) return;
        row[key] = validatedRecovery[key];
      });
      row.justificativa_por = row.justificativa_por || 'Fechamento JUN/2026';
    }
    current[taskId] = row;
  });
  if (!options.defer_write) writeClickUpMilestoneClosingCurrent_(sheet, headers, current);
  return current;
}

function loadClickUpMilestoneClosingCurrent_(sheet) {
  ensureHeaders_(sheet, getClickUpMilestoneClosingHeaders_());
  var values = sheet.getDataRange().getValues();
  var current = {};
  if (values.length > 1) {
    values.slice(1).forEach(function(row) {
      var item = rowToObject_(values[0], row);
      if (item.task_id) current[String(item.task_id)] = item;
    });
  }
  return current;
}

function writeClickUpMilestoneClosingCurrent_(sheet, headers, current) {
  headers = headers && headers.length ? headers : getClickUpMilestoneClosingHeaders_();
  writeObjectsToSheet_(sheet, Object.keys(current).map(function(id) { return current[id]; }), headers);
  sheet.setFrozenRows(1);
}

function upsertClickUpMilestoneClosingEntries_(entries, options) {
  options = options || {};
  if (!entries || !entries.length) return;
  var sheet = getClickUpMilestoneClosingSheet_();
  var headers = getClickUpMilestoneClosingHeaders_();
  var current = loadClickUpMilestoneClosingCurrent_(sheet);
  entries.forEach(function(entry) {
    upsertClickUpMilestoneClosing_(entry.mapping, entry.normalized, {
      current: current,
      defer_write: true,
      fetch_comments: options.fetch_comments !== false,
      authoritative: options.authoritative === true,
      validation_comments_only: options.validation_comments_only === true,
      preserve_closed_history: options.preserve_closed_history !== false
    });
  });
  writeClickUpMilestoneClosingCurrent_(sheet, headers, current);
}

function clickUpMilestoneSituation_(status) {
  var key = normalizeKey_(status);
  if (/REPROVAD.*GESTAO/.test(key)) return 'reprovado';
  if (/APROVAD.*GESTAO/.test(key)) return 'aprovado';
  if (/CLOSED|CLOSE|CONCLUID|FINALIZ|DONE|COMPLETE/.test(key)) return 'aguardando';
  return 'outro';
}

function clickUpMonthReference_(value) {
  if (!value) return '';
  var date = new Date(value);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

function normalizeClickUpMonthReference_(value, fallbackDate) {
  var text = sanitizeText_(value);
  var isoMatch = text.match(/^(\d{4})-(\d{2})/);
  if (isoMatch) return isoMatch[1] + '-' + isoMatch[2];
  var brMatch = text.match(/^(\d{2})\/(\d{4})$/);
  if (brMatch) return brMatch[2] + '-' + brMatch[1];
  return clickUpMonthReference_(fallbackDate || value);
}

function clickUpMilestoneConsultant_(value) {
  var seen = {};
  var names = sanitizeText_(value).split(/\s*,\s*/).filter(function(name) {
    var key = normalizeKey_(name);
    if (!key || /ADMINISTRATIVO|ADMINISTRADOR|MULTSOFT|SUPORTE/.test(key) || seen[key]) return false;
    seen[key] = true;
    return true;
  });
  return names[0] || sanitizeText_(value).split(/\s*,\s*/)[0] || 'Sem consultor';
}

function parseJsonArray_(value) {
  if (Array.isArray(value)) return value;
  try {
    var parsed = JSON.parse(String(value || '[]'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function parseJsonObject_(value) {
  if (value && typeof value === 'object' && !Array.isArray(value)) return value;
  try {
    var parsed = JSON.parse(String(value || '{}'));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
  } catch (error) {
    return null;
  }
}

function parseLatestClickUpTaskComment_(response) {
    var comments = response && response.comments || [];
    if (!comments.length) return { text: '', user: '' };
    comments.sort(function(a, b) {
      return Number(b.date || b.date_created || 0) - Number(a.date || a.date_created || 0);
    });
    var meaningful = comments.map(function(comment) {
      var text = Array.isArray(comment.comment_text)
      ? comment.comment_text.map(function(part) {
        return sanitizeText_(part && (part.text || part.value || part.content));
      }).filter(function(part) { return !!part; }).join(' ')
      : sanitizeText_(comment.comment_text || comment.comment || comment.text);
      var attachments = (comment.attachments || []).map(function(attachment) {
        return sanitizeText_(attachment && (attachment.title || attachment.name || attachment.filename));
      }).filter(function(name) { return !!name; });
      if (attachments.length) text += (text ? ' | ' : '') + 'Anexo(s): ' + attachments.join(', ');
      return {
        text: sanitizeText_(text),
        user: sanitizeText_(comment.user && (comment.user.username || comment.user.name || comment.user.email))
      };
    }).filter(function(comment) { return !!comment.text; });
    if (!meaningful.length) return { text: '', user: '' };
    return {
      text: meaningful.slice(0, 5).map(function(comment) { return comment.text; }).join(' | '),
      user: meaningful[0].user
    };
}

function isProjectClosingMilestone_(milestone) {
  return isProjectClosingDeliveryItem_(milestone, milestone && milestone.fase_nome);
}

function isProjectClosingApprovalStatus_(status) {
  return normalizeKey_(status) === 'APROVAR';
}

function getProjectClosingDecisionSheet_() {
  var sheet = getOrCreateSheet_('CLICKUP_PROJECT_CLOSING_DECISIONS');
  ensureHeaders_(sheet, [
    'project_key', 'project_name', 'consultant', 'item_id', 'item_name',
    'decision', 'notes', 'decided_at', 'decided_by', 'month', 'bonus_value',
    'clickup_status', 'clickup_updated_at'
  ]);
  return sheet;
}

function projectClosingDecisionDateText_(value) {
  if (value instanceof Date) return value.toISOString();
  return sanitizeText_(value);
}

function getProjectClosingDecisions_(params) {
  requireUser_(params || {});
  var sheet = getProjectClosingDecisionSheet_();
  var values = sheet.getDataRange().getValues();
  var header = values[0] || [];
  var items = values.slice(1).map(function(row) {
    var item = rowToObject_(header, row);
    return {
      project_key: sanitizeText_(item.project_key),
      project_name: sanitizeText_(item.project_name),
      consultant: sanitizeText_(item.consultant),
      item_id: sanitizeText_(item.item_id),
      item_name: sanitizeText_(item.item_name),
      decision: sanitizeText_(item.decision).toLowerCase(),
      notes: sanitizeText_(item.notes),
      decided_at: projectClosingDecisionDateText_(item.decided_at),
      decided_by: sanitizeText_(item.decided_by),
      month: sanitizeText_(item.month),
      bonus_value: Number(item.bonus_value || 0),
      clickup_status: sanitizeText_(item.clickup_status),
      clickup_updated_at: projectClosingDecisionDateText_(item.clickup_updated_at)
    };
  }).filter(function(item) { return item.project_key && item.decision; });
  return { ok: true, items: items };
}

function projectClosingCandidateFromTask_(task, mappings) {
  var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
  var normalized = buildNormalizedMilestoneCoverageProject_(mapping, task);
  var milestone = normalized.marcos && normalized.marcos[0] || {};
  var projectKey = sanitizeText_(mapping.project_key || '') ||
    normalizeProjectKey_(buildProjectKey_(mapping.mes || '', mapping.cliente || '')) ||
    ('CLICKUP|' + sanitizeText_(milestone.id || task && task.id));
  var projectName = sanitizeText_(mapping.cliente || normalized.cliente) ||
    sanitizeText_(task && task.folder && task.folder.name) ||
    sanitizeText_(task && task.project && task.project.name) ||
    sanitizeText_(task && task.list && task.list.name) ||
    'Projeto nao mapeado';
  var consultant = sanitizeText_(mapping.consultor || normalized.consultor || milestone.responsaveis);
  return {
    project_key: projectKey,
    project_name: projectName,
    consultant: consultant,
    item_id: sanitizeText_(milestone.id || task && task.id),
    item_name: sanitizeText_(milestone.nome || task && task.name || 'Fase 8 - Break Off'),
    item_status: sanitizeText_(milestone.status_original || task && task.status && (task.status.status || task.status.label)),
    phase_name: sanitizeText_(milestone.fase_nome || task && task._project_closing_phase_name),
    phase_status: sanitizeText_(milestone.fase_status_original || task && task._project_closing_phase_status),
    item_url: sanitizeText_(milestone.task_url || task && (task.url || task.permalink || task.link || task.html_url)) ||
      ('https://app.clickup.com/t/' + sanitizeText_(milestone.id || task && task.id)),
    list_id: sanitizeText_(mapping.list_id || task && task.list && task.list.id),
    folder_id: sanitizeText_(mapping.folder_id || task && task.folder && task.folder.id),
    space_id: sanitizeText_(mapping.space_id || task && task.space && task.space.id),
    marcador_entrega: 'sim',
    updated_at: sanitizeText_(milestone.updated_at) || (task && task.date_updated ? fromMillisIso_(task.date_updated) : ''),
    project_closing_rule_version: CLICKUP_PROJECT_CLOSING_RULE_VERSION
  };
}

function projectClosingCandidateSourceKey_(mapping) {
  if (!mapping) return '';
  if (mapping.list_id) return 'list|' + mapping.list_id;
  if (mapping.view_id) return 'view|' + mapping.view_id;
  if (mapping.folder_id) return 'folder|' + mapping.folder_id;
  if (mapping.space_id) return 'space|' + mapping.space_id + '|' + normalizeKey_(mapping.cliente);
  return normalizeProjectKey_(mapping.project_key || buildProjectKey_(mapping.mes, mapping.cliente));
}

function loadProjectClosingCandidateMappings_() {
  var seen = {};
  var out = [];
  loadProjectSyncMappings_().forEach(function(mapping) {
    if (!mapping || !mapping.enabled) return;
    if (!(mapping.list_id || mapping.view_id || mapping.folder_id || mapping.space_id || mapping.project_url)) return;
    var key = projectClosingCandidateSourceKey_(mapping);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(mapping);
  });
  return out;
}

function projectClosingCandidateFromNormalizedItem_(mapping, normalized, item) {
  var projectKey = sanitizeText_(normalized && normalized.project_key || mapping && mapping.project_key || '') ||
    normalizeProjectKey_(buildProjectKey_(mapping && mapping.mes || '', mapping && mapping.cliente || '')) ||
    ('CLICKUP|' + sanitizeText_(item && item.id));
  var projectName = sanitizeText_(normalized && normalized.cliente || mapping && mapping.cliente) ||
    'Projeto ClickUp';
  var consultant = sanitizeText_(normalized && normalized.consultor || mapping && mapping.consultor || item && item.responsaveis);
  var itemId = sanitizeText_(item && item.id);
  return {
    project_key: projectKey,
    project_name: projectName,
    consultant: consultant,
    item_id: itemId,
    item_name: sanitizeText_(item && item.nome) || 'Fase 8 - Break Off',
    item_status: sanitizeText_(item && item.status_original) || 'APROVAR',
    phase_name: sanitizeText_(item && item.fase_nome) || sanitizeText_(item && item.nome),
    phase_status: sanitizeText_(item && item.fase_status_original),
    item_url: sanitizeText_(item && item.task_url) || (itemId ? ('https://app.clickup.com/t/' + itemId) : ''),
    list_id: sanitizeText_(mapping && mapping.list_id || normalized && normalized.list_id),
    folder_id: sanitizeText_(mapping && mapping.folder_id),
    space_id: sanitizeText_(mapping && mapping.space_id),
    marcador_entrega: 'sim',
    updated_at: sanitizeText_(item && item.updated_at),
    project_closing_rule_version: CLICKUP_PROJECT_CLOSING_RULE_VERSION
  };
}

function projectClosingCandidatesFromMapping_(mapping, options) {
  var normalized = buildNormalizedProjectFromClickUp_(mapping, options || {});
  var seen = {};
  var source = []
    .concat(normalized && normalized.marcos || [])
    .concat(normalized && normalized.tasks || []);
  return source.filter(function(item) {
    if (!item || !item.id || seen[item.id]) return false;
    if (!isProjectClosingDeliveryItem_(item, item.fase_nome || item.nome, {
      status_original: item.fase_status_original || item.status_original,
      marcador_entrega: item.marcador_entrega,
      custom_item_id: item.custom_item_id,
      custom_item_name: item.custom_item_name
    })) return false;
    seen[item.id] = true;
    return true;
  }).map(function(item) {
    return projectClosingCandidateFromNormalizedItem_(mapping, normalized, item);
  });
}

function projectClosingSavedBreakOffCandidates_() {
  var seen = {};
  var out = [];
  MONTHS.forEach(function(month) {
    try {
      getMonthlyProjectsFromSheet_(month).forEach(function(project) {
        var payload = parseJsonObject_(project && project.clickup_json);
        if (!payload) return;
        var projectKey = sanitizeText_(payload.project_key) ||
          buildProjectKey_(month, project.cliente) + '|ROW|' + String(project._sheet_row || '');
        var source = []
          .concat(Array.isArray(payload.fases) ? payload.fases : [])
          .concat(Array.isArray(payload.marcos) ? payload.marcos : [])
          .concat(Array.isArray(payload.tasks) ? payload.tasks : []);
        source.forEach(function(item) {
          var itemId = sanitizeText_(item && (item.id || item.task_id));
          var key = itemId || projectKey;
          if (!key || seen[key]) return;
          var isBreakOff = isProjectBreakOffText_(item && (item.fase_nome || item.fase || item.phase)) ||
            isProjectBreakOffText_(item && (item.nome || item.marco || item.name));
          if (!isBreakOff || !isProjectDeliveryClosingCandidate_(item)) return;
          seen[key] = true;
          out.push({
            project_key: projectKey,
            project_name: sanitizeText_(payload.cliente || project.cliente) || 'Projeto ClickUp',
            consultant: sanitizeText_(payload.consultor || project.consultor),
            item_id: itemId,
            item_name: sanitizeText_(item && (item.nome || item.marco || item.name)) || 'Fase 8 - Break Off',
            item_status: sanitizeText_(item && (item.status_original || item.status || item.status_atual)),
            phase_name: sanitizeText_(item && (item.fase_nome || item.fase || item.phase || item.nome || item.name)),
            phase_status: sanitizeText_(item && item.fase_status_original),
            item_url: sanitizeText_(item && (item.task_url || item.url || item.link)) ||
              sanitizeText_(payload.project_url || project.projeto_link || project.link_projeto),
            list_id: sanitizeText_(payload.list_id || project.list_id),
            folder_id: sanitizeText_(payload.folder_id),
            space_id: sanitizeText_(payload.space_id),
            marcador_entrega: 'sim',
            updated_at: sanitizeText_(item && item.updated_at),
            project_closing_rule_version: CLICKUP_PROJECT_CLOSING_RULE_VERSION
          });
        });
      });
    } catch (error) {}
  });
  return out;
}

function refreshProjectClosingCandidatesByTaskId_(candidates) {
  var byId = {};
  fetchClickUpTasksByIds_((candidates || []).map(function(item) {
    return item && item.item_id;
  })).forEach(function(task) {
    byId[normalizeClickUpId_(task && task.id)] = task;
  });
  return (candidates || []).map(function(candidate) {
    var task = byId[normalizeClickUpId_(candidate && candidate.item_id)];
    if (!task) return candidate;
    var status = clickUpTaskStatusText_(task);
    var itemName = sanitizeText_(task.name) || candidate.item_name;
    return Object.assign({}, candidate, {
      item_name: itemName,
      item_status: status || candidate.item_status,
      item_url: sanitizeText_(task.url || task.permalink || task.link || task.html_url) || candidate.item_url,
      updated_at: task.date_updated ? fromMillisIso_(task.date_updated) : candidate.updated_at,
      marcador_entrega: isProjectDeliveryTask_(task) ? 'sim' : candidate.marcador_entrega
    });
  });
}

function getProjectClosingCandidates_(params) {
  requireUser_(params || {});
  var started = new Date();
  var savedCandidates = projectClosingSavedBreakOffCandidates_();
  var offset = Math.max(0, toInt_((params || {}).offset, 0));
  var limit = Math.max(1, Math.min(toInt_((params || {}).limit, 50), 100));
  var batch = savedCandidates.slice(offset, offset + limit);
  var seen = {};
  var errors = [];
  var refreshed = [];
  try {
    refreshed = refreshProjectClosingCandidatesByTaskId_(batch);
  } catch (error) {
    errors.push({ project_name: 'Atualizacao direta por ID', error: simplifyErrorMessage_(error) });
    refreshed = batch;
  }
  var items = refreshed.filter(function(item) {
    if (!isProjectClosingApprovalStatus_(item && item.item_status) &&
        !isProjectClosingApprovalStatus_(item && item.phase_status)) return false;
    var key = sanitizeText_(item && item.item_id) || sanitizeText_(item && item.project_key);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
  var nextOffset = Math.min(savedCandidates.length, offset + batch.length);
  return {
    ok: true,
    source: 'clickup_project_closing_saved_ids',
    project_closing_rule_version: CLICKUP_PROJECT_CLOSING_RULE_VERSION,
    offset: offset,
    limit: limit,
    processed: nextOffset,
    total: savedCandidates.length,
    next_offset: nextOffset,
    has_more: nextOffset < savedCandidates.length,
    done: nextOffset >= savedCandidates.length,
    scanned: batch.length,
    errors: errors.length,
    error_details: errors.slice(0, 20),
    items: items,
    generated_at: started.toISOString()
  };
}

function setProjectClosingDecision_(params) {
  var admin = requireAdmin_(params || {});
  var projectKey = sanitizeText_(params.project_key);
  var decision = sanitizeText_(params.decision).toLowerCase();
  var notes = sanitizeText_(params.notes || params.observacao);
  if (!projectKey) throw new Error('Projeto obrigatorio.');
  if (['approved', 'rejected'].indexOf(decision) < 0) throw new Error('Decisao invalida.');
  if (decision === 'rejected' && !notes) throw new Error('Informe o motivo da reprovacao.');

  var itemId = sanitizeText_(params.item_id);
  if (!itemId) throw new Error('Item Fase 8 nao encontrado para atualizar no ClickUp. Sincronize os projetos e tente novamente.');
  var desiredStatus = decision === 'approved' ? 'APROVADO GESTAO' : 'REPROVADO GESTAO';
  var clickupUpdate = updateProjectClosingStatusInClickUp_(itemId, desiredStatus, notes, admin);

  var now = new Date();
  var month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  var item = {
    project_key: projectKey,
    project_name: sanitizeText_(params.project_name),
    consultant: sanitizeText_(params.consultant),
    item_id: itemId,
    item_name: sanitizeText_(params.item_name),
    decision: decision,
    notes: notes,
    decided_at: now,
    decided_by: admin.name || admin.username,
    month: month,
    bonus_value: decision === 'approved' ? CLICKUP_PROJECT_CLOSING_BONUS_VALUE : 0,
    clickup_status: clickupUpdate.status,
    clickup_updated_at: now
  };
  upsertProjectClosingDecisionRow_(item);
  item.decided_at = now.toISOString();
  item.clickup_updated_at = now.toISOString();
  item.clickup_warning = clickupUpdate.warning || '';
  return { ok: true, item: item };
}

function upsertProjectClosingDecisionRow_(item) {
  var sheet = getProjectClosingDecisionSheet_();
  var values = sheet.getDataRange().getValues();
  var header = values[0] || [];
  var keyIndex = header.indexOf('project_key');
  var targetRow = 0;
  for (var i = 1; i < values.length; i++) {
    if (sanitizeText_(values[i][keyIndex]) === sanitizeText_(item.project_key)) {
      targetRow = i + 1;
      break;
    }
  }
  var row = header.map(function(name) { return item[name] == null ? '' : item[name]; });
  if (targetRow) sheet.getRange(targetRow, 1, 1, header.length).setValues([row]);
  else sheet.appendRow(row);
}

function reconcileProjectClosingDecisionFromNormalized_(mapping, normalized) {
  var items = []
    .concat(normalized && normalized.fases || [])
    .concat(normalized && normalized.marcos || [])
    .concat(normalized && normalized.tasks || []);
  var finalItem = items.filter(function(item) {
    if (!isProjectBreakOffText_(item && (item.nome || item.fase_nome))) return false;
    if (!isProjectDeliveryClosingCandidate_(item)) return false;
    var statusKey = normalizeKey_(item && item.status_original);
    return statusKey === 'APROVADO GESTAO' || statusKey === 'REPROVADO GESTAO';
  })[0];
  if (!finalItem) return;
  var statusKey = normalizeKey_(finalItem.status_original);
  var approved = statusKey === 'APROVADO GESTAO';
  var decidedAtText = sanitizeText_(finalItem.updated_at) || new Date().toISOString();
  var decidedAt = new Date(decidedAtText);
  if (isNaN(decidedAt.getTime())) decidedAt = new Date();
  upsertProjectClosingDecisionRow_({
    project_key: mapping.project_key,
    project_name: normalized.cliente || mapping.cliente,
    consultant: normalized.consultor || mapping.consultor,
    item_id: finalItem.id || '',
    item_name: finalItem.nome || 'Fase 8 - Break Off',
    decision: approved ? 'approved' : 'rejected',
    notes: 'Situacao reconciliada pelo status atual do ClickUp.',
    decided_at: decidedAt,
    decided_by: 'ClickUp',
    month: Utilities.formatDate(decidedAt, Session.getScriptTimeZone(), 'yyyy-MM'),
    bonus_value: approved ? CLICKUP_PROJECT_CLOSING_BONUS_VALUE : 0,
    clickup_status: finalItem.status_original || '',
    clickup_updated_at: decidedAt
  });
}

function updateProjectClosingStatusInClickUp_(taskId, desiredStatus, notes, user) {
  var task = clickupRequest_('get', '/task/' + encodeURIComponent(taskId));
  var listId = sanitizeText_(task && task.list && task.list.id);
  if (!listId) throw new Error('Lista do item Fase 8 nao encontrada no ClickUp.');
  var list = clickupRequest_('get', '/list/' + encodeURIComponent(listId));
  var statuses = Array.isArray(list && list.statuses) ? list.statuses : [];
  var desiredKey = normalizeKey_(desiredStatus);
  var matched = statuses.filter(function(status) {
    return normalizeKey_(status && (status.status || status.name || status.label)) === desiredKey;
  })[0];
  if (!matched) {
    throw new Error('O status ' + desiredStatus + ' nao existe na lista ' + sanitizeText_(list && list.name || listId) + '. Crie esse status no ClickUp e tente novamente.');
  }
  var statusLabel = sanitizeText_(matched.status || matched.name || matched.label);
  var updated = clickupRequest_('put', '/task/' + encodeURIComponent(taskId), { status: statusLabel });
  var confirmedStatus = sanitizeText_(updated && updated.status && (updated.status.status || updated.status.label));
  if (confirmedStatus && normalizeKey_(confirmedStatus) !== normalizeKey_(statusLabel)) {
    throw new Error('O ClickUp respondeu com status diferente do solicitado: ' + confirmedStatus + '.');
  }
  var warning = '';
  if (notes) {
    try {
      clickupRequest_('post', '/task/' + encodeURIComponent(taskId) + '/comment', {
        comment_text: notes + '\n\nDecisao registrada por ' + sanitizeText_(user && (user.name || user.username)),
        notify_all: false
      });
    } catch (commentError) {
      warning = 'Status atualizado, mas o comentario nao foi enviado ao ClickUp.';
    }
  }
  return { status: statusLabel, warning: warning };
}

function hasProjectClosingApprovalSignal_(task, status) {
  if (isProjectClosingApprovalStatus_(status)) return true;
  var taskStatus = sanitizeText_(task && task.status && (task.status.status || task.status.type || task.status.label));
  return isProjectClosingApprovalStatus_(taskStatus);
}

function hasProjectClosingApprovalSignalByTaskOrPhase_(task, status, phase) {
  if (hasProjectClosingApprovalSignal_(task, status)) return true;
  var phaseStatus = sanitizeText_(phase && phase.status_original);
  return isProjectClosingApprovalStatus_(phaseStatus);
}

function fetchLatestClickUpTaskComment_(taskId) {
  try {
    return parseLatestClickUpTaskComment_(
      clickupRequest_('get', '/task/' + normalizeClickUpId_(taskId) + '/comment')
    );
  } catch (error) {
    return { text: '', user: '' };
  }
}

function fetchLatestClickUpTaskCommentsByIds_(taskIds) {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  var ids = (taskIds || []).map(normalizeClickUpId_).filter(function(id, index, all) {
    return !!id && all.indexOf(id) === index;
  });
  var commentsById = {};
  if (!ids.length) return commentsById;
  var batchSize = 10;
  for (var offset = 0; offset < ids.length; offset += batchSize) {
    if (offset > 0) Utilities.sleep(1100);
    var batch = ids.slice(offset, offset + batchSize);
    var responses = UrlFetchApp.fetchAll(batch.map(function(id) {
      return {
        url: CLICKUP_API_BASE + '/task/' + id + '/comment',
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          Authorization: token,
          Accept: 'application/json',
          'Content-Type': 'application/json'
        }
      };
    }));
    responses.forEach(function(response, index) {
      var id = batch[index];
      if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
        commentsById[id] = { text: '', user: '' };
        return;
      }
      try {
        commentsById[id] = parseLatestClickUpTaskComment_(JSON.parse(response.getContentText() || '{}'));
      } catch (error) {
        commentsById[id] = { text: '', user: '' };
      }
    });
  }
  return commentsById;
}

function startClickUpMilestoneClosingBackground_(params) {
  params = params || {};
  var props = PropertiesService.getScriptProperties();
  var forceHistory = String(params.force_history || params.force_full || '') === '1';
  var monthCount = clickUpMilestoneClosingDistinctMonths_().length;
  var fullProjectScan = monthCount <= 1 ||
    String(params.full_project_scan || params.force_project_scan || '') === '1';
  migrateClickUpMilestoneClosingSchema_();
  if (forceHistory) {
    props.deleteProperty('CLICKUP_MILESTONE_HISTORY_REBUILT');
    props.deleteProperty('CLICKUP_CLOSED_INCREMENTAL_SINCE');
  }
  if (props.getProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE') === '1') {
    // A second click must never restart an in-flight historical rebuild. Restarting
    // at "monthly" publishes the partial monthly cache again and loses progress.
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    scheduleClickUpMilestoneClosingBackground_(1000);
    return {
      ok: true,
      scheduled: true,
      already_active: true,
      background_sync: getClickUpMilestoneClosingBackgroundStatus_()
    };
  }
  props.setProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE', '1');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_OFFSET', '0');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_PROCESSED', '0');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', '0');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', '0');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_SKIP_PROJECT_SCAN', fullProjectScan ? '0' : '1');
  // Mesmo uma reconstrução completa precisa começar pelo incremental recente:
  // é nessa fase que capturamos retornos de Aprovado/Reprovado para Closed.
  props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'project_approvals');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_STARTED_AT', new Date().toISOString());
  props.deleteProperty('CLICKUP_MILESTONE_CLOSING_ERROR');
  if (fullProjectScan) {
    props.deleteProperty('CLICKUP_MILESTONE_HISTORY_REBUILT');
  }
  props.setProperty('CLICKUP_MILESTONE_CLOSING_TOTAL', fullProjectScan ? String(loadClickUpMilestoneClosingMappings_().length) : '0');
  props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
  scheduleClickUpMilestoneClosingBackground_(1000);
  return {
    ok: true,
    scheduled: true,
    background_sync: getClickUpMilestoneClosingBackgroundStatus_()
  };
}

function clickUpMilestoneClosingDistinctMonths_() {
  var sheet = getClickUpMilestoneClosingSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var monthColumn = getClickUpMilestoneClosingHeaders_().indexOf('mes_fechamento') + 1;
  var found = {};
  sheet.getRange(2, monthColumn, lastRow - 1, 1).getDisplayValues().forEach(function(row) {
    var month = normalizeClickUpMonthReference_(row[0]);
    if (month) found[month] = true;
  });
  return Object.keys(found);
}

function migrateClickUpMilestoneClosingSchema_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_MILESTONE_CLOSING_SCHEMA_VERSION') === CLICKUP_MILESTONE_CLOSING_SCHEMA_VERSION) return;
  var sheet = getClickUpMilestoneClosingSheet_();
  ensureHeaders_(sheet, getClickUpMilestoneClosingHeaders_());
  props.setProperty('CLICKUP_MILESTONE_CLOSING_SCHEMA_VERSION', CLICKUP_MILESTONE_CLOSING_SCHEMA_VERSION);
}

function sincronizarFechamentoMarcosClickUp() {
  var result = startClickUpMilestoneClosingBackground_({});
  try {
    SpreadsheetApp.getUi().alert(
      'Fechamento de marcos',
      'Atualização agendada em segundo plano. A guia CLICKUP_FECHAMENTO_MARCOS será atualizada por lotes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {}
  return result;
}

function restaurarFechamentoMarcosHistoricoMensal() {
  var result = restoreClickUpMilestoneClosingFromMonthlyHistory_({});
  try {
    SpreadsheetApp.getUi().alert(
      'Fechamento de marcos restaurado',
      'Aba CLICKUP_FECHAMENTO_MARCOS restaurada pelo histórico mensal.' +
        '\nAntes: ' + result.before_rows +
        '\nDepois: ' + result.after_rows +
        '\nMarcos lidos no histórico: ' + result.detected +
        '\nProjetos com histórico: ' + result.projects +
        '\nErros: ' + result.errors,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {}
  return result;
}

function restoreClickUpMilestoneClosingFromMonthlyHistory_(params) {
  params = params || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Outra restauração/sincronização de fechamento ainda está em andamento. Tente novamente em alguns segundos.');
  }
  try {
    migrateClickUpMilestoneClosingSchema_();
    var sheet = getClickUpMilestoneClosingSheet_();
    var beforeRows = Math.max(0, sheet.getLastRow() - 1);
    var result = syncClickUpMilestoneClosingFromMonthlySheets_();
    var afterRows = Math.max(0, sheet.getLastRow() - 1);
    var props = PropertiesService.getScriptProperties();
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE', '0');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'monthly_restored');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', String(result.detected || 0));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', String(result.errors || 0));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    props.setProperty('CLICKUP_MILESTONE_CLOSING_COMPLETED_AT', new Date().toISOString());
    if (result.last_error) props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', result.last_error);
    clearClickUpMilestoneClosingBackgroundTriggers_();
    return {
      ok: true,
      restored: true,
      before_rows: beforeRows,
      after_rows: afterRows,
      detected: result.detected || 0,
      projects: result.projects || 0,
      errors: result.errors || 0,
      last_error: result.last_error || '',
      source: 'monthly_clickup_json'
    };
  } finally {
    lock.releaseLock();
  }
}

function syncClickUpMilestoneClosingTrigger() {
  var status = getClickUpMilestoneClosingBackgroundStatus_();
  if (!status.active) startClickUpMilestoneClosingBackground_({});
}

function continueClickUpMilestoneClosingBackgroundTrigger() {
  var props = PropertiesService.getScriptProperties();
  try {
    continueClickUpMilestoneClosingBackgroundWorker_();
  } catch (error) {
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', simplifyErrorMessage_(error));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    if (props.getProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE') === '1') {
      scheduleClickUpMilestoneClosingBackground_(60000);
    }
  }
}

function continueClickUpMilestoneClosingBackgroundWorker_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE') !== '1') {
    clearClickUpMilestoneClosingBackgroundTriggers_();
    return;
  }
  var phase = props.getProperty('CLICKUP_MILESTONE_CLOSING_PHASE') || 'recent';
  if (phase === 'project_approvals') {
    var approvals = syncClickUpProjectClosingApprovalCoverage_();
    props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', String(approvals.detected || 0));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', String(approvals.errors || 0));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'recent');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    if (approvals.last_error) props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', approvals.last_error);
    scheduleClickUpMilestoneClosingBackground_(1000);
    return;
  }
  if (phase === 'recent') {
    var recent = syncClickUpRecentMilestoneCoverage_({
      authoritative: true,
      validation_comments_only: true,
      preserve_closed_history: true
    });
    props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', String(recent.detected || 0));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', String(recent.errors || 0));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'history');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    if (recent.last_error) props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', recent.last_error);
    scheduleClickUpMilestoneClosingBackground_(5000);
    return;
  }
  if (phase === 'history') {
    var history = ensureClickUpMilestoneHistoryCoverage_();
    props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', String(
      toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_DETECTED'), 0) + Number(history.detected || 0)
    ));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', String(
      toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_ERRORS'), 0) + Number(history.errors || 0)
    ));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'monthly');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    if (history.last_error) props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', history.last_error);
    scheduleClickUpMilestoneClosingBackground_(5000);
    return;
  }
  if (phase === 'monthly') {
    var monthly = syncClickUpMilestoneClosingFromMonthlySheets_();
    props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', String(
      toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_DETECTED'), 0) + Number(monthly.detected || 0)
    ));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', String(
      toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_ERRORS'), 0) + Number(monthly.errors || 0)
    ));
    props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
    if (monthly.last_error) props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', monthly.last_error);
    if (props.getProperty('CLICKUP_MILESTONE_CLOSING_SKIP_PROJECT_SCAN') === '1') {
      props.setProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE', '0');
      props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'complete_fast');
      props.setProperty('CLICKUP_MILESTONE_CLOSING_COMPLETED_AT', new Date().toISOString());
      clearClickUpMilestoneClosingBackgroundTriggers_();
      return;
    }
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'projects');
    scheduleClickUpMilestoneClosingBackground_(5000);
    return;
  }
  if (props.getProperty('CLICKUP_MILESTONE_CLOSING_SKIP_PROJECT_SCAN') !== '0') {
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE', '0');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'complete_fast');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_COMPLETED_AT', new Date().toISOString());
    clearClickUpMilestoneClosingBackgroundTriggers_();
    return;
  }
  var offset = toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_OFFSET'), 0);
  // A single large project can consume most of the Apps Script execution window.
  // Keep each continuation small so the offset is persisted instead of retrying
  // the same batch forever after a timeout.
  var result = syncClickUpMilestoneClosingBatch_(offset, 3);
  var processed = toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_PROCESSED'), 0) + result.processed;
  var errors = toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_ERRORS'), 0) + result.errors;
  var detected = toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_DETECTED'), 0) + result.detected;
  props.setProperty('CLICKUP_MILESTONE_CLOSING_OFFSET', String(result.next_offset));
  props.setProperty('CLICKUP_MILESTONE_CLOSING_PROCESSED', String(processed));
  props.setProperty('CLICKUP_MILESTONE_CLOSING_ERRORS', String(errors));
  props.setProperty('CLICKUP_MILESTONE_CLOSING_DETECTED', String(detected));
  props.setProperty('CLICKUP_MILESTONE_CLOSING_TOTAL', String(result.total));
  props.setProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT', new Date().toISOString());
  if (result.last_error) props.setProperty('CLICKUP_MILESTONE_CLOSING_ERROR', result.last_error);
  if (result.done) {
    props.setProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE', '0');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_PHASE', 'complete');
    props.setProperty('CLICKUP_MILESTONE_CLOSING_COMPLETED_AT', new Date().toISOString());
    clearClickUpMilestoneClosingBackgroundTriggers_();
    return;
  }
  scheduleClickUpMilestoneClosingBackground_(1000);
}

function syncClickUpMilestoneClosingBatch_(offset, limit) {
  var mappings = loadClickUpMilestoneClosingMappings_();
  offset = Math.max(0, Number(offset || 0));
  limit = Math.max(1, Number(limit || 5));
  var batch = mappings.slice(offset, offset + limit);
  var processed = 0;
  var errors = 0;
  var detected = 0;
  var coverageDetected = 0;
  var lastError = '';
  var pendingUpserts = [];
  batch.forEach(function(mapping) {
    try {
      var normalized = buildNormalizedProjectFromClickUp_(mapping);
      detected += (normalized.marcos || []).length;
      pendingUpserts.push({ mapping: mapping, normalized: normalized });
    } catch (error) {
      errors += 1;
      lastError = simplifyErrorMessage_(error);
    }
    processed += 1;
  });
  upsertClickUpMilestoneClosingEntries_(pendingUpserts);
  var reachedEnd = !batch.length || offset + batch.length >= mappings.length;
  if (reachedEnd) {
    var coverageTasks = [];
    try {
      coverageTasks = fetchClickUpMilestoneCoverageTasks_();
    } catch (coverageError) {
      errors += 1;
      lastError = 'Varredura geral: ' + simplifyErrorMessage_(coverageError);
    }
    getClickUpMilestoneAuditTaskIds_().forEach(function(taskId) {
      try {
        coverageTasks.push(clickupRequest_('get', '/task/' + normalizeClickUpId_(taskId)));
      } catch (directError) {
        errors += 1;
        lastError = 'Leitura direta do marco ' + taskId + ': ' + simplifyErrorMessage_(directError);
      }
    });
    coverageTasks = dedupeTasks_(coverageTasks);
    coverageDetected = coverageTasks.length;
    var coverageUpserts = coverageTasks.map(function(task) {
      var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
      var normalized = buildNormalizedMilestoneCoverageProject_(mapping, task);
      return { mapping: mapping, normalized: normalized };
    });
    upsertClickUpMilestoneClosingEntries_(coverageUpserts);
  }
  return {
    total: mappings.length,
    processed: processed,
    errors: errors,
    detected: detected + coverageDetected,
    coverage_detected: coverageDetected,
    last_error: lastError,
    next_offset: offset + batch.length,
    done: reachedEnd
  };
}

function clickUpProjectClosingApprovalStatusAliases_() {
  return ['APROVAR', 'Aprovar', 'aprovar'];
}

function fetchClickUpProjectClosingApprovalTasks_(options) {
  options = options || {};
  var workspaceId = getClickUpWorkspaceId_();
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado.');
  var maxPages = Math.max(1, Math.min(toInt_(options.max_pages, 4), 10));
  var tasks = [];
  clickUpProjectClosingApprovalStatusAliases_().forEach(function(status) {
    for (var page = 0; page < maxPages; page++) {
      var response = clickupRequest_('get', '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'custom_items[]=1',
        'page=' + page,
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'));
      var batch = response.tasks || [];
      tasks = tasks.concat(batch.filter(function(task) {
        return isProjectClosingDeliveryItem_(task, sanitizeText_(task && task.list && task.list.name), null);
      }));
      if (batch.length < 100) break;
      Utilities.sleep(250);
    }
  });
  return dedupeTasks_(tasks);
}

function syncClickUpProjectClosingApprovalCoverage_() {
  var mappings = loadClickUpMilestoneClosingMappings_();
  var errors = 0;
  var lastError = '';
  var tasks = [];
  try {
    tasks = fetchClickUpProjectClosingApprovalTasks_({ max_pages: 4 });
  } catch (error) {
    errors += 1;
    lastError = 'Aprovar fechamento projeto: ' + simplifyErrorMessage_(error);
  }
  var entries = tasks.map(function(task) {
    var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
    return { mapping: mapping, normalized: buildNormalizedMilestoneCoverageProject_(mapping, task) };
  });
  upsertClickUpMilestoneClosingEntries_(entries, {
    fetch_comments: false,
    authoritative: false,
    validation_comments_only: true,
    preserve_closed_history: true
  });
  return { detected: entries.length, errors: errors, last_error: lastError };
}

function syncClickUpMilestoneClosingFromMonthlySheets_() {
  var sheet = getClickUpMilestoneClosingSheet_();
  var headers = getClickUpMilestoneClosingHeaders_();
  var current = loadClickUpMilestoneClosingCurrent_(sheet);
  var detected = 0;
  var projects = 0;
  var errors = 0;
  var lastError = '';

  MONTHS.forEach(function(month) {
    try {
      getMonthlyProjectsFromSheet_(month).forEach(function(project) {
        var payload = parseJsonObject_(project.clickup_json);
        if (!payload || !Array.isArray(payload.marcos) || !payload.marcos.length) return;
        var normalized = payload;
        normalized.cliente = sanitizeText_(normalized.cliente) || project.cliente;
        normalized.consultor = sanitizeText_(normalized.consultor) || project.consultor;
        var mapping = {
          enabled: true,
          mes: month,
          cliente: project.cliente,
          project_key: sanitizeText_(payload.project_key) || buildProjectKey_(month, project.cliente),
          project_url: sanitizeText_(payload.project_url) || sanitizeText_(project.projeto_link || project.link_projeto),
          view_id: sanitizeText_(payload.view_id) || normalizeClickUpId_(project.view_id),
          list_id: sanitizeText_(payload.list_id) || normalizeClickUpId_(project.list_id),
          folder_id: sanitizeText_(payload.folder_id || ''),
          space_id: sanitizeText_(payload.space_id || ''),
          consultor: project.consultor
        };
        upsertClickUpMilestoneClosing_(mapping, normalized, {
          current: current,
          defer_write: true,
          fetch_comments: false,
          authoritative: false,
          validation_comments_only: true,
          preserve_closed_history: true
        });
        detected += payload.marcos.length;
        projects += 1;
      });
    } catch (error) {
      errors += 1;
      lastError = 'Recuperacao mensal ' + month + ': ' + simplifyErrorMessage_(error);
    }
  });
  writeClickUpMilestoneClosingCurrent_(sheet, headers, current);
  return { detected: detected, projects: projects, errors: errors, last_error: lastError };
}

function loadClickUpMilestoneClosingMappings_() {
  var seen = {};
  return loadProjectMappings_().filter(function(item) {
    if (!item.enabled || !(item.list_id || item.view_id || item.folder_id || item.space_id)) return false;
    var key = item.list_id ? ('list|' + item.list_id) :
      item.view_id ? ('view|' + item.view_id) :
      item.folder_id ? ('folder|' + item.folder_id) :
      ('space|' + item.space_id);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function getClickUpMilestoneAuditTaskIds_() {
  var configured = sanitizeText_(getScriptProperty_('CLICKUP_MILESTONE_AUDIT_TASK_IDS', ''));
  var ids = CLICKUP_MILESTONE_AUDIT_TASK_IDS.slice();
  if (configured) ids = ids.concat(configured.split(/[\s,;]+/));
  var seen = {};
  return ids.map(normalizeClickUpId_).filter(function(id) {
    if (!id || seen[id]) return false;
    seen[id] = true;
    return true;
  });
}

function syncClickUpMilestoneAuditTasks_() {
  var mappings = loadProjectMappings_().filter(function(item) { return item.enabled; });
  var detected = 0;
  var errors = 0;
  var lastError = '';
  getClickUpMilestoneAuditTaskIds_().forEach(function(taskId) {
    try {
      var task = clickupRequest_('get', '/task/' + normalizeClickUpId_(taskId));
      var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
      upsertClickUpMilestoneClosing_(mapping, buildNormalizedMilestoneCoverageProject_(mapping, task));
      detected += 1;
    } catch (error) {
      errors += 1;
      lastError = 'Leitura direta do marco ' + taskId + ': ' + simplifyErrorMessage_(error);
    }
  });
  return { detected: detected, errors: errors, last_error: lastError };
}

function syncClickUpRecentMilestoneCoverage_(options) {
  options = options || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(options.interactive ? 15000 : 1000)) {
    return { detected: 0, scanned: 0, errors: 0, already_running: true, last_error: '' };
  }
  try {
  var props = PropertiesService.getScriptProperties();
  var mappings = loadClickUpMilestoneClosingMappings_();
  var tasks = [];
  var errors = 0;
  var lastError = '';
  var startedAt = new Date().getTime();
  var previousCursor = toInt_(props.getProperty('CLICKUP_MILESTONE_INCREMENTAL_SINCE'), 0);
  var normalSince = previousCursor
    ? Math.max(0, previousCursor - 5 * 60 * 1000)
    : startedAt - 24 * 60 * 60 * 1000;
  var since = options.interactive
    ? Math.max(normalSince, startedAt - 2 * 60 * 60 * 1000)
    : normalSince;
  try {
    tasks = fetchClickUpRecentMilestoneCoverageTasks_(since, {
      max_pages: options.interactive ? 1 : 10
    });
  } catch (error) {
    errors += 1;
    lastError = 'Atualizacao recente: ' + simplifyErrorMessage_(error);
  }
  if (options.confirm_task_ids && options.confirm_task_ids.length) {
    try {
      tasks = tasks.concat(fetchClickUpTasksByIds_(options.confirm_task_ids));
    } catch (error) {
      errors += 1;
      lastError = 'Confirmacao direta dos marcos: ' + simplifyErrorMessage_(error);
    }
  }
  getClickUpMilestoneAuditTaskIds_().forEach(function(taskId) {
    try {
      tasks.push(clickupRequest_('get', '/task/' + normalizeClickUpId_(taskId)));
    } catch (error) {
      errors += 1;
      lastError = 'Leitura direta do marco ' + taskId + ': ' + simplifyErrorMessage_(error);
    }
  });
  tasks = dedupeTasks_(tasks);
  tasks.sort(function(a, b) {
    return Number(b && b.date_updated || 0) - Number(a && a.date_updated || 0);
  });
  var entries = tasks.filter(function(task) {
    return isMilestoneTask_(task);
  }).map(function(task) {
    var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
    return { mapping: mapping, normalized: buildNormalizedMilestoneCoverageProject_(mapping, task) };
  });
  upsertClickUpMilestoneClosingEntries_(entries, {
    authoritative: options.authoritative === true,
    validation_comments_only: options.validation_comments_only === true,
    preserve_closed_history: options.preserve_closed_history === true
  });
  if (!errors) props.setProperty('CLICKUP_MILESTONE_INCREMENTAL_SINCE', String(startedAt));
  return {
    detected: entries.length,
    scanned: tasks.length,
    since: new Date(since).toISOString(),
    errors: errors,
    last_error: lastError
  };
  } finally {
    lock.releaseLock();
  }
}

function syncClickUpRecentMilestoneAndGetClosing_(params) {
  var confirmTaskIds = sanitizeText_((params || {}).task_ids).split(/[\s,;]+/).filter(function(id) {
    return !!normalizeClickUpId_(id);
  });
  var recent = syncClickUpRecentMilestoneCoverage_({
    authoritative: true,
    interactive: true,
    validation_comments_only: true,
    confirm_task_ids: confirmTaskIds
  });
  if (recent.already_running) {
    throw new Error('Outra atualização recente ainda está finalizando. Tente novamente em alguns segundos.');
  }
  var result = getClickUpMilestoneClosing_(params || {});
  result.recent_sync = recent;
  return result;
}

function diagnoseClickUpMilestoneTask_(params) {
  params = params || {};
  var taskId = normalizeClickUpId_(params.task_id || params.id || params.query || params.link || params.url);
  if (!taskId) throw new Error('Informe task_id.');
  var task = clickupRequest_('get', '/task/' + encodeURIComponent(taskId));
  var status = clickUpTaskStatusText_(task);
  var mapping = findProjectMappingForTask_(task, loadClickUpMilestoneClosingMappings_()) || fallbackProjectMappingForTask_(task);
  var normalized = buildNormalizedMilestoneCoverageProject_(mapping, task);
  var milestone = normalized.marcos && normalized.marcos[0] || {};
  return {
    ok: true,
    task_id: String(task && task.id || taskId),
    name: sanitizeText_(task && task.name),
    status: status,
    situation: clickUpMilestoneSituation_(status),
    is_milestone: isMilestoneTask_(task),
    is_delivery: isProjectDeliveryTask_(task),
    is_closing_tracked: isClosingTrackedTask_(task),
    custom_item_id: String(task && task.custom_item_id || ''),
    custom_item_name: clickUpTaskCustomItemName_(task),
    list_id: normalizeClickUpId_(task && task.list && task.list.id),
    list_name: sanitizeText_(task && task.list && task.list.name),
    folder_id: normalizeClickUpId_(task && task.folder && task.folder.id),
    folder_name: sanitizeText_(task && task.folder && task.folder.name),
    project_key: mapping && mapping.project_key || '',
    project_name: mapping && mapping.cliente || '',
    date_closed: task && task.date_closed ? fromMillisIso_(task.date_closed) : '',
    date_updated: task && task.date_updated ? fromMillisIso_(task.date_updated) : '',
    normalized_marco: milestone
  };
}

function syncClickUpMilestoneTask_(params) {
  params = params || {};
  var diagnosis = diagnoseClickUpMilestoneTask_(params);
  if (!diagnosis.is_closing_tracked) {
    diagnosis.synced = false;
    diagnosis.reason = 'Task nao reconhecida como marco/entrega rastreavel.';
    return diagnosis;
  }
  var task = clickupRequest_('get', '/task/' + encodeURIComponent(diagnosis.task_id));
  var mappings = loadClickUpMilestoneClosingMappings_();
  var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
  upsertClickUpMilestoneClosing_(mapping, buildNormalizedMilestoneCoverageProject_(mapping, task), {
    fetch_comments: false,
    authoritative: false,
    validation_comments_only: true,
    preserve_closed_history: true
  });
  diagnosis.synced = true;
  diagnosis.after = getClickUpMilestoneClosing_({ month: sanitizeText_(params.month || params.mes).slice(0, 7) });
  return diagnosis;
}

function confirmClickUpMilestoneStatuses_(params) {
  var startedAt = new Date().getTime();
  var month = sanitizeText_((params || {}).month || (params || {}).mes).slice(0, 7);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('A base ainda está finalizando outra gravação. Tente novamente em alguns segundos.');
  }
  try {
    var sheet = getClickUpMilestoneClosingSheet_();
    var headers = getClickUpMilestoneClosingHeaders_();
    var current = loadClickUpMilestoneClosingCurrent_(sheet);
    var changed = 0;
    Object.keys(current).forEach(function(taskId) {
      var item = current[taskId] || {};
      if ((!month || sanitizeText_(item.mes_fechamento).slice(0, 7) === month) &&
          item.situacao === 'aguardando' &&
          (sanitizeText_(item.justificativa) || sanitizeText_(item.justificativa_por))) {
        item.justificativa = '';
        item.justificativa_por = '';
        changed += 1;
      }
    });
    var validatedIds = Object.keys(current).filter(function(taskId) {
      var item = current[taskId] || {};
      return (!month || sanitizeText_(item.mes_fechamento).slice(0, 7) === month) &&
        (item.situacao === 'aprovado' || item.situacao === 'reprovado');
    });
    var tasks = fetchClickUpValidationAndCurrentTasks_(validatedIds);
    tasks = dedupeTasks_(tasks).filter(function(task) {
      var taskId = String(task && task.id || '');
      var existing = current[taskId] || {};
      var taskClosedAt = task && task.date_closed ? fromMillisIso_(task.date_closed) : '';
      var taskMonth = normalizeClickUpMonthReference_(existing.mes_fechamento, taskClosedAt);
      return !month || taskMonth === month;
    });
    var validationIds = tasks.filter(function(task) {
      var status = task && task.status && (task.status.status || task.status.type || task.status.label) || '';
      var situation = clickUpMilestoneSituation_(status);
      return situation === 'aprovado' || situation === 'reprovado';
    }).map(function(task) { return String(task.id || ''); });
    var commentsById = fetchLatestClickUpTaskCommentsByIds_(validationIds);
    tasks.forEach(function(task) {
      var taskId = String(task && task.id || '');
      var previous = current[taskId] || {};
      var beforeStatus = sanitizeText_(previous.status_atual);
      var mapping = {
        project_key: sanitizeText_(previous.project_key),
        cliente: sanitizeText_(previous.projeto) || sanitizeText_(task && task.folder && task.folder.name),
        consultor: sanitizeText_(previous.consultor),
        list_id: normalizeClickUpId_(task && task.list && task.list.id),
        folder_id: normalizeClickUpId_(task && task.folder && task.folder.id),
        space_id: normalizeClickUpId_(task && task.space && task.space.id)
      };
      upsertClickUpMilestoneClosing_(mapping, buildNormalizedMilestoneCoverageProject_(mapping, task), {
        current: current,
        defer_write: true,
        fetch_comments: false,
        authoritative: true,
        validation_comments_only: true
      });
      var after = current[taskId];
      if (after) {
        if (after.situacao === 'aprovado' || after.situacao === 'reprovado') {
          var comment = commentsById[taskId] || { text: '', user: '' };
          after.justificativa = comment.text || '';
          after.justificativa_por = comment.user || '';
        } else {
          after.justificativa = '';
          after.justificativa_por = '';
        }
      }
      if (!after || sanitizeText_(after.status_atual) !== beforeStatus) changed += 1;
    });
    writeClickUpMilestoneClosingCurrent_(sheet, headers, current);
    var result = getClickUpMilestoneClosing_(params || {});
    result.confirmed = tasks.length;
    result.changed = changed;
    result.elapsed_ms = new Date().getTime() - startedAt;
    return result;
  } finally {
    lock.releaseLock();
  }
}

function clickUpMilestoneTaskMatchesMonth_(task, existing, month, useValidationMonth) {
  if (!month) return true;
  existing = existing || {};

  var closedAt = task && task.date_closed ? fromMillisIso_(task.date_closed) : '';
  var updatedAt = task && task.date_updated ? fromMillisIso_(task.date_updated) : '';
  var taskClosedMonth = normalizeClickUpMonthReference_(closedAt);
  var taskUpdatedMonth = normalizeClickUpMonthReference_(updatedAt);
  var existingClosedMonth = normalizeClickUpMonthReference_(existing.mes_fechamento, existing.closed_at);
  var existingValidationMonth = normalizeClickUpMonthReference_(existing.mes_validacao, existing.validation_at);

  if (useValidationMonth) {
    return taskUpdatedMonth === month || existingValidationMonth === month;
  }

  return taskClosedMonth === month ||
    taskUpdatedMonth === month ||
    existingClosedMonth === month ||
    existingValidationMonth === month;
}

function fetchClickUpMilestonesBySituationOptimized_(situation, currentIds, options) {
  options = options || {};
  var workspaceId = getClickUpWorkspaceId_();
  if (!workspaceId) throw new Error('CLICKUP_TEAM_ID nao configurado.');

  var month = sanitizeText_(options.month || '').slice(0, 7);
  var maxPages = Math.max(1, Math.min(toInt_(options.max_pages, 6), 20));
  var directLimit = Math.max(0, Math.min(toInt_(options.direct_limit, 120), 250));
  var deadlineMs = Number(options.deadline_ms || (new Date().getTime() + 240000));
  var current = options.current || {};
  var statuses = clickUpMilestoneStatusAliases_(situation);
  var tasks = [];
  var seen = {};

  function assertDeadline() {
    if (new Date().getTime() > deadlineMs) {
      throw new Error('Tempo seguro excedido durante busca de marcos ' + situation + '. Rode novamente em alguns segundos.');
    }
  }

  function addTask(task, existing) {
    if (!task || !task.id) return;
    var taskId = String(task.id || '');
    if (seen[taskId]) return;
    if (!clickUpMilestoneStatusMatchesSituation_(clickUpTaskStatusText_(task), situation)) return;
    if (!clickUpMilestoneTaskMatchesMonth_(task, existing || {}, month, situation !== 'aguardando')) return;
    if (!isClosingTrackedTask_(task)) return;
    task._confirmed_milestone = true;
    seen[taskId] = true;
    tasks.push(task);
  }

  statuses.forEach(function(status) {
    for (var page = 0; page < maxPages; page++) {
      assertDeadline();
      var response = clickupRequest_('get', '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'custom_items[]=1',
        'page=' + page,
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'));
      var batch = response.tasks || [];
      batch.forEach(function(task) { addTask(task, current[String(task && task.id || '')] || {}); });
      if (batch.length < 100) break;
      Utilities.sleep(300);
    }
  });

  statuses.forEach(function(status) {
    for (var page = 0; page < Math.min(maxPages, 5); page++) {
      assertDeadline();
      var response = clickupRequest_('get', '/team/' + workspaceId + '/task?' + [
        'include_closed=true',
        'subtasks=true',
        'page=' + page,
        'statuses[]=' + encodeURIComponent(status)
      ].join('&'));
      var batch = response.tasks || [];
      batch.filter(isProjectDeliveryTask_).forEach(function(task) { addTask(task, current[String(task && task.id || '')] || {}); });
      if (batch.length < 100) break;
      Utilities.sleep(300);
    }
  });

  var directIds = (currentIds || []).map(function(id) {
    return sanitizeText_(id);
  }).filter(function(id, index, all) {
    return !!id && all.indexOf(id) === index;
  }).slice(0, directLimit);
  if (directIds.length) {
    assertDeadline();
    try {
      fetchClickUpTasksByIds_(directIds).forEach(function(task) {
        addTask(task, current[String(task && task.id || '')] || {});
      });
      Utilities.sleep(300);
    } catch (error) {
      Logger.log('[Marcos] Falha ao confirmar lote direto: ' + simplifyErrorMessage_(error));
    }
  }

  return dedupeTasks_(tasks);
}

function clickUpMilestoneReconcileBatchIds_(current, month, mode, limit) {
  current = current || {};
  var ids = Object.keys(current).filter(function(taskId) {
    return clickUpMilestoneTaskMatchesMonth_(null, current[taskId] || {}, month, false);
  });
  ids.sort(function(a, b) {
    function score(id) {
      var item = current[id] || {};
      if (mode === 'closed') {
        if (item.situacao === 'aprovado' || item.situacao === 'reprovado') return 0;
        if (item.situacao === 'aguardando') return 1;
        return 2;
      }
      if (item.situacao === 'aguardando') return 0;
      if (item.situacao === mode) return 1;
      return 2;
    }
    var diff = score(a) - score(b);
    if (diff) return diff;
    return String(current[b] && current[b].updated_at || current[b] && current[b].sincronizado_em || '').localeCompare(
      String(current[a] && current[a].updated_at || current[a] && current[a].sincronizado_em || '')
    );
  });
  limit = Math.max(1, Math.min(Number(limit || 140), 150));
  if (!ids.length) return { ids: [], total: 0, offset: 0, next_offset: 0, done: true };
  var props = PropertiesService.getScriptProperties();
  var key = 'CLICKUP_MILESTONE_RECONCILE_CURSOR_' + mode + '_' + String(month || 'all').replace(/\D/g, '');
  var offset = Math.max(0, toInt_(props.getProperty(key), 0));
  if (offset >= ids.length) offset = 0;
  var batch = ids.slice(offset, offset + limit);
  var nextOffset = offset + batch.length;
  var done = nextOffset >= ids.length;
  props.setProperty(key, done ? '0' : String(nextOffset));
  return { ids: batch, total: ids.length, offset: offset, next_offset: done ? 0 : nextOffset, done: done };
}

function syncClickUpClosedMilestones_(params) {
  params = params || {};
  var startedAt = new Date().getTime();
  var month = sanitizeText_(params.month || params.mes).slice(0, 7) ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Outra atualização está finalizando. Tente novamente em alguns segundos.');
  try {
    Logger.log('🔄 [Closed] Iniciando. Mes=' + month);
    var sheet = getClickUpMilestoneClosingSheet_();
    var headers = getClickUpMilestoneClosingHeaders_();
    var current = loadClickUpMilestoneClosingCurrent_(sheet);
    var mappings = loadClickUpMilestoneClosingMappings_();
    var deadlineMs = startedAt + 240000;
    var reconcile = clickUpMilestoneReconcileBatchIds_(current, month, 'closed', 140);

    var tasks = fetchClickUpMilestonesBySituationOptimized_('aguardando', reconcile.ids, {
      month: month,
      max_pages: 4,
      direct_limit: 140,
      deadline_ms: deadlineMs,
      current: current
    });

    var recentSince = Math.max(0, new Date(month + '-01T00:00:00').getTime() || (startedAt - 75 * 24 * 60 * 60 * 1000));
    try {
      tasks = tasks.concat(fetchClickUpRecentMilestoneCoverageTasks_(recentSince, { max_pages: 3 }));
    } catch (recentError) {
      Logger.log('[Closed] Busca recente parcial: ' + simplifyErrorMessage_(recentError));
    }

    tasks = dedupeTasks_(tasks).filter(function(task) {
      var taskId = String(task && task.id || '');
      var existing = current[taskId] || {};
      return clickUpMilestoneStatusMatchesSituation_(clickUpTaskStatusText_(task), 'aguardando') &&
        clickUpMilestoneTaskMatchesMonth_(task, existing, month, false);
    });
    Logger.log('📊 [Closed] Marcos para processar: ' + tasks.length);

    var changed = 0;
    tasks.forEach(function(task) {
      var taskId = String(task && task.id || '');
      var previous = current[taskId] || {};
      var beforeStatus = sanitizeText_(previous.status_atual);
      var beforeSituation = sanitizeText_(previous.situacao);
      var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
      upsertClickUpMilestoneClosing_(mapping, buildNormalizedMilestoneCoverageProject_(mapping, task), {
        current: current,
        defer_write: true,
        fetch_comments: false,
        authoritative: false,
        validation_comments_only: true,
        preserve_closed_history: true
      });
      if (current[taskId]) {
        current[taskId].justificativa = '';
        current[taskId].justificativa_por = '';
      }
      if (sanitizeText_(current[taskId] && current[taskId].status_atual) !== beforeStatus ||
          sanitizeText_(current[taskId] && current[taskId].situacao) !== beforeSituation) changed += 1;
    });

    writeClickUpMilestoneClosingCurrent_(sheet, headers, current);
    Logger.log('📊 [Closed] Finalizado. Confirmados=' + tasks.length + ' Mudancas=' + changed);

    var summary = {
      ok: true,
      confirmed: tasks.length,
      changed: changed,
      elapsed_ms: new Date().getTime() - startedAt,
      sync_type: 'closed',
      month: month,
      reconcile_processed: reconcile.ids.length,
      reconcile_total: reconcile.total,
      reconcile_next_offset: reconcile.next_offset,
      reconcile_done: reconcile.done
    };
    if (params.skip_result) return summary;
    var resultParams = {};
    Object.keys(params).forEach(function(key) { resultParams[key] = params[key]; });
    resultParams.month = month;
    delete resultParams.mes;
    var result = getClickUpMilestoneClosing_(resultParams);
    Object.keys(summary).forEach(function(key) { result[key] = summary[key]; });
    return result;
  } finally {
    lock.releaseLock();
  }
}

function syncClickUpValidationSituation_(params, situation) {
  params = params || {};
  var startedAt = new Date().getTime();
  var month = sanitizeText_(params.month || params.mes).slice(0, 7) ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Outra atualização está finalizando. Tente novamente em alguns segundos.');
  try {
    Logger.log('🔄 [' + situation + '] Iniciando. Mes=' + month);
    var sheet = getClickUpMilestoneClosingSheet_();
    var headers = getClickUpMilestoneClosingHeaders_();
    var current = loadClickUpMilestoneClosingCurrent_(sheet);
    var mappings = loadClickUpMilestoneClosingMappings_();
    var deadlineMs = startedAt + 240000;
    var reconcile = clickUpMilestoneReconcileBatchIds_(current, month, situation, 140);

    var tasks = fetchClickUpMilestonesBySituationOptimized_(situation, reconcile.ids, {
      month: month,
      max_pages: 4,
      direct_limit: 140,
      deadline_ms: deadlineMs,
      current: current
    });

    var monthStartMs = new Date(month + '-01T00:00:00').getTime();
    if (monthStartMs) {
      try {
        tasks = tasks.concat(fetchClickUpRecentMilestoneCoverageTasks_(monthStartMs, { max_pages: 3 }));
      } catch (recentError) {
        Logger.log('[' + situation + '] Busca recente parcial: ' + simplifyErrorMessage_(recentError));
      }
    }

    tasks = dedupeTasks_(tasks).filter(function(task) {
      var taskId = String(task && task.id || '');
      var existing = current[taskId] || {};
      return clickUpMilestoneStatusMatchesSituation_(clickUpTaskStatusText_(task), situation) &&
        clickUpMilestoneTaskMatchesMonth_(task, existing, month, true);
    });
    Logger.log('📊 [' + situation + '] Marcos para processar: ' + tasks.length);

    var matchingIds = tasks.filter(function(task) {
      var taskId = String(task && task.id || '');
      var existing = current[taskId] || {};
      return !existing.task_id ||
        existing.situacao !== situation ||
        sanitizeText_(existing.status_atual) !== clickUpTaskStatusText_(task) ||
        !sanitizeText_(existing.justificativa);
    }).map(function(task) {
      return String(task && task.id || '');
    }).filter(function(id, index, all) {
      return !!id && all.indexOf(id) === index;
    });
    var commentsById = fetchLatestClickUpTaskCommentsByIds_(matchingIds);
    var changed = 0;

    tasks.forEach(function(task) {
      var taskId = String(task && task.id || '');
      var previous = current[taskId] || {};
      var beforeStatus = sanitizeText_(previous.status_atual);
      var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
      upsertClickUpMilestoneClosing_(mapping, buildNormalizedMilestoneCoverageProject_(mapping, task), {
        current: current,
        defer_write: true,
        fetch_comments: false,
        authoritative: false,
        validation_comments_only: true,
        preserve_closed_history: true
      });
      var after = current[taskId];
      if (after && after.situacao === situation) {
        var comment = commentsById[taskId] || { text: '', user: '' };
        after.justificativa = comment.text || '';
        after.justificativa_por = comment.user || '';
      } else if (after) {
        after.justificativa = '';
        after.justificativa_por = '';
      }
      if (!after ||
          sanitizeText_(after.status_atual) !== beforeStatus ||
          sanitizeText_(after.situacao) !== sanitizeText_(previous.situacao)) changed += 1;
    });
    writeClickUpMilestoneClosingCurrent_(sheet, headers, current);
    var result = getClickUpMilestoneClosing_(params);
    result.ok = true;
    result.month = month;
    result.confirmed = tasks.length;
    result.changed = changed;
    result.elapsed_ms = new Date().getTime() - startedAt;
    result.sync_type = situation;
    result.reconcile_processed = reconcile.ids.length;
    result.reconcile_total = reconcile.total;
    result.reconcile_next_offset = reconcile.next_offset;
    result.reconcile_done = reconcile.done;
    Logger.log('📊 [' + situation + '] Finalizado. Confirmados=' + tasks.length + ' Mudancas=' + changed);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function syncClickUpMilestoneHistoryCoverage_() {
  var mappings = loadClickUpMilestoneClosingMappings_();
  var tasks = [];
  var errors = 0;
  var lastError = '';
  try {
    tasks = fetchClickUpMilestoneCoverageTasks_();
  } catch (error) {
    errors += 1;
    lastError = 'Historico de marcos: ' + simplifyErrorMessage_(error);
  }
  var entries = tasks.map(function(task) {
    var mapping = findProjectMappingForTask_(task, mappings) || fallbackProjectMappingForTask_(task);
    return { mapping: mapping, normalized: buildNormalizedMilestoneCoverageProject_(mapping, task) };
  });
  // Comentarios continuam sendo enriquecidos pela atualizacao incremental e
  // pela carga detalhada. Evita uma chamada extra por marco no historico.
  upsertClickUpMilestoneClosingEntries_(entries, { fetch_comments: false });
  return { detected: entries.length, errors: errors, last_error: lastError };
}

function ensureClickUpMilestoneHistoryCoverage_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_MILESTONE_HISTORY_REBUILT') === '1') {
    return { detected: 0, errors: 0, already_rebuilt: true, last_error: '' };
  }
  var result = syncClickUpMilestoneHistoryCoverage_();
  if (!result.errors) props.setProperty('CLICKUP_MILESTONE_HISTORY_REBUILT', '1');
  return result;
}

function findProjectMappingForTask_(task, mappings) {
  var listId = normalizeClickUpId_(task && task.list && task.list.id);
  var folderId = normalizeClickUpId_(task && task.folder && task.folder.id);
  var spaceId = normalizeClickUpId_(task && task.space && task.space.id);
  var taskNames = [
    sanitizeText_(task && task.folder && task.folder.name),
    sanitizeText_(task && task.project && task.project.name),
    sanitizeText_(task && task.list && task.list.name)
  ].filter(function(name) { return !!name; }).map(normalizeKey_);
  return (mappings || loadProjectMappings_()).filter(function(mapping) {
    if (listId && String(mapping.list_id || '') === listId) return true;
    if (folderId && String(mapping.folder_id || '') === folderId) return true;
    if (spaceId && String(mapping.space_id || '') === spaceId && taskNames.indexOf(normalizeKey_(mapping.cliente)) >= 0) return true;
    return taskNames.indexOf(normalizeKey_(mapping.cliente)) >= 0;
  })[0] || null;
}

function fallbackProjectMappingForTask_(task) {
  var projectName = sanitizeText_(task && task.folder && task.folder.name) ||
    sanitizeText_(task && task.project && task.project.name) ||
    sanitizeText_(task && task.list && task.list.name) ||
    'Projeto nao mapeado';
  var hierarchyId = normalizeClickUpId_(task && task.folder && task.folder.id) ||
    normalizeClickUpId_(task && task.list && task.list.id) ||
    normalizeClickUpId_(task && task.space && task.space.id) ||
    normalizeClickUpId_(task && task.id);
  return {
    enabled: true,
    mes: '',
    cliente: projectName,
    project_key: 'CLICKUP|' + hierarchyId,
    project_url: '',
    view_id: '',
    list_id: normalizeClickUpId_(task && task.list && task.list.id),
    folder_id: normalizeClickUpId_(task && task.folder && task.folder.id),
    space_id: normalizeClickUpId_(task && task.space && task.space.id)
  };
}

function buildNormalizedMilestoneCoverageProject_(mapping, task) {
  if (!isClosingTrackedTask_(task)) {
    return { cliente: mapping.cliente || '', consultor: '', marcos: [] };
  }
  var status = task && task.status && (task.status.status || task.status.type || task.status.label) || '';
  var responsible = (task && task.assignees || []).map(function(user) {
    return sanitizeText_(user && (user.username || user.name || user.email));
  }).filter(function(name) { return !!name; }).join(', ');
  return {
    cliente: mapping.cliente || '',
    consultor: responsible,
    marcos: [{
      id: String(task && task.id || ''),
      nome: sanitizeText_(task && task.name),
      fase_nome: sanitizeText_(task && task.list && task.list.name),
      status_original: status,
      custom_item_id: String(task && task.custom_item_id || ''),
      custom_item_name: clickUpTaskCustomItemName_(task),
      responsaveis: responsible,
      task_url: sanitizeText_(task && (task.url || task.permalink || task.link || task.html_url)) ||
        ('https://app.clickup.com/t/' + String(task && task.id || '')),
      date_closed: task && task.date_closed ? fromMillisIso_(task.date_closed) : '',
      updated_at: task && task.date_updated ? fromMillisIso_(task.date_updated) : ''
    }]
  };
}

function getClickUpMilestoneClosingBackgroundStatus_() {
  var props = PropertiesService.getScriptProperties();
  return {
    active: props.getProperty('CLICKUP_MILESTONE_CLOSING_ACTIVE') === '1',
    total: toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_TOTAL'), 0),
    processed: Math.max(
      toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_PROCESSED'), 0),
      toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_OFFSET'), 0)
    ),
    errors: toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_ERRORS'), 0),
    detected: toInt_(props.getProperty('CLICKUP_MILESTONE_CLOSING_DETECTED'), 0),
    phase: props.getProperty('CLICKUP_MILESTONE_CLOSING_PHASE') || '',
    started_at: props.getProperty('CLICKUP_MILESTONE_CLOSING_STARTED_AT') || '',
    updated_at: props.getProperty('CLICKUP_MILESTONE_CLOSING_UPDATED_AT') || '',
    completed_at: props.getProperty('CLICKUP_MILESTONE_CLOSING_COMPLETED_AT') || '',
    error: props.getProperty('CLICKUP_MILESTONE_CLOSING_ERROR') || ''
  };
}

function scheduleClickUpMilestoneClosingBackground_(delayMs) {
  clearClickUpMilestoneClosingBackgroundTriggers_();
  ScriptApp.newTrigger('continueClickUpMilestoneClosingBackgroundTrigger')
    .timeBased()
    .after(Math.max(1000, Number(delayMs || 15000)))
    .create();
}

function clearClickUpMilestoneClosingBackgroundTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'continueClickUpMilestoneClosingBackgroundTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sincronizarAtividadeUsuariosClickUp() {
  var result = syncClickUpUserActivity_({});
  SpreadsheetApp.getUi().alert(
    'Atividade ClickUp sincronizada',
    'Resumo: ' + result.users + ' usuarios.\nModo: ' + (result.mode || 'audit_log') + '.\nEventos/sinais lidos: ' + result.events + '.\n\nVeja a aba "' + result.summary_sheet + '".',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function syncClickUpUserActivity_(params) {
  params = params || {};
  var workspaceId = normalizeClickUpId_(params.workspace_id) || getClickUpWorkspaceId_();
  if (!workspaceId) throw new Error('Configure CLICKUP_TEAM_ID nas propriedades do Apps Script ou informe workspace_id.');

  var now = new Date();
  var days = Math.max(1, Math.min(toInt_(params.days, toInt_(getScriptProperty_('CLICKUP_ACTIVITY_DAYS', 90), 90)), 365));
  var endMs = parseClickUpAuditTimeParam_(params.end_time, now.getTime());
  var startMs = parseClickUpAuditTimeParam_(params.start_time, endMs - days * 24 * 60 * 60 * 1000);
  if (startMs >= endMs) throw new Error('Periodo invalido para auditoria ClickUp.');

  var members = fetchClickUpWorkspaceMembers_(
    workspaceId,
    toInt_(params.scan_offset, 0) === 0 &&
      String(params.resume_scan || '') !== '1' &&
      String(params.retry_errors || '') !== '1'
  );
  if (String(params.force_estimated || '') === '1') {
    return syncClickUpUserActivityApprox_(params, {
      workspace_id: workspaceId,
      start_ms: startMs,
      end_ms: endMs,
      fetched_at: now,
      members: members,
      audit_warnings: ['Modo estimado solicitado pelo painel.']
    });
  }
  var audit = fetchClickUpAuditEvents_(workspaceId, {
    start_ms: startMs,
    end_ms: endMs,
    max_events: toInt_(params.max_events, toInt_(getScriptProperty_('CLICKUP_ACTIVITY_MAX_EVENTS', 2000), 2000)),
    page_rows: toInt_(params.page_rows, toInt_(getScriptProperty_('CLICKUP_ACTIVITY_PAGE_ROWS', 100), 100)),
    max_pages: toInt_(params.max_pages, toInt_(getScriptProperty_('CLICKUP_ACTIVITY_MAX_PAGES', 8), 8)),
    applicabilities: getClickUpAuditApplicabilities_(params.applicabilities)
  });
  if (!audit.events.length) {
    return syncClickUpUserActivityApprox_(params, {
      workspace_id: workspaceId,
      start_ms: startMs,
      end_ms: endMs,
      fetched_at: now,
      members: members,
      audit_warnings: audit.warnings
    });
  }
  var summary = buildClickUpUserActivitySummary_(members, audit.events, {
    start_ms: startMs,
    end_ms: endMs,
    fetched_at: now
  });

  writeClickUpUserActivitySummary_(summary.rows);
  writeClickUpAuditLogRows_(audit.events, {
    start_ms: startMs,
    end_ms: endMs,
    fetched_at: now
  });

  return {
    ok: true,
    workspace_id: workspaceId,
    start_time: new Date(startMs).toISOString(),
    end_time: new Date(endMs).toISOString(),
    mode: 'audit_log',
    users: summary.rows.length,
    events: audit.events.length,
    warnings: audit.warnings,
    summary_sheet: getClickUpUserActivitySheetName_(),
    raw_sheet: getClickUpAuditLogSheetName_()
  };
}

function startClickUpUserActivityBackground_(params) {
  params = params || {};
  var props = PropertiesService.getScriptProperties();
  var alreadyActive = props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE') === '1';
  var existingRows = readClickUpUserActivityRows_();
  var previousComplete = existingRows.length &&
    String(existingRows[0].sincronizacao_completa_controle || '').toLowerCase() === 'sim';
  if (alreadyActive && previousComplete) {
    props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE', '0');
    clearClickUpUserActivityBackgroundTriggers_();
    alreadyActive = false;
  }
  if (!alreadyActive && previousComplete) {
    existingRows.forEach(function(row) {
      row.projetos_lidos_controle = 0;
      row.projetos_com_erro_controle = 0;
      row.projetos_erros_json_controle = '[]';
      row.projetos_proximo_offset_controle = 0;
      row.sincronizacao_completa_controle = 'nao';
    });
    writeClickUpUserActivitySummary_(existingRows, { auto_resize: false });
  }
  props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE', '1');
  props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_FAILURES', '0');
  if (!alreadyActive) props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_STARTED_AT', new Date().toISOString());
  props.deleteProperty('CLICKUP_ACTIVITY_BACKGROUND_ERROR');
  scheduleClickUpUserActivityBackground_(1000);
  return {
    ok: true,
    scheduled: true,
    already_active: alreadyActive,
    message: 'Estimativa agendada em segundo plano. A pagina pode ser fechada.'
  };
}

function continueClickUpUserActivityBackgroundTrigger() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE') !== '1') {
    clearClickUpUserActivityBackgroundTriggers_();
    return;
  }
  try {
    var batchSize = Math.max(3, Math.min(
      toInt_(getScriptProperty_('CLICKUP_ACTIVITY_BACKGROUND_BATCH_SIZE', '12'), 12),
      20
    ));
    var result = syncClickUpUserActivity_({
      force_estimated: '1',
      resume_scan: '1',
      scan_batch_size: String(batchSize)
    });
    if (result && result.busy) {
      scheduleClickUpUserActivityBackground_(15000);
      return;
    }
    props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_FAILURES', '0');
    props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_UPDATED_AT', new Date().toISOString());
    props.deleteProperty('CLICKUP_ACTIVITY_BACKGROUND_ERROR');
    if (result && result.done) {
      props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE', '0');
      props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_COMPLETED_AT', new Date().toISOString());
      clearClickUpUserActivityBackgroundTriggers_();
      return;
    }
    scheduleClickUpUserActivityBackground_(5000);
  } catch (error) {
    var failures = toInt_(props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_FAILURES'), 0) + 1;
    props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_FAILURES', String(failures));
    props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_ERROR', simplifyErrorMessage_(error));
    if (failures >= 20) {
      props.setProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE', '0');
      clearClickUpUserActivityBackgroundTriggers_();
      return;
    }
    scheduleClickUpUserActivityBackground_(30000);
  }
}

function scheduleClickUpUserActivityBackground_(delayMs) {
  clearClickUpUserActivityBackgroundTriggers_();
  ScriptApp.newTrigger('continueClickUpUserActivityBackgroundTrigger')
    .timeBased()
    .after(Math.max(1000, Number(delayMs || 60000)))
    .create();
}

function clearClickUpUserActivityBackgroundTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'continueClickUpUserActivityBackgroundTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function syncClickUpUserActivityApprox_(params, meta) {
  params = params || {};
  meta = meta || {};
  var syncLock = LockService.getScriptLock();
  if (!syncLock.tryLock(3000)) {
    return {
      ok: false,
      busy: true,
      error: 'Uma sincronizacao ClickUp ainda esta em andamento.',
      retry_after_seconds: 15
    };
  }
  try {
  var configuredMappings = enrichClickUpActivityMappingsWithConsultants_(loadProjectMappings_()).filter(function(mapping) {
    return mapping.enabled;
  });
  var eligibleMappings = configuredMappings.filter(function(mapping) {
    return mapping.list_id || mapping.view_id || mapping.folder_id || mapping.space_id;
  });
  var existingRows = readClickUpUserActivityRows_();
  var retryProjectKey = sanitizeText_(params.retry_project_key);
  var retryMode = String(params.retry_errors || '') === '1' && !!retryProjectKey;
  var existingErrorDetails = existingRows.length ? parseClickUpActivityErrors_(existingRows[0].projetos_erros_json_controle) : [];
  var resumeScan = String(params.resume_scan || '') === '1';
  var storedNextOffset = existingRows.length ? toInt_(existingRows[0].projetos_proximo_offset_controle, 0) : 0;
  var storedComplete = existingRows.length && String(existingRows[0].sincronizacao_completa_controle || '').toLowerCase() === 'sim';
  var scanOffset = retryMode ? storedNextOffset : (resumeScan && storedNextOffset > 0 && !storedComplete
    ? storedNextOffset
    : Math.max(0, toInt_(params.scan_offset, 0)));
  var scanBatchSize = Math.max(0, Math.min(toInt_(params.scan_batch_size, 0), 30));
  var requestedLimit = toInt_(params.max_projects, 0);
  var mappings = retryMode
    ? eligibleMappings.filter(function(mapping) { return String(mapping.project_key || '') === retryProjectKey; }).slice(0, 1)
    : (scanBatchSize > 0
    ? eligibleMappings.slice(scanOffset, scanOffset + scanBatchSize)
    : (requestedLimit > 0 ? eligibleMappings.slice(0, requestedLimit) : eligibleMappings));
  var approx = buildApproxClickUpUserActivityFromTasks_(mappings, {
    members: meta.members || [],
    start_ms: meta.start_ms,
    end_ms: meta.end_ms,
    fetched_at: meta.fetched_at || new Date(),
    execution_deadline_ms: new Date().getTime() + 240000,
    project_timeout_ms: Math.max(30000, Math.min(
      toInt_(getScriptProperty_('CLICKUP_ACTIVITY_PROJECT_TIMEOUT_MS', '45000'), 45000),
      90000
    ))
  });
  var accumulatedRows = (scanOffset > 0 || retryMode)
    ? mergeClickUpUserActivityRows_(existingRows, approx.rows, meta.fetched_at || new Date())
    : approx.rows;
  var previousRead = (scanOffset > 0 || retryMode) && existingRows.length ? toInt_(existingRows[0].projetos_lidos_controle, 0) : 0;
  var errorDetails = retryMode
    ? mergeClickUpActivityErrors_(existingErrorDetails.filter(function(item) { return String(item.project_key || '') !== retryProjectKey; }), approx.errors)
    : mergeClickUpActivityErrors_(scanOffset > 0 ? existingErrorDetails : [], approx.errors);
  var cumulativeErrors = errorDetails.length;
  var nextOffset = retryMode ? storedNextOffset : scanOffset + mappings.length;
  var scanDone = retryMode ? storedComplete : nextOffset >= eligibleMappings.length;
  // Progress represents attempted projects. Failures remain visible separately and
  // must not make a completed scan look permanently stuck below the total.
  var cumulativeRead = retryMode
    ? Math.max(previousRead, storedNextOffset)
    : Math.min(eligibleMappings.length, nextOffset);
  accumulatedRows.forEach(function(item) {
    item.projetos_configurados_controle = configuredMappings.length;
    item.projetos_elegiveis_controle = eligibleMappings.length;
    item.projetos_selecionados_controle = eligibleMappings.length;
    item.projetos_lidos_controle = cumulativeRead;
    item.projetos_com_erro_controle = cumulativeErrors;
    item.projetos_erros_json_controle = JSON.stringify(errorDetails);
    item.projetos_proximo_offset_controle = nextOffset;
    item.sincronizacao_completa_controle = scanDone ? 'sim' : 'nao';
  });

  writeClickUpUserActivitySummary_(accumulatedRows, { auto_resize: scanDone });
  if (scanOffset === 0) {
    writeClickUpAuditLogRows_([], {
      start_ms: meta.start_ms,
      end_ms: meta.end_ms,
      fetched_at: meta.fetched_at || new Date()
    });
  }

  return {
    ok: true,
    workspace_id: meta.workspace_id || '',
    start_time: new Date(meta.start_ms).toISOString(),
    end_time: new Date(meta.end_ms).toISOString(),
    mode: 'estimated_from_tasks',
    users: accumulatedRows.length,
    events: approx.events,
    projects_configured: configuredMappings.length,
    projects_eligible: eligibleMappings.length,
    projects_selected: eligibleMappings.length,
    projects_read: cumulativeRead,
    projects_errors: cumulativeErrors,
    batch_projects: mappings.length,
    scan_offset: scanOffset,
    next_offset: nextOffset,
    done: scanDone,
    resumed: scanOffset > 0,
    retry_mode: retryMode,
    remaining_errors: cumulativeErrors,
    errors: approx.errors,
    warnings: (meta.audit_warnings || []).concat([
      'Audit Log indisponivel no plano atual. Controle gerado por estimativa usando tarefas, responsaveis, criadores e datas de atualizacao.'
    ]),
    summary_sheet: getClickUpUserActivitySheetName_(),
    raw_sheet: getClickUpAuditLogSheetName_()
  };
  } finally {
    syncLock.releaseLock();
  }
}

function getClickUpUserActivity_(params) {
  requireAdmin_(params || {});
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(getClickUpUserActivitySheetName_());
  var backgroundSync = getClickUpUserActivityBackgroundStatus_();
  if (!sheet) return { ok: true, users: [], total: 0, needs_generate: true, background_sync: backgroundSync, sheet: getClickUpUserActivitySheetName_() };
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, users: [], total: 0, needs_generate: true, background_sync: backgroundSync, sheet: getClickUpUserActivitySheetName_() };
  var header = values[0];
  var users = values.slice(1).map(function(row) {
    return rowToObject_(header, row);
  });
  var stale = users.length > 0 && !users.some(function(user) {
    return String(user.modo_controle || '') === 'estimado_por_tarefas';
  });
  return {
    ok: true,
    users: users,
    total: users.length,
    stale_schema: stale,
    needs_generate: stale,
    background_sync: backgroundSync,
    sheet: getClickUpUserActivitySheetName_()
  };
}

function getClickUpUserActivityBackgroundStatus_() {
  var props = PropertiesService.getScriptProperties();
  var rows = readClickUpUserActivityRows_();
  var progress = rows.length ? rows[0] : {};
  var complete = String(progress.sincronizacao_completa_controle || '').toLowerCase() === 'sim';
  var active = props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_ACTIVE') === '1' && !complete;
  return {
    active: active,
    started_at: props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_STARTED_AT') || '',
    updated_at: props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_UPDATED_AT') || '',
    completed_at: props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_COMPLETED_AT') || '',
    error: props.getProperty('CLICKUP_ACTIVITY_BACKGROUND_ERROR') || '',
    projects_read: toInt_(progress.projetos_lidos_controle, 0),
    projects_total: toInt_(progress.projetos_selecionados_controle, 0),
    projects_errors: toInt_(progress.projetos_com_erro_controle, 0)
  };
}

function readClickUpUserActivityRows_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(getClickUpUserActivitySheetName_());
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var header = values[0];
  return values.slice(1).map(function(row) {
    return rowToObject_(header, row);
  });
}

function mergeClickUpUserActivityRows_(existingRows, batchRows, fetchedAt) {
  var byKey = {};
  (existingRows || []).forEach(function(row) {
    var key = clickUpUserActivityKey_(row.email, row.user_id);
    if (key) byKey[key] = row;
  });
  (batchRows || []).forEach(function(batch) {
    var key = clickUpUserActivityKey_(batch.email, batch.user_id);
    if (!key) return;
    var current = byKey[key];
    if (!current) {
      byKey[key] = batch;
      return;
    }
    mergeClickUpActivityIdentity_(current, batch);
    mergeClickUpActivityFirst_(current, batch, 'primeiro_evento', []);
    mergeClickUpActivityFirst_(current, batch, 'primeira_acao', ['tipo_primeira_acao', 'primeira_acao_contexto', 'primeira_acao_link']);
    mergeClickUpActivityLast_(current, batch, 'ultimo_evento', []);
    mergeClickUpActivityLast_(current, batch, 'ultima_acao', ['tipo_ultima_acao', 'ultima_acao_contexto', 'ultima_acao_link']);
    [
      'total_eventos_periodo',
      'total_acoes_periodo',
      'tarefas_atribuidas',
      'tarefas_concluidas_estimadas',
      'tarefas_criadas_estimadas',
      'tarefas_atualizadas_hoje',
      'tarefas_concluidas_hoje',
      'tarefas_criadas_hoje'
    ].forEach(function(field) {
      current[field] = toInt_(current[field], 0) + toInt_(batch[field], 0);
    });
    var previousProjectCount = toInt_(current.projetos_associados, 0);
    var batchProjectCount = toInt_(batch.projetos_associados, 0);
    current.projetos_carteira_json = mergeClickUpPortfolioProjectsJson_(
      current.projetos_carteira_json,
      batch.projetos_carteira_json
    );
    var portfolioProjectCount = parseClickUpPortfolioProjects_(current.projetos_carteira_json).length;
    current.projetos_associados = portfolioProjectCount || (previousProjectCount + batchProjectCount);
    current.atividades_hoje_json = mergeClickUpTodayActionsJson_(current.atividades_hoje_json, batch.atividades_hoje_json);
    current.atividades_7_dias_json = mergeClickUpTodayActionsJson_(current.atividades_7_dias_json, batch.atividades_7_dias_json, 300);
    current.sincronizado_em = batch.sincronizado_em || current.sincronizado_em;
    current.periodo_inicio = current.periodo_inicio || batch.periodo_inicio;
    current.periodo_fim = batch.periodo_fim || current.periodo_fim;
    current.modo_controle = 'estimado_por_tarefas';
  });
  var now = fetchedAt instanceof Date ? fetchedAt : new Date(fetchedAt || new Date());
  return Object.keys(byKey).map(function(key) {
    var row = byKey[key];
    row.dias_sem_acao = row.ultima_acao ? daysBetween_(new Date(row.ultima_acao), now) : '';
    return row;
  }).sort(function(a, b) {
    var aLast = a.ultima_acao || a.clickup_last_active || '';
    var bLast = b.ultima_acao || b.clickup_last_active || '';
    return String(bLast).localeCompare(String(aLast));
  });
}

function mergeClickUpActivityIdentity_(current, batch) {
  ['user_id', 'nome', 'email', 'role', 'date_joined', 'date_invited'].forEach(function(field) {
    if (!current[field] && batch[field]) current[field] = batch[field];
  });
  current.clickup_last_active = laterIso_(current.clickup_last_active, batch.clickup_last_active);
}

function mergeClickUpActivityFirst_(current, batch, dateField, relatedFields) {
  var currentMs = normalizeClickUpDateMillis_(current[dateField]);
  var batchMs = normalizeClickUpDateMillis_(batch[dateField]);
  if (!batchMs || (currentMs && currentMs <= batchMs)) return;
  current[dateField] = batch[dateField];
  (relatedFields || []).forEach(function(field) {
    current[field] = batch[field] || '';
  });
}

function mergeClickUpActivityLast_(current, batch, dateField, relatedFields) {
  var currentMs = normalizeClickUpDateMillis_(current[dateField]);
  var batchMs = normalizeClickUpDateMillis_(batch[dateField]);
  if (!batchMs || currentMs >= batchMs) return;
  current[dateField] = batch[dateField];
  (relatedFields || []).forEach(function(field) {
    current[field] = batch[field] || '';
  });
}

function normalizeClickUpDateMillis_(value) {
  if (!value) return 0;
  var date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? 0 : date.getTime();
}

function mergeClickUpTodayActionsJson_(currentJson, batchJson, limit) {
  var seen = {};
  var merged = [];
  [currentJson, batchJson].forEach(function(raw) {
    var actions = [];
    if (Array.isArray(raw)) {
      actions = raw;
    } else if (raw) {
      try {
        actions = JSON.parse(String(raw));
      } catch (e) {
        actions = [];
      }
    }
    (Array.isArray(actions) ? actions : []).forEach(function(action) {
      var key = sanitizeText_(action && action.key) || [
        sanitizeText_(action && action.link),
        sanitizeText_(action && action.tipo),
        sanitizeText_(action && action.horario)
      ].join('|');
      if (!key || seen[key]) return;
      seen[key] = true;
      merged.push(action);
    });
  });
  merged.sort(function(a, b) {
    return normalizeClickUpDateMillis_(b && b.horario) - normalizeClickUpDateMillis_(a && a.horario);
  });
  if (Number(limit || 0) > 0) merged = merged.slice(0, Number(limit));
  return JSON.stringify(merged);
}

function parseClickUpPortfolioProjects_(raw) {
  if (Array.isArray(raw)) return raw;
  if (!raw) return [];
  try {
    var parsed = JSON.parse(String(raw));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function mergeClickUpPortfolioProjectsJson_(currentJson, batchJson) {
  var byKey = {};
  parseClickUpPortfolioProjects_(currentJson).concat(parseClickUpPortfolioProjects_(batchJson)).forEach(function(project) {
    project = project || {};
    var key = sanitizeText_(project.project_key || project.cliente || project.project_url);
    if (!key) return;
    byKey[key] = {
      project_key: key,
      cliente: sanitizeText_(project.cliente || key),
      mes: sanitizeText_(project.mes),
      project_url: sanitizeText_(project.project_url),
      list_id: normalizeClickUpId_(project.list_id),
      view_id: normalizeClickUpId_(project.view_id)
    };
  });
  return JSON.stringify(Object.keys(byKey).map(function(key) {
    return byKey[key];
  }).sort(function(a, b) {
    return String(a.cliente || '').localeCompare(String(b.cliente || ''));
  }));
}

function fetchClickUpWorkspaceMembers_(workspaceId, forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'clickup_workspace_members_' + String(workspaceId);
  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {}
    }
  }
  var response = clickupRequest_('get', '/team', null);
  var teams = response && response.teams ? response.teams : [];
  var members = [];
  teams.forEach(function(team) {
    if (String(team.id || '') !== String(workspaceId)) return;
    (team.members || []).forEach(function(member) {
      var user = member.user || member || {};
      members.push({
        id: String(user.id || ''),
        name: sanitizeText_(user.username || user.name || ''),
        email: sanitizeText_(user.email || '').toLowerCase(),
        role: sanitizeText_(user.role_key || user.role || user.role_subtype || ''),
        last_active: normalizeClickUpEventDate_(user.last_active),
        date_joined: normalizeClickUpEventDate_(user.date_joined),
        date_invited: normalizeClickUpEventDate_(user.date_invited)
      });
    });
  });
  try {
    cache.put(cacheKey, JSON.stringify(members), 600);
  } catch (e) {}
  return members;
}

function buildApproxClickUpUserActivityFromTasks_(mappings, options) {
  options = options || {};
  var byKey = {};
  var errors = [];
  var eventCount = 0;
  var projectsRead = 0;
  options.day_start_ms = startOfDayMillis_(options.fetched_at || new Date());
  options.day_end_ms = options.day_start_ms + 24 * 60 * 60 * 1000 - 1;
  options.seven_day_start_ms = options.day_start_ms - 6 * 24 * 60 * 60 * 1000;

  (options.members || []).forEach(function(member) {
    ensureApproxUser_(byKey, member.email, member.id, member.name, member.role, {
      clickup_last_active: member.last_active,
      date_joined: member.date_joined,
      date_invited: member.date_invited
    });
  });

  (mappings || []).forEach(function(mapping) {
    try {
      associateApproxProjectWithConsultant_(byKey, mapping);
      var projectDeadline = new Date().getTime() + Math.max(30000, Number(options.project_timeout_ms || 150000));
      if (options.execution_deadline_ms) projectDeadline = Math.min(projectDeadline, Number(options.execution_deadline_ms));
      var payload = fetchProjectTasks_(mapping, { deadline_ms: projectDeadline });
      var tasks = payload.tasks || [];
      projectsRead += 1;
      tasks.forEach(function(task) {
        eventCount += aggregateApproxTaskForUsers_(byKey, task, mapping, options);
      });
    } catch (error) {
      errors.push({
        project_key: mapping.project_key,
        cliente: mapping.cliente,
        project_url: mapping.project_url || buildProjectUrl_(mapping, []),
        ocorrido_em: new Date().toISOString(),
        error: simplifyErrorMessage_(error)
      });
    }
  });

  var rows = Object.keys(byKey).map(function(key) {
    var item = byKey[key];
    var first = item._first_action ? new Date(item._first_action).toISOString() : '';
    var last = item._last_action ? new Date(item._last_action).toISOString() : '';
    var taskCreated = item._first_created ? new Date(item._first_created).toISOString() : '';
    item.sincronizado_em = options.fetched_at.toISOString();
    item.periodo_inicio = new Date(options.start_ms).toISOString();
    item.periodo_fim = new Date(options.end_ms).toISOString();
    item.primeiro_login = '';
    item.ultimo_login = '';
    item.primeiro_evento = taskCreated || first;
    item.ultimo_evento = last;
    item.primeira_acao = first || taskCreated;
    item.ultima_acao = last || taskCreated;
    item.tipo_primeira_acao = first ? 'Primeira tarefa atribuida/atualizada no periodo' : (taskCreated ? 'Primeira tarefa criada no periodo' : '');
    item.tipo_ultima_acao = last ? 'Ultima tarefa atribuida/atualizada no periodo' : (taskCreated ? 'Tarefa criada no periodo' : '');
    item.primeira_acao_link = (item._first_action_ctx || item._first_created_ctx || {}).link || '';
    item.primeira_acao_contexto = (item._first_action_ctx || item._first_created_ctx || {}).contexto || '';
    item.ultima_acao_link = (item._last_action_ctx || item._first_created_ctx || {}).link || '';
    item.ultima_acao_contexto = (item._last_action_ctx || item._first_created_ctx || {}).contexto || '';
    item.dias_sem_login = '';
    item.dias_sem_acao = item.ultima_acao ? daysBetween_(new Date(item.ultima_acao), options.fetched_at) : '';
    item.total_eventos_periodo = item.total_eventos_periodo || 0;
    item.total_acoes_periodo = item.total_acoes_periodo || 0;
    var portfolioProjects = Object.keys(item._portfolio_projects || {}).map(function(projectKey) {
      return item._portfolio_projects[projectKey];
    }).filter(function(project) { return !!project; }).sort(function(a, b) {
      return String(a.cliente || '').localeCompare(String(b.cliente || ''));
    });
    item.projetos_associados = portfolioProjects.length;
    item.projetos_carteira_json = JSON.stringify(portfolioProjects);
    item.tarefas_atribuidas = item.tarefas_atribuidas || 0;
    item.tarefas_concluidas_estimadas = item.tarefas_concluidas_estimadas || 0;
    item.tarefas_criadas_estimadas = item.tarefas_criadas_estimadas || 0;
    item.tarefas_atualizadas_hoje = item.tarefas_atualizadas_hoje || 0;
    item.tarefas_concluidas_hoje = item.tarefas_concluidas_hoje || 0;
    item.tarefas_criadas_hoje = item.tarefas_criadas_hoje || 0;
    item.atividades_hoje_json = JSON.stringify(item._today_actions || []);
    item.atividades_7_dias_json = JSON.stringify((item._seven_day_actions || []).sort(function(a, b) {
      return normalizeClickUpDateMillis_(b && b.horario) - normalizeClickUpDateMillis_(a && a.horario);
    }).slice(0, 300));
    item.modo_controle = 'estimado_por_tarefas';
    delete item._projects;
    delete item._portfolio_projects;
    delete item._first_action;
    delete item._last_action;
    delete item._first_created;
    delete item._first_action_ctx;
    delete item._last_action_ctx;
    delete item._first_created_ctx;
    delete item._today_actions;
    delete item._seven_day_actions;
    delete item._seven_day_action_keys;
    return item;
  }).filter(function(item) {
    return item.total_eventos_periodo || item.clickup_last_active || item.email || item.nome;
  });

  rows.sort(function(a, b) {
    var aLast = a.ultima_acao || a.clickup_last_active || '';
    var bLast = b.ultima_acao || b.clickup_last_active || '';
    return String(bLast).localeCompare(String(aLast));
  });

  return {
    rows: rows,
    events: eventCount,
    projects_read: projectsRead,
    errors: errors
  };
}

function parseClickUpActivityErrors_(raw) {
  if (Array.isArray(raw)) return raw;
  if (!raw) return [];
  try {
    var parsed = JSON.parse(String(raw));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function mergeClickUpActivityErrors_(current, incoming) {
  var byKey = {};
  (current || []).concat(incoming || []).forEach(function(item) {
    item = item || {};
    var key = sanitizeText_(item.project_key || item.cliente);
    if (!key) return;
    byKey[key] = {
      project_key: key,
      cliente: sanitizeText_(item.cliente || key),
      project_url: sanitizeText_(item.project_url),
      ocorrido_em: item.ocorrido_em || new Date().toISOString(),
      error: sanitizeText_(item.error || 'Erro não identificado')
    };
  });
  return Object.keys(byKey).map(function(key) { return byKey[key]; });
}

function ensureApproxUser_(byKey, email, id, name, role, extras) {
  var key = clickUpUserActivityKey_(email, id);
  if (!key) return null;
  extras = extras || {};
  if (!byKey[key]) {
    byKey[key] = {
      user_id: sanitizeText_(id),
      nome: sanitizeText_(name),
      email: sanitizeText_(email).toLowerCase(),
      role: sanitizeText_(role),
      clickup_last_active: extras.clickup_last_active || '',
      date_joined: extras.date_joined || '',
      date_invited: extras.date_invited || '',
      total_eventos_periodo: 0,
      total_acoes_periodo: 0,
      tarefas_atribuidas: 0,
      tarefas_concluidas_estimadas: 0,
      tarefas_criadas_estimadas: 0,
      tarefas_atualizadas_hoje: 0,
      tarefas_concluidas_hoje: 0,
      tarefas_criadas_hoje: 0,
      projetos_associados: 0,
      _projects: {},
      _portfolio_projects: {},
      _first_action: 0,
      _last_action: 0,
      _first_created: 0,
      _first_action_ctx: null,
      _last_action_ctx: null,
      _first_created_ctx: null,
      _today_actions: [],
      _seven_day_actions: [],
      _seven_day_action_keys: {}
    };
  }
  var item = byKey[key];
  if (!item.user_id && id) item.user_id = sanitizeText_(id);
  if (!item.nome && name) item.nome = sanitizeText_(name);
  if (!item.email && email) item.email = sanitizeText_(email).toLowerCase();
  if (!item.role && role) item.role = sanitizeText_(role);
  if (!item.clickup_last_active && extras.clickup_last_active) item.clickup_last_active = extras.clickup_last_active;
  if (!item.date_joined && extras.date_joined) item.date_joined = extras.date_joined;
  if (!item.date_invited && extras.date_invited) item.date_invited = extras.date_invited;
  return item;
}

function findApproxUserByConsultant_(byKey, consultant) {
  var target = normalizeKey_(consultant);
  if (!target) return null;
  var exact = null;
  var partial = null;
  Object.keys(byKey || {}).some(function(key) {
    var item = byKey[key] || {};
    var name = normalizeKey_(item.nome);
    var email = normalizeKey_(item.email);
    if (target === name || target === email) {
      exact = item;
      return true;
    }
    if (!partial && ((name && (target.indexOf(name) >= 0 || name.indexOf(target) >= 0)) ||
        (email && (target.indexOf(email) >= 0 || email.indexOf(target) >= 0)))) {
      partial = item;
    }
    return false;
  });
  return exact || partial;
}

function associateApproxProjectWithConsultant_(byKey, mapping) {
  var item = findApproxUserByConsultant_(byKey, mapping && mapping.consultor);
  var projectKey = sanitizeText_(mapping && (mapping.project_key || mapping.cliente));
  if (!item || !projectKey) return null;
  if (!item._portfolio_projects) item._portfolio_projects = {};
  item._portfolio_projects[projectKey] = {
    project_key: projectKey,
    cliente: sanitizeText_(mapping && mapping.cliente),
    mes: sanitizeText_(mapping && mapping.mes),
    project_url: sanitizeText_(mapping && mapping.project_url) || buildProjectUrl_(mapping || {}, []),
    list_id: normalizeClickUpId_(mapping && mapping.list_id),
    view_id: normalizeClickUpId_(mapping && mapping.view_id)
  };
  return item;
}

function associateApproxConsultantMovement_(byKey, task, mapping, context, flags) {
  flags = flags || {};
  var item = associateApproxProjectWithConsultant_(byKey, mapping);
  if (!item) return;
  if (flags.require_direct_user_signal) return;
  if (flags.updated_today || flags.created_today) {
    addApproxTodayAction_(item, task, mapping, context, {
      updated: flags.updated_today,
      created: flags.created_today,
      completed: !!(flags.completed && flags.updated_today),
      timestamp: flags.updated_today ? flags.updated : flags.created
    });
  }
  if (flags.updated_seven_days || flags.created_seven_days) {
    addApproxSevenDayAction_(item, task, mapping, context, {
      updated: flags.updated_seven_days,
      created: flags.created_seven_days,
      completed: !!(flags.completed && flags.updated_seven_days),
      timestamp: flags.updated_seven_days ? flags.updated : flags.created
    });
  }
}

function addApproxTaskMovementForUser_(item, task, mapping, context, flags) {
  if (!item) return;
  flags = flags || {};
  if (flags.updated_today || flags.created_today) {
    addApproxTodayAction_(item, task, mapping, context, {
      updated: flags.updated_today,
      created: flags.created_today,
      completed: !!(flags.completed && flags.updated_today),
      timestamp: flags.updated_today ? flags.updated : flags.created
    });
  }
  if (flags.updated_seven_days || flags.created_seven_days) {
    addApproxSevenDayAction_(item, task, mapping, context, {
      updated: flags.updated_seven_days,
      created: flags.created_seven_days,
      completed: !!(flags.completed && flags.updated_seven_days),
      timestamp: flags.updated_seven_days ? flags.updated : flags.created
    });
  }
}

function aggregateApproxTaskForUsers_(byKey, task, mapping, options) {
  options = options || {};
  task = task || {};
  var count = 0;
  var updated = Number(task.date_updated || 0);
  var created = Number(task.date_created || 0);
  var updatedInPeriod = updated && updated >= Number(options.start_ms || 0) && updated <= Number(options.end_ms || new Date().getTime());
  var createdInPeriod = created && created >= Number(options.start_ms || 0) && created <= Number(options.end_ms || new Date().getTime());
  var updatedToday = updated && updated >= Number(options.day_start_ms || 0) && updated <= Number(options.day_end_ms || 0);
  var createdToday = created && created >= Number(options.day_start_ms || 0) && created <= Number(options.day_end_ms || 0);
  var updatedSevenDays = updated && updated >= Number(options.seven_day_start_ms || 0) && updated <= Number(options.day_end_ms || 0);
  var createdSevenDays = created && created >= Number(options.seven_day_start_ms || 0) && created <= Number(options.day_end_ms || 0);
  var done = isClosedStatus_(task.status && (task.status.status || task.status.type || task.status.label) || '');
  var projectKey = sanitizeText_(mapping.project_key || mapping.cliente || '');
  var context = buildApproxTaskContext_(task, mapping);
  var seen = {};

  associateApproxConsultantMovement_(byKey, task, mapping, context, {
    updated_today: !!updatedToday,
    created_today: !!createdToday,
    updated_seven_days: !!updatedSevenDays,
    created_seven_days: !!createdSevenDays,
    completed: !!done,
    updated: updated,
    created: created,
    require_direct_user_signal: true
  });

  (task.assignees || []).forEach(function(user) {
    var item = ensureApproxUser_(
      byKey,
      user && user.email,
      user && user.id,
      user && (user.username || user.name || user.email),
      user && (user.role || user.role_key),
      {}
    );
    if (!item) return;
    var key = clickUpUserActivityKey_(item.email, item.user_id);
    if (seen[key]) return;
    seen[key] = true;
    item.tarefas_atribuidas += 1;
    if (updatedInPeriod || createdInPeriod) {
      item.total_eventos_periodo += 1;
      item.total_acoes_periodo += 1;
    }
    if (done) item.tarefas_concluidas_estimadas += 1;
    if (updatedToday) item.tarefas_atualizadas_hoje += 1;
    if (done && updatedToday) item.tarefas_concluidas_hoje += 1;
    if (createdToday) item.tarefas_criadas_hoje += 1;
    addApproxTaskMovementForUser_(item, task, mapping, context, {
      updated_today: !!updatedToday,
      created_today: !!createdToday,
      updated_seven_days: !!updatedSevenDays,
      created_seven_days: !!createdSevenDays,
      completed: !!done,
      updated: updated,
      created: created
    });
    if (projectKey) item._projects[projectKey] = true;
    if (created && (!item._first_created || created < item._first_created)) {
      item._first_created = created;
      item._first_created_ctx = context;
    }
    if (updated && (!item._first_action || updated < item._first_action)) {
      item._first_action = updated;
      item._first_action_ctx = context;
    }
    if (updated && updated > item._last_action) {
      item._last_action = updated;
      item._last_action_ctx = context;
    }
    count += 1;
  });

  var creator = task.creator || task.created_by || null;
  if (creator) {
    var creatorItem = ensureApproxUser_(
      byKey,
      creator.email,
      creator.id,
      creator.username || creator.name || creator.email,
      creator.role || creator.role_key,
      {}
    );
    if (creatorItem) {
      creatorItem.tarefas_criadas_estimadas += 1;
      if (createdToday) creatorItem.tarefas_criadas_hoje += 1;
      if (createdInPeriod) {
        creatorItem.total_eventos_periodo += 1;
        creatorItem.total_acoes_periodo += 1;
      }
      if (projectKey) creatorItem._projects[projectKey] = true;
      if (created && (!creatorItem._first_created || created < creatorItem._first_created)) {
        creatorItem._first_created = created;
        creatorItem._first_created_ctx = context;
      }
      if (createdToday || createdSevenDays) {
        addApproxTaskMovementForUser_(creatorItem, task, mapping, context, {
          updated_today: false,
          created_today: !!createdToday,
          updated_seven_days: false,
          created_seven_days: !!createdSevenDays,
          completed: !!done,
          updated: updated,
          created: created
        });
      }
      if (created && created > creatorItem._last_action) {
        creatorItem._last_action = created;
        creatorItem._last_action_ctx = context;
      }
      count += 1;
    }
  }

  return count;
}

function addApproxTodayAction_(item, task, mapping, context, flags) {
  flags = flags || {};
  if (!item._today_actions) item._today_actions = [];
  var tipo = flags.completed ? 'Concluida' : flags.created ? 'Criada' : 'Atualizada';
  var taskId = sanitizeText_(task && task.id);
  var key = [taskId, tipo, flags.timestamp || ''].join('|');
  if (item._today_actions.some(function(action) { return action.key === key; })) return;
  item._today_actions.push({
    key: key,
    horario: flags.timestamp ? new Date(flags.timestamp).toISOString() : '',
    tipo: tipo,
    projeto: sanitizeText_(mapping && mapping.cliente || mapping && mapping.project_key || ''),
    item_tipo: isMilestoneTask_(task || {}) ? 'Marco' : 'Task',
    item_nome: sanitizeText_(task && task.name || ''),
    lista: sanitizeText_(task && task.list && task.list.name || ''),
    link: context && context.link || '',
    contexto: context && context.contexto || ''
  });
}

function addApproxSevenDayAction_(item, task, mapping, context, flags) {
  flags = flags || {};
  if (!item._seven_day_actions) item._seven_day_actions = [];
  if (!item._seven_day_action_keys) item._seven_day_action_keys = {};
  var tipo = flags.completed ? 'Concluida' : flags.created ? 'Criada' : 'Atualizada';
  var taskId = sanitizeText_(task && task.id);
  var key = [taskId, tipo, flags.timestamp || ''].join('|');
  if (item._seven_day_action_keys[key]) return;
  item._seven_day_action_keys[key] = true;
  item._seven_day_actions.push({
    key: key,
    horario: flags.timestamp ? new Date(flags.timestamp).toISOString() : '',
    tipo: tipo,
    projeto: sanitizeText_(mapping && mapping.cliente || mapping && mapping.project_key || ''),
    item_tipo: isMilestoneTask_(task || {}) ? 'Marco' : 'Task',
    item_nome: sanitizeText_(task && task.name || ''),
    lista: sanitizeText_(task && task.list && task.list.name || ''),
    link: context && context.link || '',
    contexto: context && context.contexto || ''
  });
}

function buildApproxTaskContext_(task, mapping) {
  task = task || {};
  mapping = mapping || {};
  var taskId = sanitizeText_(task.id);
  var projectName = sanitizeText_(mapping.cliente || mapping.project_key || '');
  var itemName = sanitizeText_(task.name || '');
  var typeName = isMilestoneTask_(task) ? 'Marco' : 'Task';
  var listName = sanitizeText_(task.list && task.list.name || '');
  var link =
    sanitizeText_(task.url || task.permalink || task.link || task.html_url) ||
    (taskId ? ('https://app.clickup.com/' + getClickUpWorkspaceId_() + '/t/' + taskId) : '') ||
    sanitizeText_(mapping.project_url);
  return {
    link: link,
    contexto: [projectName, typeName + ': ' + itemName, listName ? ('Lista: ' + listName) : ''].filter(function(part) {
      return !!sanitizeText_(part);
    }).join(' | ')
  };
}

function getClickUpWorkspaceId_() {
  return normalizeClickUpId_(getScriptProperty_('CLICKUP_TEAM_ID', '')) || normalizeClickUpId_(CLICKUP_DEFAULT_WORKSPACE_ID);
}

function fetchClickUpAuditEvents_(workspaceId, options) {
  options = options || {};
  var applicabilities = options.applicabilities || [];
  var maxEvents = Math.max(1, Math.min(Number(options.max_events || 2000), 10000));
  var pageRows = Math.max(1, Math.min(Number(options.page_rows || 100), 500));
  var maxPages = Math.max(1, Math.min(Number(options.max_pages || 8), 50));
  var events = [];
  var warnings = [];
  var seen = {};

  applicabilities.forEach(function(applicability) {
    var pageTimestamp = Number(options.end_ms || new Date().getTime());
    for (var page = 0; page < maxPages && events.length < maxEvents; page++) {
      var body = {
        applicability: applicability,
        filter: {
          workspaceId: String(workspaceId),
          startTime: Number(options.start_ms),
          endTime: Number(options.end_ms)
        },
        pagination: {
          pageRows: pageRows,
          pageTimestamp: pageTimestamp,
          pageDirection: 'before'
        }
      };
      var response;
      try {
        response = clickupRequestAbsolute_('post', 'https://api.clickup.com/api/v3/workspaces/' + workspaceId + '/auditlogs', body);
      } catch (error) {
        warnings.push(applicability + ': ' + simplifyErrorMessage_(error));
        break;
      }
      var rows = extractClickUpAuditRows_(response);
      if (!rows.length) break;

      var oldestTs = pageTimestamp;
      rows.forEach(function(row) {
        var event = normalizeClickUpAuditEvent_(row, applicability);
        if (!event.timestamp_ms) return;
        var key = event.id || [event.timestamp_ms, event.user_email, event.event_type, event.title, event.entity_id].join('|');
        if (seen[key]) return;
        seen[key] = true;
        events.push(event);
        if (event.timestamp_ms < oldestTs) oldestTs = event.timestamp_ms;
      });

      if (rows.length < pageRows || oldestTs >= pageTimestamp) break;
      pageTimestamp = oldestTs - 1;
    }
  });

  events.sort(function(a, b) { return a.timestamp_ms - b.timestamp_ms; });
  return { events: events.slice(0, maxEvents), warnings: warnings };
}

function buildClickUpUserActivitySummary_(members, events, meta) {
  var byKey = {};
  (members || []).forEach(function(member) {
    var key = clickUpUserActivityKey_(member.email, member.id);
    if (!key) return;
    byKey[key] = {
      user_id: member.id,
      nome: member.name,
      email: member.email,
      role: member.role,
      clickup_last_active: member.last_active,
      date_joined: member.date_joined,
      date_invited: member.date_invited,
      primeiro_login: '',
      ultimo_login: '',
      primeiro_evento: '',
      ultimo_evento: '',
      primeira_acao: '',
      ultima_acao: '',
      tipo_primeira_acao: '',
      tipo_ultima_acao: '',
      total_eventos_periodo: 0,
      total_acoes_periodo: 0
    };
  });

  (events || []).forEach(function(event) {
    var key = clickUpUserActivityKey_(event.user_email, event.user_id);
    if (!key) return;
    if (!byKey[key]) {
      byKey[key] = {
        user_id: event.user_id,
        nome: event.user_name,
        email: event.user_email,
        role: event.user_role,
        clickup_last_active: '',
        date_joined: '',
        date_invited: '',
        primeiro_login: '',
        ultimo_login: '',
        primeiro_evento: '',
        ultimo_evento: '',
        primeira_acao: '',
        ultima_acao: '',
        tipo_primeira_acao: '',
        tipo_ultima_acao: '',
        total_eventos_periodo: 0,
        total_acoes_periodo: 0
      };
    }
    var item = byKey[key];
    if (!item.nome) item.nome = event.user_name;
    if (!item.email) item.email = event.user_email;
    if (!item.role) item.role = event.user_role;
    item.total_eventos_periodo += 1;
    item.primeiro_evento = earlierIso_(item.primeiro_evento, event.timestamp);
    item.ultimo_evento = laterIso_(item.ultimo_evento, event.timestamp);
    if (isClickUpLoginEvent_(event)) {
      item.primeiro_login = earlierIso_(item.primeiro_login, event.timestamp);
      item.ultimo_login = laterIso_(item.ultimo_login, event.timestamp);
    }
    if (isClickUpUsageActionEvent_(event)) {
      item.total_acoes_periodo += 1;
      if (!item.primeira_acao || event.timestamp < item.primeira_acao) {
        item.primeira_acao = event.timestamp;
        item.tipo_primeira_acao = event.title || event.event_type;
      }
      if (!item.ultima_acao || event.timestamp > item.ultima_acao) {
        item.ultima_acao = event.timestamp;
        item.tipo_ultima_acao = event.title || event.event_type;
      }
    }
  });

  var rows = Object.keys(byKey).map(function(key) {
    var item = byKey[key];
    item.periodo_inicio = new Date(meta.start_ms).toISOString();
    item.periodo_fim = new Date(meta.end_ms).toISOString();
    item.sincronizado_em = meta.fetched_at.toISOString();
    item.dias_sem_acao = item.ultima_acao ? daysBetween_(new Date(item.ultima_acao), meta.fetched_at) : '';
    item.dias_sem_login = item.ultimo_login ? daysBetween_(new Date(item.ultimo_login), meta.fetched_at) : '';
    return item;
  });
  rows.sort(function(a, b) {
    var aLast = a.ultima_acao || a.ultimo_login || a.clickup_last_active || '';
    var bLast = b.ultima_acao || b.ultimo_login || b.clickup_last_active || '';
    return String(bLast).localeCompare(String(aLast));
  });
  return { rows: rows };
}

function writeClickUpUserActivitySummary_(rows, options) {
  options = options || {};
  var sheet = getClickUpUserActivitySheet_();
  var headers = getClickUpUserActivityHeaders_();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows.map(function(item) {
      return headers.map(function(header) {
        return item[header] === undefined ? '' : item[header];
      });
    }));
  }
  sheet.setFrozenRows(1);
  if (options.auto_resize !== false) sheet.autoResizeColumns(1, headers.length);
}

function writeClickUpAuditLogRows_(events, meta) {
  var sheet = getClickUpAuditLogSheet_();
  var headers = getClickUpAuditLogHeaders_();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (events && events.length) {
    sheet.getRange(2, 1, events.length, headers.length).setValues(events.map(function(event) {
      var item = {
        sincronizado_em: meta.fetched_at.toISOString(),
        periodo_inicio: new Date(meta.start_ms).toISOString(),
        periodo_fim: new Date(meta.end_ms).toISOString(),
        timestamp: event.timestamp,
        applicability: event.applicability,
        user_id: event.user_id,
        user_name: event.user_name,
        user_email: event.user_email,
        user_role: event.user_role,
        event_type: event.event_type,
        title: event.title,
        event_status: event.event_status,
        entity_id: event.entity_id,
        raw_json: JSON.stringify(event.raw || {})
      };
      return headers.map(function(header) {
        return item[header] === undefined ? '' : item[header];
      });
    }));
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, Math.min(headers.length, 12));
}

function getClickUpUserActivityHeaders_() {
  return [
    'sincronizado_em',
    'periodo_inicio',
    'periodo_fim',
    'user_id',
    'nome',
    'email',
    'role',
    'clickup_last_active',
    'date_joined',
    'date_invited',
    'primeiro_login',
    'ultimo_login',
    'primeiro_evento',
    'ultimo_evento',
    'primeira_acao',
    'tipo_primeira_acao',
    'primeira_acao_contexto',
    'primeira_acao_link',
    'ultima_acao',
    'tipo_ultima_acao',
    'ultima_acao_contexto',
    'ultima_acao_link',
    'dias_sem_login',
    'dias_sem_acao',
    'total_eventos_periodo',
    'total_acoes_periodo',
    'tarefas_atribuidas',
    'tarefas_concluidas_estimadas',
    'tarefas_criadas_estimadas',
    'tarefas_atualizadas_hoje',
    'tarefas_concluidas_hoje',
    'tarefas_criadas_hoje',
    'atividades_hoje_json',
    'projetos_associados',
    'projetos_carteira_json',
    'projetos_configurados_controle',
    'projetos_elegiveis_controle',
    'projetos_selecionados_controle',
    'projetos_lidos_controle',
    'projetos_com_erro_controle',
    'projetos_erros_json_controle',
    'projetos_proximo_offset_controle',
    'sincronizacao_completa_controle',
    'modo_controle',
    'atividades_7_dias_json'
  ];
}

function getClickUpAuditLogHeaders_() {
  return [
    'sincronizado_em',
    'periodo_inicio',
    'periodo_fim',
    'timestamp',
    'applicability',
    'user_id',
    'user_name',
    'user_email',
    'user_role',
    'event_type',
    'title',
    'event_status',
    'entity_id',
    'raw_json'
  ];
}

function getClickUpInventoryHeaders_() {
  return [
    'mes', 'cliente', 'consultor', 'status', 'project_key', 'project_url', 'view_id', 'list_id', 'folder_id', 'space_id',
    'tasks_concluidas', 'tasks_pendentes', 'marcos_concluidos', 'marcos_pendentes', 'fases_total', 'progresso',
    'data_ultima_atualizacao', 'dias_sem_atualizacao', 'clickup_json', 'ultima_sync_clickup', 'sync_status_clickup', 'sync_error_clickup'
  ];
}

function getClickUpMilestoneClosingHeaders_() {
  return [
    'task_id', 'project_key', 'projeto', 'consultor', 'marco', 'fase',
    'status_atual', 'situacao', 'closed_at', 'mes_fechamento',
    'validation_at', 'mes_validacao', 'justificativa', 'justificativa_por',
    'valor_bonus', 'link', 'responsaveis', 'updated_at', 'sincronizado_em',
    'status_history_json', 'item_tipo', 'marcador_entrega', 'fechamento_projeto'
  ];
}

function inferProjectStatusFromSummary_(resumo) {
  var total = (resumo.tasks_concluidas || 0) + (resumo.tasks_pendentes || 0) + (resumo.marcos_concluidos || 0) + (resumo.marcos_pendentes || 0);
  var pend = (resumo.tasks_pendentes || 0) + (resumo.marcos_pendentes || 0);
  if (total && !pend) return 'Aguardando entrega';
  return 'Em Andamento';
}

function inferConsultorFromTasks_(tasks) {
  var counts = {};
  (tasks || []).forEach(function(task) {
    (task.assignees || []).forEach(function(user) {
      var name = sanitizeText_(user && (user.username || user.name || user.email));
      if (!name) return;
      counts[name] = (counts[name] || 0) + 1;
    });
  });
  var best = '';
  Object.keys(counts).forEach(function(name) {
    if (!best || counts[name] > counts[best]) best = name;
  });
  return best;
}

function clickupRequest_(method, path, body) {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  var options = {
    method: String(method || 'get').toUpperCase(),
    muteHttpExceptions: true,
    headers: {
      Authorization: token,
      'Content-Type': 'application/json'
    }
  };
  if (body) options.payload = JSON.stringify(body);

  var url = CLICKUP_API_BASE + path;
  var maxAttempts = 3;
  var lastError = null;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var text = response.getContentText() || '{}';
      if (code >= 200 && code < 300) {
        return JSON.parse(text);
      }
      lastError = new Error('ClickUp API error ' + code + ': ' + text);
      if (attempt >= maxAttempts || !isRecoverableClickUpHttpError_(code, text)) {
        throw lastError;
      }
    } catch (error) {
      lastError = error;
      if (attempt >= maxAttempts || !isRecoverableClickUpTransportError_(error)) {
        throw error;
      }
    }

    Utilities.sleep(500 * attempt);
  }

  throw lastError || new Error('Falha desconhecida ao consultar o ClickUp.');
}

function clickupRequestAbsolute_(method, url, body) {
  var token = getScriptProperty_('CLICKUP_TOKEN');
  if (!token) throw new Error('Missing CLICKUP_TOKEN script property');
  var options = {
    method: String(method || 'get').toUpperCase(),
    muteHttpExceptions: true,
    headers: {
      Authorization: token,
      Accept: 'application/json',
      'Content-Type': 'application/json'
    }
  };
  if (body) options.payload = JSON.stringify(body);

  var maxAttempts = 3;
  var lastError = null;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var text = response.getContentText() || '{}';
      if (code >= 200 && code < 300) {
        return text ? JSON.parse(text) : {};
      }
      lastError = new Error('ClickUp API error ' + code + ': ' + text);
      if (attempt >= maxAttempts || !isRecoverableClickUpHttpError_(code, text)) {
        throw lastError;
      }
    } catch (error) {
      lastError = error;
      if (attempt >= maxAttempts || !isRecoverableClickUpTransportError_(error)) {
        throw error;
      }
    }
    Utilities.sleep(500 * attempt);
  }
  throw lastError || new Error('Falha desconhecida ao consultar o ClickUp.');
}

function getClickUpAuditApplicabilities_(value) {
  var raw = sanitizeText_(value || getScriptProperty_('CLICKUP_AUDIT_APPLICABILITIES', 'auth-and-security,user-activity,hierarchy-activity,custom-fields'));
  return raw.split(',').map(function(item) {
    return sanitizeText_(item);
  }).filter(function(item, index, arr) {
    return item && arr.indexOf(item) === index;
  });
}

function extractClickUpAuditRows_(response) {
  if (!response) return [];
  if (Array.isArray(response)) return response;
  var directKeys = ['data', 'events', 'logs', 'audit_logs', 'items', 'results', 'rows'];
  for (var i = 0; i < directKeys.length; i++) {
    if (Array.isArray(response[directKeys[i]])) return response[directKeys[i]];
  }
  var best = [];
  Object.keys(response).forEach(function(key) {
    if (Array.isArray(response[key]) && response[key].length > best.length) best = response[key];
  });
  return best;
}

function normalizeClickUpAuditEvent_(row, applicability) {
  row = row || {};
  var timestampValue = deepFindFirst_(row, ['startTime', 'start_time', 'timestamp', 'time', 'date', 'created_at', 'createdAt', 'endTime']);
  var timestamp = normalizeClickUpEventDate_(timestampValue);
  var eventType = sanitizeText_(deepFindFirst_(row, ['eventType', 'event_type', 'type', 'action']));
  var title = sanitizeText_(deepFindFirst_(row, ['title', 'event', 'name', 'actionName']));
  return {
    id: sanitizeText_(deepFindFirst_(row, ['id', 'eventId', 'event_id', 'uuid'])),
    applicability: applicability,
    timestamp: timestamp,
    timestamp_ms: timestamp ? new Date(timestamp).getTime() : 0,
    user_id: sanitizeText_(deepFindFirst_(row, ['userId', 'user_id', 'actorId', 'actor_id'])),
    user_name: sanitizeText_(deepFindFirst_(row, ['userName', 'user_name', 'username', 'actorName', 'actor_name'])),
    user_email: sanitizeText_(deepFindFirst_(row, ['userEmail', 'user_email', 'email', 'actorEmail', 'actor_email'])).toLowerCase(),
    user_role: sanitizeText_(deepFindFirst_(row, ['userRole', 'user_role', 'role'])),
    event_type: eventType,
    title: title || eventType,
    event_status: sanitizeText_(deepFindFirst_(row, ['eventStatus', 'event_status', 'status'])),
    entity_id: sanitizeText_(deepFindFirst_(row, ['entityId', 'entity_id', 'taskId', 'task_id', 'listId', 'list_id', 'folderId', 'folder_id'])),
    raw: row
  };
}

function deepFindFirst_(value, keys) {
  var keyMap = {};
  (keys || []).forEach(function(key) {
    keyMap[String(key).toLowerCase()] = true;
  });
  return deepFindFirstInValue_(value, keyMap, 0);
}

function deepFindFirstInValue_(value, keyMap, depth) {
  if (value === null || value === undefined || depth > 6) return '';
  if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) {
      var fromArray = deepFindFirstInValue_(value[i], keyMap, depth + 1);
      if (fromArray !== '' && fromArray !== null && fromArray !== undefined) return fromArray;
    }
    return '';
  }
  if (typeof value !== 'object') return '';
  var keys = Object.keys(value);
  for (var j = 0; j < keys.length; j++) {
    if (keyMap[String(keys[j]).toLowerCase()]) return value[keys[j]];
  }
  for (var k = 0; k < keys.length; k++) {
    var found = deepFindFirstInValue_(value[keys[k]], keyMap, depth + 1);
    if (found !== '' && found !== null && found !== undefined) return found;
  }
  return '';
}

function normalizeClickUpEventDate_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'number' || /^\d+$/.test(String(value))) {
    var n = Number(value);
    if (!n) return '';
    if (n < 10000000000) n = n * 1000;
    return new Date(n).toISOString();
  }
  var date = new Date(String(value));
  return isNaN(date.getTime()) ? String(value) : date.toISOString();
}

function parseClickUpAuditTimeParam_(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  if (typeof value === 'number' || /^\d+$/.test(String(value))) {
    var n = Number(value);
    return n < 10000000000 ? n * 1000 : n;
  }
  var date = new Date(String(value));
  if (isNaN(date.getTime())) return fallback;
  return date.getTime();
}

function clickUpUserActivityKey_(email, id) {
  email = sanitizeText_(email).toLowerCase();
  id = sanitizeText_(id);
  return email || (id ? 'id:' + id : '');
}

function isClickUpLoginEvent_(event) {
  var text = normalizeKey_((event.event_type || '') + ' ' + (event.title || ''));
  return text.indexOf('USER_LOGIN') >= 0 || /\bLOGIN\b/.test(text);
}

function isClickUpUsageActionEvent_(event) {
  var text = normalizeKey_((event.event_type || '') + ' ' + (event.title || ''));
  if (!text) return false;
  if (text.indexOf('USER_LOGIN') >= 0 || text.indexOf('USER_LOGOUT') >= 0) return false;
  if (/\bLOGIN\b/.test(text) || /\bLOGOUT\b/.test(text)) return false;
  return true;
}

function earlierIso_(current, candidate) {
  if (!candidate) return current || '';
  if (!current) return candidate;
  return new Date(candidate).getTime() < new Date(current).getTime() ? candidate : current;
}

function laterIso_(current, candidate) {
  if (!candidate) return current || '';
  if (!current) return candidate;
  return new Date(candidate).getTime() > new Date(current).getTime() ? candidate : current;
}

function daysBetween_(fromDate, toDate) {
  if (!fromDate || isNaN(fromDate.getTime())) return '';
  return Math.floor((toDate.getTime() - fromDate.getTime()) / (24 * 60 * 60 * 1000));
}

function startOfDayMillis_(date) {
  date = date instanceof Date ? date : new Date(date || new Date());
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function loadProjectMappings_() {
  var sheet = getConfigSheet_();
  var values = sheet.getDataRange().getValues();
  var seen = {};
  var out = [];

  function addMapping(mapping) {
    if (!mapping || !mapping.cliente || !mapping.mes) return;
    var key = normalizeProjectKey_(mapping.project_key || buildProjectKey_(mapping.mes, mapping.cliente));
    if (seen[key] !== undefined) {
      var existing = out[seen[key]];
      var existingHasId = !!(existing && (existing.list_id || existing.view_id || existing.folder_id || existing.space_id || existing.project_url));
      var mappingHasId = !!(mapping.list_id || mapping.view_id || mapping.folder_id || mapping.space_id || mapping.project_url);
      if (!existingHasId && mappingHasId) out[seen[key]] = mapping;
      return;
    }
    seen[key] = out.length;
    out.push(mapping);
  }

  if (values.length > 1) {
    getConfigHeaderCandidates_(values).forEach(function(candidate) {
      values.slice(candidate.index + 1).forEach(function(row) {
        addMapping(projectMappingFromConfigItem_(rowToObject_(candidate.header, row)));
      });
    });
  }
  loadMonthlyProjectLinkMappings_().forEach(addMapping);
  return out;
}

function loadProjectSyncMappings_() {
  var configMappings = loadProjectMappings_();
  var out = [];
  var seenRows = {};

  function addProject(project) {
    if (!project || !monthlyProjectExpectsClickUpSummary_(project)) return;
    var rowKey = [sanitizeMonth_(project.mes), String(project._sheet_row || ''), normalizeKey_(project.cliente)].join('|');
    if (seenRows[rowKey]) return;
    var match = findDiagnosticMappingForProject_(project, configMappings);
    var base = match && match.mapping || {};
    var projectUrl = sanitizeText_(project.link_projeto || project.projeto_link || base.project_url);
    var mapping = projectMappingFromConfigItem_(Object.assign({}, base, {
      enabled: true,
      mes: project.mes,
      cliente: project.cliente,
      consultor: project.consultor || base.consultor || '',
      project_key: buildProjectKey_(project.mes, project.cliente) + '|ROW|' + String(project._sheet_row || ''),
      project_url: projectUrl,
      link_projeto: projectUrl,
      view_id: project.view_id || base.view_id,
      list_id: project.list_id || base.list_id,
      folder_id: base.folder_id,
      space_id: base.space_id
    }));
    if (!mapping || !mapping.enabled) return;
    if (!(mapping.list_id || mapping.view_id || mapping.folder_id || mapping.space_id || mapping.project_url)) return;
    seenRows[rowKey] = true;
    out.push(mapping);
  }

  MONTHS.forEach(function(month) {
    try {
      getMonthlyProjectsFromSheet_(month).forEach(addProject);
    } catch (error) {}
  });
  return out;
}

function loadMonthlyProjectLinkMappings_() {
  var out = [];
  MONTHS.forEach(function(month) {
    try {
      getMonthlyProjectsFromSheet_(month).forEach(function(project) {
        var projectUrl = sanitizeText_(project && (project.link_projeto || project.projeto_link));
        if (!projectUrl || !isRecognizedClickUpSourceUrl_(projectUrl)) return;
        var rowKey = String(project._sheet_row || '').trim();
        var mapping = projectMappingFromConfigItem_({
          enabled: true,
          mes: month,
          cliente: project.cliente,
          consultor: project.consultor,
          project_key: buildProjectKey_(month, project.cliente) + (rowKey ? ('|ROW|' + rowKey) : ''),
          project_url: projectUrl,
          view_id: project.view_id,
          list_id: project.list_id,
          notes: 'Link lido da aba mensal'
        });
        if (mapping) out.push(mapping);
      });
    } catch (error) {}
  });
  return out;
}

function isRecognizedClickUpSourceUrl_(url) {
  return !!(
    extractClickUpIdFromUrl_(url, 'view') ||
    extractClickUpIdFromUrl_(url, 'list') ||
    extractClickUpIdFromUrl_(url, 'folder') ||
    extractClickUpIdFromUrl_(url, 'space')
  );
}

function enrichClickUpActivityMappingsWithConsultants_(mappings) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'clickup_activity_consultants_v2';
  var consultantMap = null;
  try {
    var cached = cache.get(cacheKey);
    if (cached) consultantMap = JSON.parse(cached);
  } catch (e) {
    consultantMap = null;
  }
  if (!consultantMap) {
    consultantMap = { by_project: {}, by_client: {} };
    function addSource(item) {
      var consultant = sanitizeText_(item && item.consultor);
      var clientKey = normalizeKey_(item && item.cliente);
      var projectKey = normalizeProjectKey_(item && item.project_key);
      if (!consultant) return;
      if (projectKey) consultantMap.by_project[projectKey] = consultant;
      if (clientKey) consultantMap.by_client[clientKey] = consultant;
    }
    MONTHS.forEach(function(month) {
      getMonthlyProjectsFromSheet_(month).forEach(addSource);
    });
    var inventorySheet = getClickUpInventorySheet_();
    var inventoryValues = inventorySheet.getDataRange().getDisplayValues();
    if (inventoryValues.length > 1) {
      var inventoryHeader = inventoryValues[0];
      inventoryValues.slice(1).forEach(function(row) {
        addSource(rowToObject_(inventoryHeader, row));
      });
    }
    try {
      cache.put(cacheKey, JSON.stringify(consultantMap), 600);
    } catch (e) {}
  }
  return (mappings || []).map(function(mapping) {
    var enriched = {};
    Object.keys(mapping || {}).forEach(function(key) { enriched[key] = mapping[key]; });
    enriched.consultor = sanitizeText_(mapping && mapping.consultor) ||
      consultantMap.by_project[normalizeProjectKey_(mapping && mapping.project_key)] ||
      consultantMap.by_client[normalizeKey_(mapping && mapping.cliente)] ||
      '';
    return enriched;
  });
}

function projectMappingFromConfigItem_(item) {
  item = item || {};
  var mes = sanitizeMonth_(item.mes);
  var cliente = sanitizeText_(item.cliente);
  var projectUrl = sanitizeText_(item.project_url || item.link_projeto || item.link_do_projeto || item.link_clickup);
  var urlViewId = extractClickUpIdFromUrl_(projectUrl, 'view');
  var urlListId = extractClickUpIdFromUrl_(projectUrl, 'list');
  var urlFolderId = extractClickUpIdFromUrl_(projectUrl, 'folder');
  var urlSpaceId = extractClickUpIdFromUrl_(projectUrl, 'space');
  var urlDefinesSource = !!(urlViewId || urlListId || urlFolderId || urlSpaceId);
  var viewId = urlDefinesSource ? urlViewId : normalizeClickUpId_(item.view_id);
  var listId = urlDefinesSource ? urlListId : normalizeClickUpId_(item.list_id);
  var folderId = urlDefinesSource ? urlFolderId : normalizeClickUpNumericId_(item.folder_id);
  var spaceId = urlDefinesSource ? urlSpaceId : normalizeClickUpNumericId_(item.space_id);
  if (!mes && !cliente && !projectUrl && !viewId && !listId && !folderId && !spaceId) return null;
  return {
    enabled: normalizeBoolean_(item.enabled, true),
    mes: mes,
    cliente: cliente,
    consultor: sanitizeText_(item.consultor || item.consultora || item.responsavel),
    project_key: String(item.project_key || buildProjectKey_(mes, cliente)),
    project_url: projectUrl,
    view_id: viewId,
    list_id: listId,
    folder_id: folderId,
    space_id: spaceId,
    sync_mode: sanitizeText_(item.sync_mode || 'list'),
    notes: sanitizeText_(item.notes)
  };
}

function validarClickUpConfig() {
  var sheet = getConfigSheet_();
  var values = sheet.getDataRange().getValues();
  var diagnostics = [];
  if (values.length <= 1) {
    diagnostics.push({
      row: '',
      status: 'erro',
      mes: '',
      cliente: '',
      project_url: '',
      view_id: '',
      list_id: '',
      diagnostico: 'CLICKUP_CONFIG vazio ou sem linhas de projeto.'
    });
  } else {
    getConfigHeaderCandidates_(values).forEach(function(candidate) {
      values.slice(candidate.index + 1).forEach(function(row, offset) {
      var item = rowToObject_(candidate.header, row);
      var mapping = projectMappingFromConfigItem_(item);
      if (!mapping) return;
      var rawCliente = mapping.cliente;
      var rawMes = mapping.mes;
      var projectUrl = mapping.project_url;
      var viewId = mapping.view_id;
      var listId = mapping.list_id;
      var folderId = mapping.folder_id;
      var spaceId = mapping.space_id;
      var enabled = mapping.enabled;
      var rowNumber = candidate.index + 2 + offset;
      var problems = [];
      if (!enabled) problems.push('Linha desabilitada.');
      if (!rawMes) problems.push('Mes nao encontrado.');
      if (!rawCliente) problems.push('Cliente/projeto nao encontrado.');
      if (!viewId && !listId && !folderId && !spaceId) problems.push('Sem view_id/list_id/folder_id/space_id e nao foi possivel extrair do link.');
      diagnostics.push({
        row: rowNumber,
        status: problems.length ? 'revisar' : 'ok',
        mes: rawMes,
        cliente: rawCliente,
        project_url: projectUrl,
        view_id: viewId,
        list_id: listId,
        folder_id: folderId,
        space_id: spaceId,
        diagnostico: problems.length ? problems.join(' ') : 'Configuracao minima OK.'
      });
      });
    });
  }
  var out = getOrCreateSheet_('CLICKUP_CONFIG_DIAGNOSTICO');
  writeObjectsToSheet_(out, diagnostics, ['row', 'status', 'mes', 'cliente', 'project_url', 'view_id', 'list_id', 'folder_id', 'space_id', 'diagnostico']);
  return {
    ok: true,
    total: diagnostics.length,
    ok_count: diagnostics.filter(function(item) { return item.status === 'ok'; }).length,
    revisar_count: diagnostics.filter(function(item) { return item.status !== 'ok'; }).length,
    sheet: 'CLICKUP_CONFIG_DIAGNOSTICO'
  };
}

function findProjectMapping_(projectKey) {
  var raw = String(projectKey || '').trim();
  if (!raw) return null;
  var normalizedInput = normalizeProjectKey_(raw);
  return loadProjectMappings_().filter(function(item) {
    if (!item) return false;
    var rawItemKey = String(item.project_key || '').trim();
    if (rawItemKey === raw) return true;
    if (normalizeProjectKey_(rawItemKey) === normalizedInput) return true;
    if (buildProjectKey_(item.mes, item.cliente) === normalizedInput) return true;
    return false;
  })[0] || null;
}

function findProjectMappingByListId_(listId) {
  return loadProjectMappings_().filter(function(item) {
    return String(item.list_id) === String(listId || '');
  })[0] || null;
}

function findProjectMappingByTaskId_(taskId) {
  var task = clickupRequest_('get', '/task/' + taskId);
  var listId = task && task.list && task.list.id;
  return listId ? findProjectMappingByListId_(listId) : null;
}

function writeProjectSummaryToMonthlySheet_(mapping, normalized) {
  var sheet = getMonthSheet_(mapping.mes);
  var headerInfo = ensureMonthlyOutputColumns_(sheet);
  var clientRow = findProjectRowForMapping_(sheet, mapping);
  if (!clientRow) throw new Error('Client row not found in month sheet: ' + mapping.cliente + ' (' + mapping.mes + ')');

  var values = {};
  values.tasks_concluidas = normalized.resumo.tasks_concluidas;
  values.tasks_pendentes = normalized.resumo.tasks_pendentes;
  values.marcos_concluidos = normalized.resumo.marcos_concluidos;
  values.marcos_pendentes = normalized.resumo.marcos_pendentes;
  values.fases_total = normalized.resumo.fases_total;
  values.progresso = normalized.resumo.progresso;
  values.data_ultima_atualizacao = normalized.resumo.data_ultima_atualizacao || normalized.synced_at || '';
  values.dias_sem_atualizacao = diffDaysFromIso_(normalized.resumo.data_ultima_atualizacao);
  values.link_projeto = normalized.project_url || '';
  values.view_id = normalized.view_id || '';
  values.list_id = normalized.list_id || '';
  values.clickup_json = buildSheetSafeClickUpJson_(normalized);
  values.ultima_sync_clickup = new Date();
  values.sync_status_clickup = 'OK';
  values.sync_error_clickup = '';

  OUTPUT_COLUMNS.forEach(function(column) {
    var col = headerInfo[column];
    sheet.getRange(clientRow, col).setValue(values[column] !== undefined ? values[column] : '');
  });
}

function writeSyncStatus_(mapping, status, errorMessage) {
  try {
    var sheet = getMonthSheet_(mapping.mes);
    var headerInfo = ensureMonthlyOutputColumns_(sheet);
    var clientRow = findProjectRowForMapping_(sheet, mapping);
    if (!clientRow) return;
    sheet.getRange(clientRow, headerInfo.sync_status_clickup).setValue(String(status || ''));
    sheet.getRange(clientRow, headerInfo.sync_error_clickup).setValue(String(errorMessage || ''));
    sheet.getRange(clientRow, headerInfo.ultima_sync_clickup).setValue(new Date());
  } catch (e) {}
}

function ensureMonthlyOutputColumns_(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1));
  var headerValues = headerRange.getValues()[0];
  var map = {};

  OUTPUT_COLUMNS.forEach(function(name) {
    var index = headerValues.indexOf(name);
    if (index < 0) {
      index = headerValues.length;
      headerValues.push(name);
      sheet.getRange(1, index + 1).setValue(name);
    }
    map[name] = index + 1;
  });
  return map;
}

function findClientRow_(sheet, cliente) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var target = normalizeKey_(cliente);
  for (var i = 0; i < names.length; i++) {
    if (normalizeKey_(names[i][0]) === target) return i + 2;
  }
  return null;
}

function findProjectRowForMapping_(sheet, mapping) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var header = normalizeMonthlyHeader_(values[0]);
  var targetListId = normalizeClickUpId_(mapping && mapping.list_id);
  var targetViewId = normalizeClickUpId_(mapping && mapping.view_id);
  var targetFolderId = normalizeClickUpNumericId_(mapping && mapping.folder_id);
  var targetSpaceId = normalizeClickUpNumericId_(mapping && mapping.space_id);
  var targetProjectUrl = sanitizeText_(mapping && mapping.project_url);

  for (var i = 1; i < values.length; i++) {
    var item = monthlyProjectFromRow_(sanitizeMonth_(mapping && mapping.mes), header, values[i], i + 1);
    if (!item) continue;
    var rowUrl = sanitizeText_(item.link_projeto || item.projeto_link);
    var rowListId = normalizeClickUpId_(item.list_id) || extractClickUpIdFromUrl_(rowUrl, 'list');
    var rowViewId = normalizeClickUpId_(item.view_id) || extractClickUpIdFromUrl_(rowUrl, 'view');
    var rowFolderId = extractClickUpIdFromUrl_(rowUrl, 'folder');
    var rowSpaceId = extractClickUpIdFromUrl_(rowUrl, 'space');

    if (targetListId && rowListId && targetListId === rowListId) return i + 1;
    if (targetViewId && rowViewId && targetViewId === rowViewId) return i + 1;
    if (targetFolderId && rowFolderId && targetFolderId === rowFolderId) return i + 1;
    if (targetSpaceId && rowSpaceId && targetSpaceId === rowSpaceId) return i + 1;
    if (targetProjectUrl && rowUrl && sanitizeText_(targetProjectUrl) === rowUrl) return i + 1;
  }
  return findClientRow_(sheet, mapping && mapping.cliente);
}

function enqueueDirtyEvent_(event) {
  var sheet = getDirtyQueueSheet_();
  ensureHeaders_(sheet, ['queued_at', 'event', 'task_id', 'list_id', 'project_key']);
  var row = [
    new Date(),
    String(event.event || ''),
    String(event.task_id || ''),
    String(event.list_id || event.folder_id || ''),
    ''
  ];
  sheet.appendRow(row);
  return { queued: true };
}

function verifyWebhookSignature_(body, signature) {
  var sheet = getWebhookSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return true;
  var header = values[0];
  var secrets = values.slice(1).map(function(row) {
    return rowToObject_(header, row).secret;
  }).filter(function(secret) { return !!secret; });
  if (!secrets.length) return true;
  return secrets.some(function(secret) {
    var digest = Utilities.computeHmacSha256Signature(body, secret);
    var hex = digest.map(function(byte) {
      var v = (byte < 0 ? byte + 256 : byte).toString(16);
      return v.length === 1 ? '0' + v : v;
    }).join('');
    return hex === signature;
  });
}

function verifyWebhookRequest_(e, body) {
  var expectedToken = getScriptProperty_('CLICKUP_WEBHOOK_TOKEN', '');
  if (expectedToken) {
    var actualToken = String(((e || {}).parameter || {}).webhook_token || '');
    return actualToken === expectedToken;
  }
  return true;
}

function getMonthSheet_(month) {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(sanitizeMonth_(month));
  if (!sheet) throw new Error('Month sheet not found: ' + month);
  return sheet;
}

function getConfigSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('CLICKUP_CONFIG_SHEET', 'CLICKUP_CONFIG');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, ['enabled', 'mes', 'cliente', 'project_key', 'project_url', 'view_id', 'list_id', 'folder_id', 'space_id', 'sync_mode', 'notes']);
  return sheet;
}

function getClickUpInventorySheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('CLICKUP_INVENTORY_SHEET', 'CLICKUP_INVENTARIO');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, getClickUpInventoryHeaders_());
  return sheet;
}

function getClickUpMilestoneClosingSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('CLICKUP_MILESTONE_CLOSING_SHEET', 'CLICKUP_FECHAMENTO_MARCOS');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, getClickUpMilestoneClosingHeaders_());
  return sheet;
}

function getClickUpUserActivitySheetName_() {
  return getScriptProperty_('CLICKUP_USER_ACTIVITY_SHEET', 'CLICKUP_USER_ACTIVITY');
}

function getClickUpUserActivitySheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(getClickUpUserActivitySheetName_()) || ss.insertSheet(getClickUpUserActivitySheetName_());
  ensureHeaders_(sheet, getClickUpUserActivityHeaders_());
  return sheet;
}

function getClickUpAuditLogSheetName_() {
  return getScriptProperty_('CLICKUP_AUDIT_LOG_SHEET', 'CLICKUP_AUDIT_LOGS');
}

function getClickUpAuditLogSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(getClickUpAuditLogSheetName_()) || ss.insertSheet(getClickUpAuditLogSheetName_());
  ensureHeaders_(sheet, getClickUpAuditLogHeaders_());
  return sheet;
}

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getWebhookSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('CLICKUP_WEBHOOKS_SHEET', 'CLICKUP_WEBHOOKS');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, ['project_key', 'list_id', 'webhook_id', 'secret']);
  return sheet;
}

function getDirtyQueueSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('CLICKUP_DIRTY_SHEET', 'CLICKUP_DIRTY_QUEUE');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, ['queued_at', 'event', 'task_id', 'list_id', 'project_key']);
  return sheet;
}

function getPanelUpdateHistorySheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('PANEL_UPDATE_HISTORY_SHEET', 'PANEL_UPDATE_HISTORY');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, [
    'logged_at',
    'quando_iso',
    'quando',
    'usuario',
    'dispositivo',
    'tipo',
    'resultado',
    'motivo',
    'projetos',
    'url',
    'navegador'
  ]);
  return sheet;
}

function logPanelUpdate_(params) {
  params = params || {};
  var sheet = getPanelUpdateHistorySheet_();
  var row = [
    new Date(),
    sanitizeText_(params.quando_iso),
    sanitizeText_(params.quando),
    sanitizeText_(params.usuario),
    sanitizeText_(params.dispositivo),
    sanitizeText_(params.tipo),
    sanitizeText_(params.resultado),
    sanitizeText_(params.motivo),
    toInt_(params.projetos, 0),
    sanitizeText_(params.url),
    sanitizeText_(params.navegador)
  ];
  sheet.appendRow(row);
  return {
    ok: true,
    logged: true,
    total_rows: Math.max(0, sheet.getLastRow() - 1)
  };
}

function getPanelUpdateHistory_(limit) {
  var sheet = getPanelUpdateHistorySheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { ok: true, history: [], total: 0 };
  }
  var header = values[0];
  var rows = values.slice(1);
  var max = Math.max(1, Math.min(Number(limit || 20), 100));
  var history = rows.slice(Math.max(0, rows.length - max)).reverse().map(function(row) {
    var item = rowToObject_(header, row);
    return {
      logged_at: item.logged_at instanceof Date ? item.logged_at.toISOString() : String(item.logged_at || ''),
      quando_iso: String(item.quando_iso || ''),
      quando: item.quando ? String(item.quando) : (item.logged_at instanceof Date ? item.logged_at.toLocaleString('pt-BR') : ''),
      usuario: String(item.usuario || ''),
      dispositivo: String(item.dispositivo || ''),
      tipo: String(item.tipo || ''),
      resultado: String(item.resultado || ''),
      motivo: String(item.motivo || ''),
      projetos: toInt_(item.projetos, 0),
      url: String(item.url || ''),
      navegador: String(item.navegador || '')
    };
  });
  return {
    ok: true,
    total: rows.length,
    history: history
  };
}

function getProjectFollowupSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('PROJECT_FOLLOWUP_SHEET', 'ACOMPANHAMENTOS_PROJETOS');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, [
    'logged_at',
    'data_acompanhamento',
    'mes',
    'cliente',
    'consultor',
    'project_key',
    'link_projeto',
    'consideracao',
    'proxima_acao',
    'responsavel',
    'origem',
    'status',
    'url',
    'navegador',
    'kanban_stage',
    'followup_id',
    'followup_status',
    'status_updated_at',
    'status_updated_by',
    'status_comment'
  ]);
  return sheet;
}

function getProjectKanbanStateSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('PROJECT_KANBAN_STATE_SHEET', 'ACOMPANHAMENTOS_KANBAN');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, [
    'updated_at',
    'cycle_key',
    'project_key',
    'stage',
    'updated_by'
  ]);
  return sheet;
}

function getUsersSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var name = getScriptProperty_('PANEL_USERS_SHEET', 'PAINEL_USUARIOS');
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, [
    'created_at',
    'username',
    'name',
    'role',
    'enabled',
    'password_salt',
    'password_hash',
    'last_login',
    'seniority',
    'daily_value'
  ]);
  ensureBootstrapAdmin_(sheet);
  return sheet;
}

function ensureBootstrapAdmin_(sheet) {
  if (sheet.getLastRow() > 1) return;
  var username = sanitizeText_(getScriptProperty_('PANEL_ADMIN_USERNAME', ''));
  var name = sanitizeText_(getScriptProperty_('PANEL_ADMIN_NAME', username));
  var passwordSha = sanitizeText_(getScriptProperty_('PANEL_ADMIN_PASSWORD_SHA256', ''));
  if (!username || !passwordSha) return;
  var salt = makeAuthToken_().slice(0, 16);
  sheet.appendRow([
    new Date(),
    username.toLowerCase(),
    name || username,
    'admin',
    'TRUE',
    salt,
    hashPassword_(salt, passwordSha),
    ''
  ]);
}

function hashPassword_(salt, passwordSha) {
  var raw = String(salt || '') + '|' + String(passwordSha || '');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function hashPlainTextPasswordForPanel_(password) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function resetPanelUserPasswordFromEditor() {
  // Utilitario manual: edite os dois valores abaixo, rode no editor do Apps Script
  // e depois apague a senha deste arquivo no editor.
  var username = 'reinaldo';
  var newPassword = 'TROQUE_ESTA_SENHA';
  if (!username || newPassword === 'TROQUE_ESTA_SENHA') {
    throw new Error('Edite username e newPassword antes de executar.');
  }
  var found = findUserRow_(username);
  if (!found) throw new Error('Usuario nao encontrado: ' + username);
  var header = found.header;
  var saltCol = header.indexOf('password_salt') + 1;
  var hashCol = header.indexOf('password_hash') + 1;
  if (!saltCol || !hashCol) throw new Error('Colunas password_salt/password_hash nao encontradas.');
  var salt = makeAuthToken_().slice(0, 16);
  var passwordSha = hashPlainTextPasswordForPanel_(newPassword);
  found.sheet.getRange(found.row, saltCol).setValue(salt);
  found.sheet.getRange(found.row, hashCol).setValue(hashPassword_(salt, passwordSha));
  return { ok: true, username: String(username).toLowerCase() };
}

function makeAuthToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function publicUser_(user) {
  return {
    username: String(user.username || ''),
    name: String(user.name || ''),
    role: normalizePanelRole_(user.role),
    enabled: String(user.enabled || '').toUpperCase() !== 'FALSE',
    last_login: user.last_login instanceof Date ? user.last_login.toISOString() : String(user.last_login || ''),
    seniority: normalizeConsultantSeniority_(user.seniority),
    daily_value: CONSULTANT_SENIORITY_RATES[normalizeConsultantSeniority_(user.seniority)] || 0
  };
}

function normalizePanelRole_(role) {
  var key = normalizeKey_(role);
  if (key === 'ADMIN' || key === 'ADMINISTRADOR') return 'admin';
  if (key === 'COORDENADOR' || key === 'COORDENADORA' || key === 'COORDINATOR') return 'coordenador';
  return 'user';
}

function userHasFullProjectAccess_(user) {
  var role = normalizePanelRole_(user && user.role);
  return role === 'admin' || role === 'coordenador';
}

function userMatchesConsultor_(user, consultor) {
  if (!user) return false;
  if (userHasFullProjectAccess_(user)) return true;
  var target = normalizeKey_(consultor);
  if (!target) return false;
  var name = normalizeKey_(user.name);
  var username = normalizeKey_(user.username);
  return (name && (target === name || target.indexOf(name) >= 0 || name.indexOf(target) >= 0)) ||
    (username && (target === username || target.indexOf(username) >= 0 || username.indexOf(target) >= 0));
}

function canUserAccessProjectItem_(user, item) {
  if (userHasFullProjectAccess_(user)) return true;
  return userMatchesConsultor_(user, item && item.consultor);
}

function findUserRow_(username) {
  username = sanitizeText_(username).toLowerCase();
  if (!username) return null;
  var sheet = getUsersSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return null;
  var header = values[0];
  for (var i = 1; i < values.length; i++) {
    var item = rowToObject_(header, values[i]);
    if (String(item.username || '').toLowerCase() === username) {
      return {
        sheet: sheet,
        row: i + 1,
        header: header,
        user: item
      };
    }
  }
  return null;
}

function sessionCache_() {
  return CacheService.getScriptCache();
}

function storeSession_(user) {
  var token = makeAuthToken_();
  sessionCache_().put('session:' + token, JSON.stringify(publicUser_(user)), 21600);
  return token;
}

function requireUser_(params) {
  var token = sanitizeText_(params && params.auth_token);
  if (!token) throw new Error('Login obrigatorio.');
  var raw = sessionCache_().get('session:' + token);
  if (!raw) throw new Error('Sessao expirada. Entre novamente.');
  var user = JSON.parse(raw);
  if (!user.enabled) throw new Error('Usuario desativado.');
  return user;
}

function requireAdmin_(params) {
  var user = requireUser_(params);
  if (String(user.role || '') !== 'admin') throw new Error('Apenas administrador pode executar esta acao.');
  return user;
}

function loginUser_(params) {
  params = params || {};
  var username = sanitizeText_(params.username).toLowerCase();
  var passwordSha = sanitizeText_(params.password_sha);
  if (!username || !passwordSha) throw new Error('Usuario e senha sao obrigatorios.');
  var found = findUserRow_(username);
  if (!found) throw new Error('Usuario ou senha invalidos.');
  var user = found.user;
  if (String(user.enabled || '').toUpperCase() === 'FALSE') throw new Error('Usuario desativado.');
  var expected = hashPassword_(user.password_salt, passwordSha);
  if (expected !== String(user.password_hash || '')) throw new Error('Usuario ou senha invalidos.');
  found.sheet.getRange(found.row, found.header.indexOf('last_login') + 1).setValue(new Date());
  var token = storeSession_(user);
  return {
    ok: true,
    token: token,
    user: publicUser_(user)
  };
}

function getCurrentUser_(params) {
  return {
    ok: true,
    user: requireUser_(params)
  };
}

function listUsers_(params) {
  requireAdmin_(params);
  var sheet = getUsersSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, users: [] };
  var header = values[0];
  var users = values.slice(1).map(function(row) {
    return publicUser_(rowToObject_(header, row));
  });
  return { ok: true, users: users };
}

function createUser_(params) {
  var admin = requireAdmin_(params);
  params = params || {};
  var username = sanitizeText_(params.username).toLowerCase();
  var name = sanitizeText_(params.name);
  var role = normalizePanelRole_(params.role);
  var passwordSha = sanitizeText_(params.password_sha);
  if (!username || !name || !passwordSha) throw new Error('Nome, usuario e senha sao obrigatorios.');
  if (findUserRow_(username)) throw new Error('Usuario ja existe.');
  var sheet = getUsersSheet_();
  var salt = makeAuthToken_().slice(0, 16);
  sheet.appendRow([
    new Date(),
    username,
    name,
    role,
    'TRUE',
    salt,
    hashPassword_(salt, passwordSha),
    ''
  ]);
  return {
    ok: true,
    created_by: admin.username,
    user: { username: username, name: name, role: role, enabled: true, last_login: '' }
  };
}

function setUserEnabled_(params) {
  requireAdmin_(params);
  params = params || {};
  var username = sanitizeText_(params.username).toLowerCase();
  var enabled = String(params.enabled || '') === '1' || String(params.enabled || '').toUpperCase() === 'TRUE';
  var found = findUserRow_(username);
  if (!found) throw new Error('Usuario nao encontrado.');
  var col = found.header.indexOf('enabled') + 1;
  found.sheet.getRange(found.row, col).setValue(enabled ? 'TRUE' : 'FALSE');
  return { ok: true, username: username, enabled: enabled };
}

function buildSheetSafeClickUpJson_(normalized) {
  var full = JSON.stringify(normalized || {});
  if (full.length <= 45000) return full;

  function compactItem_(item) {
    item = item || {};
    var out = {
      id: String(item.id || ''),
      tipo: sanitizeText_(item.tipo || ''),
      nome: sanitizeText_(item.nome || ''),
      fase_nome: sanitizeText_(item.fase_nome || ''),
      status_original: sanitizeText_(item.status_original || ''),
      fase_status_original: sanitizeText_(item.fase_status_original || ''),
      custom_item_id: String(item.custom_item_id || ''),
      custom_item_name: sanitizeText_(item.custom_item_name || ''),
      marcador_entrega: sanitizeText_(item.marcador_entrega || ''),
      marcador_marco: sanitizeText_(item.marcador_marco || ''),
      task_url: sanitizeText_(item.task_url || ''),
      date_closed: sanitizeText_(item.date_closed || ''),
      updated_at: sanitizeText_(item.updated_at || ''),
      due_date: sanitizeText_(item.due_date || '')
    };
    if (item.concluido !== undefined) out.concluido = !!item.concluido;
    if (item.concluida !== undefined) out.concluida = !!item.concluida;
    return out;
  }

  var compact = {
    project_key: normalized && normalized.project_key || '',
    cliente: normalized && normalized.cliente || '',
    mes: normalized && normalized.mes || '',
    consultor: normalized && normalized.consultor || '',
    project_url: normalized && normalized.project_url || '',
    view_id: normalized && normalized.view_id || '',
    list_id: normalized && normalized.list_id || '',
    synced_at: normalized && normalized.synced_at || new Date().toISOString(),
    compactado: true,
    compactado_motivo: 'JSON completo excedeu o limite de caracteres por celula da planilha.',
    fases: (normalized && normalized.fases || []).map(function(phase) {
      return {
        tipo: 'fase',
        id: String(phase && phase.id || ''),
        nome: sanitizeText_(phase && phase.nome || ''),
        ordem: phase && phase.ordem || 0,
        status_original: sanitizeText_(phase && phase.status_original || ''),
        custom_item_id: String(phase && phase.custom_item_id || ''),
        custom_item_name: sanitizeText_(phase && phase.custom_item_name || ''),
        marcador_entrega: sanitizeText_(phase && phase.marcador_entrega || ''),
        marcador_marco: sanitizeText_(phase && phase.marcador_marco || ''),
        updated_at: sanitizeText_(phase && phase.updated_at || ''),
        task_url: sanitizeText_(phase && phase.task_url || ''),
        tasks_concluidas: phase && phase.tasks_concluidas || 0,
        tasks_pendentes: phase && phase.tasks_pendentes || 0,
        marcos_concluidos: phase && phase.marcos_concluidos || 0,
        marcos_pendentes: phase && phase.marcos_pendentes || 0,
        tasks_total: phase && phase.tasks_total || 0,
        marcos_total: phase && phase.marcos_total || 0,
        total_itens: phase && phase.total_itens || 0,
        progresso: phase && phase.progresso || 0
      };
    }),
    tasks: (normalized && normalized.tasks || []).map(compactItem_),
    marcos: (normalized && normalized.marcos || []).map(compactItem_),
    resumo: normalized && normalized.resumo || {},
    clickup_payload: normalized && normalized.clickup_payload || {}
  };

  var compactJson = JSON.stringify(compact);
  if (compactJson.length <= 45000) return compactJson;

  delete compact.tasks;
  compact.tasks_omitidas = normalized && normalized.tasks ? normalized.tasks.length : 0;
  compactJson = JSON.stringify(compact);
  if (compactJson.length <= 45000) return compactJson;

  delete compact.marcos;
  compact.marcos_omitidos = normalized && normalized.marcos ? normalized.marcos.length : 0;
  compactJson = JSON.stringify(compact);
  if (compactJson.length <= 45000) return compactJson;

  compact.fases = [];
  compact.fases_omitidas = normalized && normalized.fases ? normalized.fases.length : 0;
  compactJson = JSON.stringify(compact);
  return compactJson.length <= 45000 ? compactJson : compactJson.slice(0, 44900);
}

function setUserSeniority_(params) {
  var admin = requireAdmin_(params || {});
  var username = sanitizeText_(params.username).toLowerCase();
  var seniority = normalizeConsultantSeniority_(params.seniority);
  if (!seniority) throw new Error('Selecione Junior, Pleno, Senior ou Master.');
  var found = findUserRow_(username);
  if (!found) throw new Error('Usuario nao encontrado.');
  var seniorityCol = found.header.indexOf('seniority') + 1;
  var dailyValueCol = found.header.indexOf('daily_value') + 1;
  if (!seniorityCol || !dailyValueCol) throw new Error('Atualize as colunas do cadastro de usuarios.');
  found.sheet.getRange(found.row, seniorityCol).setValue(seniority);
  found.sheet.getRange(found.row, dailyValueCol).setValue(CONSULTANT_SENIORITY_RATES[seniority]);
  setConsultantSeniority_({
    auth_token: params.auth_token,
    consultant_name: found.user.name || found.user.username,
    seniority: seniority
  });
  return {
    ok: true,
    updated_by: admin.name || admin.username,
    user: publicUser_(Object.assign({}, found.user, { seniority: seniority }))
  };
}

var CONSULTANT_COMPENSATION_SHEET = 'CONSULTORES_REMUNERACAO';
var CONSULTANT_SENIORITY_RATES = {
  junior: 60,
  pleno: 85,
  senior: 110,
  master: 220
};

function getConsultantCompensationSheet_() {
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(CONSULTANT_COMPENSATION_SHEET) || ss.insertSheet(CONSULTANT_COMPENSATION_SHEET);
  ensureHeaders_(sheet, [
    'consultant_key',
    'consultant_name',
    'seniority',
    'daily_value',
    'updated_at',
    'updated_by'
  ]);
  return sheet;
}

function normalizeConsultantSeniority_(value) {
  var key = normalizeKey_(value).toLowerCase();
  if (key === 'junior' || key === 'pleno' || key === 'senior' || key === 'master') return key;
  return '';
}

function getConsultantCompensation_(params) {
  requireUser_(params || {});
  return getConsultantCompensationData_();
}

function getConsultantCompensationData_() {
  var sheet = getConsultantCompensationSheet_();
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var byKey = {};
  var consultants = values.length > 1 ? values.slice(1).map(function(row) {
    var item = rowToObject_(header, row);
    var seniority = normalizeConsultantSeniority_(item.seniority);
    return {
      consultant_key: sanitizeText_(item.consultant_key),
      consultant_name: sanitizeText_(item.consultant_name),
      seniority: seniority,
      daily_value: CONSULTANT_SENIORITY_RATES[seniority] || 0,
      updated_at: item.updated_at instanceof Date ? item.updated_at.toISOString() : String(item.updated_at || ''),
      updated_by: sanitizeText_(item.updated_by)
    };
  }).filter(function(item) {
    return !!item.consultant_key;
  }) : [];
  consultants.forEach(function(item) {
    byKey[normalizeKey_(item.consultant_key || item.consultant_name)] = item;
  });
  var usersSheet = getUsersSheet_();
  var userValues = usersSheet.getDataRange().getValues();
  if (userValues.length > 1) {
    var userHeader = userValues[0];
    userValues.slice(1).forEach(function(row) {
      var user = rowToObject_(userHeader, row);
      var seniority = normalizeConsultantSeniority_(user.seniority);
      var name = sanitizeText_(user.name || user.username);
      var enabled = String(user.enabled || '').toLowerCase();
      if (enabled === 'false' || enabled === '0' || enabled === 'nao' || enabled === 'não') return;
      if (!seniority || !name) return;
      var key = normalizeKey_(name);
      byKey[key] = {
        consultant_key: key,
        consultant_name: name,
        username: sanitizeText_(user.username),
        seniority: seniority,
        daily_value: CONSULTANT_SENIORITY_RATES[seniority],
        source: 'user'
      };
    });
  }
  if (!byKey.GUILHERME || !byKey.GUILHERME.seniority) {
    byKey.GUILHERME = {
      consultant_key: 'GUILHERME',
      consultant_name: 'Guilherme',
      seniority: 'master',
      daily_value: CONSULTANT_SENIORITY_RATES.master,
      source: 'default'
    };
  }
  return {
    ok: true,
    consultants: Object.keys(byKey).map(function(key) { return byKey[key]; }),
    rates: CONSULTANT_SENIORITY_RATES
  };
}

function setConsultantSeniority_(params) {
  var admin = requireAdmin_(params || {});
  var name = sanitizeText_(params.consultant_name || params.consultor);
  var key = normalizeKey_(name);
  var seniority = normalizeConsultantSeniority_(params.seniority);
  if (!name || !key) throw new Error('Consultor obrigatorio.');
  if (!seniority) throw new Error('Selecione Junior, Pleno, Senior ou Master.');
  var sheet = getConsultantCompensationSheet_();
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var rowNumber = 0;
  for (var i = 1; i < values.length; i++) {
    if (normalizeKey_(values[i][header.indexOf('consultant_key')]) === key) {
      rowNumber = i + 1;
      break;
    }
  }
  var row = [
    key,
    name,
    seniority,
    CONSULTANT_SENIORITY_RATES[seniority],
    new Date(),
    admin.name || admin.username
  ];
  if (rowNumber) sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return {
    ok: true,
    consultant: {
      consultant_key: key,
      consultant_name: name,
      seniority: seniority,
      daily_value: CONSULTANT_SENIORITY_RATES[seniority],
      updated_by: admin.name || admin.username
    }
  };
}

var BONUS_SALES_INDICATIONS_SHEET = 'BONUS_INDICACOES';
var BONUS_SALES_INDICATION_VALUES = { upgrade: 150, modulo: 50 };

function getBonusSalesIndicationsSheet_() {
  var sheet = getOrCreateSheet_(BONUS_SALES_INDICATIONS_SHEET);
  ensureHeaders_(sheet, [
    'id',
    'consultant_name',
    'client',
    'indication_type',
    'indication_date',
    'sale_date',
    'month',
    'value',
    'notes',
    'created_at',
    'created_by'
  ]);
  return sheet;
}

function normalizeBonusSalesIndicationType_(value) {
  var key = normalizeKey_(value).toLowerCase();
  if (key === 'upgrade') return 'upgrade';
  if (key === 'modulo' || key === 'módulo') return 'modulo';
  return '';
}

function bonusSalesDateText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var text = String(value || '').trim();
  var iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  return '';
}

function getBonusSalesIndications_(params) {
  requireAdmin_(params || {});
  var month = sanitizeMonth_(params.month);
  var sheet = getBonusSalesIndicationsSheet_();
  var values = sheet.getDataRange().getValues();
  var header = values[0] || [];
  var items = values.slice(1).map(function(row) {
    var item = rowToObject_(header, row);
    var indicationDate = bonusSalesDateText_(item.indication_date);
    var saleDate = bonusSalesDateText_(item.sale_date);
    var month = saleDate ? saleDate.slice(0, 7) : sanitizeMonth_(item.month);
    return {
      id: sanitizeText_(item.id),
      consultant_name: sanitizeText_(item.consultant_name),
      client: sanitizeText_(item.client),
      indication_type: normalizeBonusSalesIndicationType_(item.indication_type),
      indication_date: indicationDate,
      sale_date: saleDate,
      month: month,
      value: Number(item.value || 0),
      notes: sanitizeText_(item.notes),
      created_at: item.created_at instanceof Date ? item.created_at.toISOString() : String(item.created_at || ''),
      created_by: sanitizeText_(item.created_by)
    };
  }).filter(function(item) {
    return item.id && (!month || item.month === month);
  });
  return { ok: true, items: items, month: month || '' };
}

function saveBonusSalesIndication_(params) {
  var admin = requireAdmin_(params || {});
  var consultant = sanitizeText_(params.consultant_name || params.consultor);
  var client = sanitizeText_(params.client || params.cliente);
  var type = normalizeBonusSalesIndicationType_(params.indication_type || params.tipo);
  var indicationDate = String(params.indication_date || params.data_indicacao || '').slice(0, 10);
  var saleDate = String(params.sale_date || params.data_venda || '').slice(0, 10);
  if (!consultant || !client || !type || !/^\d{4}-\d{2}-\d{2}$/.test(indicationDate) || !/^\d{4}-\d{2}-\d{2}$/.test(saleDate)) {
    throw new Error('Informe consultor, cliente, tipo, data da indicacao e data da venda.');
  }
  var sheet = getBonusSalesIndicationsSheet_();
  var id = Utilities.getUuid();
  var month = saleDate.slice(0, 7);
  var value = BONUS_SALES_INDICATION_VALUES[type];
  var notes = sanitizeText_(params.notes || params.observacao);
  var createdAt = new Date();
  var createdBy = admin.name || admin.username;
  sheet.appendRow([
    id,
    consultant,
    client,
    type,
    indicationDate,
    saleDate,
    month,
    value,
    notes,
    createdAt,
    createdBy
  ]);
  var appendedRow = sheet.getLastRow();
  var monthColumn = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('month') + 1;
  if (monthColumn) sheet.getRange(appendedRow, monthColumn).setNumberFormat('@').setValue(month);
  return {
    ok: true,
    item: {
      id: id,
      consultant_name: consultant,
      client: client,
      indication_type: type,
      indication_date: indicationDate,
      sale_date: saleDate,
      month: month,
      value: value,
      notes: notes,
      created_at: createdAt.toISOString(),
      created_by: createdBy
    }
  };
}

function deleteBonusSalesIndication_(params) {
  requireAdmin_(params || {});
  var id = sanitizeText_(params.id);
  if (!id) throw new Error('Indicacao obrigatoria.');
  var sheet = getBonusSalesIndicationsSheet_();
  var values = sheet.getDataRange().getValues();
  var header = values[0] || [];
  var idIndex = header.indexOf('id');
  for (var i = values.length - 1; i >= 1; i--) {
    if (sanitizeText_(values[i][idIndex]) === id) {
      sheet.deleteRow(i + 1);
      return { ok: true, id: id };
    }
  }
  throw new Error('Indicacao nao encontrada.');
}

function logProjectFollowup_(params) {
  params = params || {};
  var user = requireUser_(params);
  var cliente = sanitizeText_(params.cliente);
  if (!cliente) throw new Error('cliente is required');
  var consultor = sanitizeText_(params.consultor);
  if (!canUserAccessProjectItem_(user, { consultor: consultor })) {
    throw new Error('Este usuario so pode registrar acompanhamentos da propria carteira.');
  }
  var consideracao = sanitizeText_(params.consideracao || params.observacao);
  var proximaAcao = sanitizeText_(params.proxima_acao);
  if (!consideracao && !proximaAcao) throw new Error('consideracao or proxima_acao is required');
  var sheet = getProjectFollowupSheet_();
  var dataAcompanhamento = sanitizeText_(params.data_acompanhamento) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var followupId = makeAuthToken_();
  var row = [
    new Date(),
    dataAcompanhamento,
    sanitizeMonth_(params.mes),
    cliente,
    consultor,
    sanitizeText_(params.project_key),
    sanitizeText_(params.link_projeto),
    consideracao,
    proximaAcao,
    user.name || user.username || sanitizeText_(params.responsavel),
    sanitizeText_(params.origem) || 'site',
    sanitizeText_(params.status),
    sanitizeText_(params.url),
    sanitizeText_(params.navegador),
    sanitizeText_(params.kanban_stage || params.etapa_kanban),
    followupId,
    normalizeFollowupStatus_(params.followup_status || params.observation_status || 'pendente'),
    '',
    '',
    ''
  ];
  sheet.appendRow(row);
  var rowNumber = sheet.getLastRow();
  var followup = rowToObject_(getProjectFollowupHeaders_(), row);
  followup.row_number = rowNumber;
  return {
    ok: true,
    logged: true,
    followup: followup,
    total_rows: Math.max(0, rowNumber - 1)
  };
}

function getProjectFollowupHeaders_() {
  return [
    'logged_at',
    'data_acompanhamento',
    'mes',
    'cliente',
    'consultor',
    'project_key',
    'link_projeto',
    'consideracao',
    'proxima_acao',
    'responsavel',
    'origem',
    'status',
    'url',
    'navegador',
    'kanban_stage',
    'followup_id',
    'followup_status',
    'status_updated_at',
    'status_updated_by',
    'status_comment'
  ];
}

function normalizeFollowupStatus_(value) {
  var key = normalizeKey_(value);
  if (key === 'CONCLUIDO' || key === 'CONCLUIDA' || key === 'DONE' || key === 'RESOLVIDO' || key === 'FINALIZADO') return 'concluido';
  if (key === 'PARCIAL' || key === 'PARCIALMENTE' || key.indexOf('PARCIAL') >= 0) return 'parcial';
  return 'pendente';
}

function publicProjectFollowup_(item, rowNumber) {
  item = item || {};
  var updatedAt = item.status_updated_at instanceof Date ? item.status_updated_at.toISOString() : String(item.status_updated_at || '');
  return {
    row_number: rowNumber || '',
    logged_at: item.logged_at instanceof Date ? item.logged_at.toISOString() : String(item.logged_at || ''),
    data_acompanhamento: item.data_acompanhamento instanceof Date ? Utilities.formatDate(item.data_acompanhamento, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(item.data_acompanhamento || ''),
    mes: String(item.mes || ''),
    cliente: String(item.cliente || ''),
    consultor: String(item.consultor || ''),
    project_key: String(item.project_key || ''),
    link_projeto: String(item.link_projeto || ''),
    consideracao: String(item.consideracao || ''),
    proxima_acao: String(item.proxima_acao || ''),
    responsavel: String(item.responsavel || ''),
    kanban_stage: String(item.kanban_stage || item.etapa_kanban || ''),
    origem: String(item.origem || ''),
    status: String(item.status || ''),
    url: String(item.url || ''),
    navegador: String(item.navegador || ''),
    followup_id: String(item.followup_id || ''),
    followup_status: normalizeFollowupStatus_(item.followup_status),
    status_updated_at: updatedAt,
    status_updated_by: String(item.status_updated_by || ''),
    status_comment: String(item.status_comment || '')
  };
}

function getProjectFollowups_(params, limit) {
  var user = requireUser_(params);
  var sheet = getProjectFollowupSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { ok: true, followups: [], total: 0 };
  }
  var header = values[0];
  var rows = values.slice(1);
  var max = Math.max(1, Math.min(Number(limit || 1000), 5000));
  var visibleProjectKeys = {};
  var visibleEntries = rows.map(function(row, index) {
    var item = rowToObject_(header, row);
    return {
      item: item,
      row_number: index + 2
    };
  }).filter(function(entry) {
    return canUserAccessProjectItem_(user, entry.item);
  });
  var start = Math.max(0, visibleEntries.length - max);
  var followups = visibleEntries.slice(start).map(function(entry) {
    var projectKey = sanitizeText_(entry.item.project_key).toUpperCase();
    if (projectKey) visibleProjectKeys[projectKey] = true;
    return publicProjectFollowup_(entry.item, entry.row_number);
  }).reverse();
  return {
    ok: true,
    total: visibleEntries.length,
    followups: followups,
    kanban_states: getProjectKanbanStates_(userHasFullProjectAccess_(user) ? null : visibleProjectKeys)
  };
}

function findProjectFollowupRow_(sheet, followupId, rowNumber) {
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return null;
  var header = values[0];
  var rowIndex = rowNumber && rowNumber >= 2 && rowNumber <= values.length ? rowNumber - 1 : null;
  if (rowIndex !== null) {
    var byRow = rowToObject_(header, values[rowIndex]);
    if (!followupId || String(byRow.followup_id || '') === followupId) {
      return { sheet: sheet, row: rowIndex + 1, header: header, item: byRow };
    }
  }
  if (!followupId) return null;
  for (var i = values.length - 1; i >= 1; i--) {
    var item = rowToObject_(header, values[i]);
    if (String(item.followup_id || '') === followupId) {
      return { sheet: sheet, row: i + 1, header: header, item: item };
    }
  }
  return null;
}

function setProjectFollowupStatus_(params) {
  params = params || {};
  var user = requireUser_(params);
  var followupId = sanitizeText_(params.followup_id);
  var rowNumber = toInt_(params.row_number, null);
  if (!followupId && !rowNumber) throw new Error('followup_id or row_number is required');
  var sheet = getProjectFollowupSheet_();
  var found = findProjectFollowupRow_(sheet, followupId, rowNumber);
  if (!found) throw new Error('Acompanhamento nao encontrado.');
  if (!canUserAccessProjectItem_(user, found.item)) {
    throw new Error('Este usuario so pode atualizar acompanhamentos da propria carteira.');
  }
  var status = normalizeFollowupStatus_(params.followup_status || params.status);
  var comment = sanitizeText_(params.status_comment || params.comment || params.comentario);
  var updatedAt = new Date();
  var updatedBy = user.name || user.username || '';
  var header = found.header;
  var statusCol = header.indexOf('followup_status') + 1;
  var updatedAtCol = header.indexOf('status_updated_at') + 1;
  var updatedByCol = header.indexOf('status_updated_by') + 1;
  var commentCol = header.indexOf('status_comment') + 1;
  if (!statusCol || !updatedAtCol || !updatedByCol || !commentCol) throw new Error('Colunas de status do acompanhamento nao encontradas.');
  sheet.getRange(found.row, statusCol).setValue(status);
  sheet.getRange(found.row, updatedAtCol).setValue(updatedAt);
  sheet.getRange(found.row, updatedByCol).setValue(updatedBy);
  sheet.getRange(found.row, commentCol).setValue(comment);
  SpreadsheetApp.flush();
  found.item.followup_status = status;
  found.item.status_updated_at = updatedAt;
  found.item.status_updated_by = updatedBy;
  found.item.status_comment = comment;
  return {
    ok: true,
    saved: true,
    followup: publicProjectFollowup_(found.item, found.row)
  };
}

function normalizeKanbanStage_(value) {
  var raw = String(value || '').toLowerCase();
  if (raw === 'done' || raw.indexOf('conclu') >= 0) return 'done';
  if (raw === 'doing' || raw.indexOf('interfer') >= 0 || raw.indexOf('processo') >= 0 || raw.indexOf('andamento') >= 0) return 'doing';
  return 'pending';
}

function getProjectKanbanStates_(allowedProjectKeys) {
  var sheet = getProjectKanbanStateSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return {};
  var header = values[0];
  var map = {};
  values.slice(1).forEach(function(row) {
    var item = rowToObject_(header, row);
    var cycleKey = sanitizeText_(item.cycle_key);
    var projectKey = sanitizeText_(item.project_key).toUpperCase();
    if (!cycleKey || !projectKey) return;
    if (allowedProjectKeys && !allowedProjectKeys[projectKey]) return;
    map[cycleKey + '|' + projectKey] = normalizeKanbanStage_(item.stage);
  });
  return map;
}

function findProjectKanbanStateRow_(sheet, cycleKey, projectKey) {
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return null;
  var header = values[0];
  for (var i = 1; i < values.length; i++) {
    var item = rowToObject_(header, values[i]);
    if (String(item.cycle_key || '') === cycleKey && String(item.project_key || '').toUpperCase() === projectKey) {
      return { row: i + 1, header: header };
    }
  }
  return null;
}

function setProjectKanbanStage_(params) {
  params = params || {};
  var user = requireUser_(params);
  var projectKey = sanitizeText_(params.project_key).toUpperCase();
  var cycleKey = sanitizeText_(params.cycle_key);
  if (!projectKey || !cycleKey) throw new Error('project_key and cycle_key are required');
  if (!canUserAccessProjectItem_(user, { consultor: sanitizeText_(params.consultor) })) {
    throw new Error('Este usuario so pode alterar o Kanban da propria carteira.');
  }
  var stage = normalizeKanbanStage_(params.stage);
  var sheet = getProjectKanbanStateSheet_();
  var found = findProjectKanbanStateRow_(sheet, cycleKey, projectKey);
  var row = [
    new Date(),
    cycleKey,
    projectKey,
    stage,
    user.name || user.username || ''
  ];
  if (found) {
    sheet.getRange(found.row, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  var allowedProjectKeys = null;
  if (!userHasFullProjectAccess_(user)) {
    allowedProjectKeys = {};
    allowedProjectKeys[projectKey] = true;
  }
  return {
    ok: true,
    saved: true,
    cycle_key: cycleKey,
    project_key: projectKey,
    stage: stage,
    kanban_states: getProjectKanbanStates_(allowedProjectKeys)
  };
}

function clearProjectKanbanStage_(cycleKey, projectKey) {
  cycleKey = sanitizeText_(cycleKey);
  projectKey = sanitizeText_(projectKey).toUpperCase();
  if (!cycleKey || !projectKey) return false;
  var sheet = getProjectKanbanStateSheet_();
  var found = findProjectKanbanStateRow_(sheet, cycleKey, projectKey);
  if (!found) return false;
  sheet.deleteRow(found.row);
  return true;
}

function normalizeFollowupDateForCompare_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').slice(0, 10);
}

function normalizeFollowupLoggedForCompare_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value || '');
}

function deleteProjectFollowup_(params) {
  params = params || {};
  requireAdmin_(params);
  var projectKey = sanitizeText_(params.project_key).toUpperCase();
  var cliente = sanitizeText_(params.cliente).toLowerCase();
  var loggedAt = sanitizeText_(params.logged_at);
  var dataAcompanhamento = sanitizeText_(params.data_acompanhamento);
  var followupId = sanitizeText_(params.followup_id);
  var rowNumber = toInt_(params.row_number, null);
  var cycleKey = sanitizeText_(params.cycle_key);
  if (!projectKey && !cliente) throw new Error('project_key or cliente is required');
  var sheet = getProjectFollowupSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, deleted: false, message: 'Nenhum acompanhamento encontrado.' };
  var header = values[0];
  var deleted = null;
  var start = rowNumber && rowNumber >= 2 && rowNumber <= values.length ? rowNumber - 1 : values.length - 1;
  var end = rowNumber ? start : 1;
  for (var i = start; i >= end; i--) {
    var item = rowToObject_(header, values[i]);
    if (followupId && String(item.followup_id || '') !== followupId) continue;
    var itemKey = String(item.project_key || '').toUpperCase();
    var itemCliente = String(item.cliente || '').toLowerCase();
    var keyMatches = projectKey ? itemKey === projectKey : itemCliente === cliente;
    if (!keyMatches) continue;
    if (loggedAt && normalizeFollowupLoggedForCompare_(item.logged_at) !== loggedAt) continue;
    if (dataAcompanhamento && normalizeFollowupDateForCompare_(item.data_acompanhamento) !== dataAcompanhamento) continue;
    deleted = rowToObject_(header, values[i]);
    sheet.deleteRow(i + 1);
    break;
  }
  if (deleted && cycleKey) clearProjectKanbanStage_(cycleKey, projectKey);
  return {
    ok: true,
    deleted: !!deleted,
    followup: deleted || null,
    total_rows: Math.max(0, sheet.getLastRow() - 1),
    kanban_states: getProjectKanbanStates_()
  };
}

function ensureHeaders_(sheet, headers) {
  headers = Array.isArray(headers) ? headers.filter(function(header) { return sanitizeText_(header); }) : [];
  if (!headers.length) throw new Error('Cabecalho da planilha nao informado.');
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  var current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  headers.forEach(function(header, index) {
    if (current[index] !== header) sheet.getRange(1, index + 1).setValue(header);
  });
}

function writeObjectsToSheet_(sheet, objects, headers) {
  headers = Array.isArray(headers) ? headers.filter(function(header) { return sanitizeText_(header); }) : [];
  if (!headers.length) throw new Error('Cabecalho da planilha nao informado.');
  sheet.clearContents();
  ensureHeaders_(sheet, headers);
  if (!objects.length) return;
  var rows = objects.map(function(item) {
    return headers.map(function(header) { return item[header] || ''; });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function findConfigHeaderRow_(values) {
  var candidates = getConfigHeaderCandidates_(values);
  return candidates.length ? candidates[0].index : 0;
}

function getConfigHeaderCandidates_(values) {
  var candidates = [];
  var bestIndex = 0;
  var bestScore = -1;
  var maxRows = Math.min(values.length, 15);
  for (var i = 0; i < maxRows; i++) {
    var normalized = normalizeConfigHeader_(values[i] || []);
    var score = 0;
    if (normalized.indexOf('mes') >= 0) score += 3;
    if (normalized.indexOf('cliente') >= 0) score += 3;
    if (normalized.indexOf('project_url') >= 0) score += 2;
    if (normalized.indexOf('view_id') >= 0) score += 2;
    if (normalized.indexOf('list_id') >= 0) score += 2;
    if (normalized.indexOf('folder_id') >= 0) score += 2;
    if (normalized.indexOf('space_id') >= 0) score += 2;
    if (normalized.indexOf('enabled') >= 0) score += 1;
    if (score >= 5) {
      candidates.push({ index: i, score: score, header: normalized });
    }
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  if (!candidates.length) candidates.push({ index: bestIndex, score: bestScore, header: normalizeConfigHeader_(values[bestIndex] || []) });
  candidates.sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return a.index - b.index;
  });
  return candidates;
}

function normalizeConfigHeader_(header) {
  var used = {};
  return (header || []).map(function(name) {
    var canonical = canonicalConfigHeader_(name);
    if (!canonical) canonical = sanitizeText_(name);
    if (!canonical) return '';
    if (!used[canonical]) {
      used[canonical] = 1;
      return canonical;
    }
    used[canonical] += 1;
    return canonical + '_' + used[canonical];
  });
}

function canonicalConfigHeader_(name) {
  var key = normalizeKey_(name);
  var aliases = {
    enabled: ['ENABLED', 'HABILITADO', 'ATIVO', 'SINCRONIZAR', 'SYNC', 'CLICKUP SYNC'],
    mes: ['MES', 'MONTH', 'MES DA VENDA', 'MES VENDA'],
    cliente: ['CLIENTE', 'PROJETO', 'CLIENTE PROJETO', 'NOME DO PROJETO', 'NOME PROJETO', 'RAZAO SOCIAL'],
    project_key: ['PROJECT KEY', 'CHAVE', 'CHAVE PROJETO', 'PROJECT KEY CLICKUP'],
    project_url: ['PROJECT URL', 'PROJECT URL CLICKUP', 'LINK PROJETO', 'LINK DO PROJETO', 'LINK CLICKUP', 'URL CLICKUP', 'PROJETO LINK'],
    view_id: ['VIEW ID', 'VIEWID', 'ID VIEW', 'CLICKUP VIEW ID'],
    list_id: ['LIST ID', 'LISTID', 'ID LISTA', 'CLICKUP LIST ID'],
    folder_id: ['FOLDER ID', 'FOLDERID', 'ID FOLDER', 'ID PASTA', 'PASTA ID', 'CLICKUP FOLDER ID'],
    space_id: ['SPACE ID', 'SPACEID', 'ID SPACE', 'ID ESPACO', 'ESPACO ID', 'CLICKUP SPACE ID'],
    sync_mode: ['SYNC MODE', 'MODO SYNC', 'MODO SINCRONIZACAO'],
    notes: ['NOTES', 'OBS', 'OBSERVACAO', 'NOTAS']
  };
  var found = '';
  Object.keys(aliases).some(function(canonical) {
    if (aliases[canonical].indexOf(key) >= 0) {
      found = canonical;
      return true;
    }
    return false;
  });
  return found;
}

function rowToObject_(header, row) {
  var out = {};
  header.forEach(function(key, index) {
    out[String(key || '').trim()] = row[index];
  });
  return out;
}

function jsonOutput_(payload, callbackName) {
  var json = JSON.stringify(payload);
  if (callbackName) {
    return ContentService.createTextOutput(String(callbackName) + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function simplifyErrorMessage_(error) {
  var message = error && error.message ? error.message : String(error || '');
  if (/list_id invalido|view_id invalido|folder_id invalido|space_id invalido/i.test(message)) {
    return 'O CLICKUP_CONFIG possui list_id/view_id/folder_id/space_id vazio, invalido ou mal formatado.';
  }
  if (/largura de banda|bandwidth/i.test(message)) {
    return 'ClickUp recusou a sincronizacao por limite temporario de banda. Aguarde alguns minutos e tente novamente.';
  }
  if (/ClickUp API error 401/i.test(message)) {
    return 'CLICKUP_TOKEN invalido ou expirado nas propriedades do Apps Script.';
  }
  if (/ClickUp API error 403/i.test(message)) {
    if (/auditlogs|audit logs|Enterprise Plan|Enterprise/i.test(message)) {
      return 'O Audit Log do ClickUp exige Workspace Enterprise e permissao de owner/admin com acesso aos logs.';
    }
    return 'O token nao tem permissao para acessar esta lista/view do ClickUp.';
  }
  if (/ClickUp API error 404/i.test(message)) {
    return 'A lista/view/folder/space informado no CLICKUP_CONFIG nao foi encontrado no ClickUp. Confira list_id/view_id/folder_id/space_id.';
  }
  if (/ClickUp API error 400/i.test(message)) {
    return 'O ClickUp rejeitou a requisicao. Normalmente isso indica id ClickUp invalido ou configuracao incompleta.';
  }
  if (/ClickUp API error 5\d\d/i.test(message)) {
    return 'O ClickUp retornou erro temporario no servidor. Tente sincronizar novamente em alguns minutos.';
  }
  if (/Project mapping not found/i.test(message)) {
    return 'Projeto nao encontrado no CLICKUP_CONFIG para a chave informada.';
  }
  if (/Missing CLICKUP_TOKEN/i.test(message)) {
    return 'CLICKUP_TOKEN nao configurado nas propriedades do Apps Script.';
  }
  return message || 'Erro inesperado na sincronizacao.';
}

function isClickUpBandwidthError_(error) {
  var message = error && error.message ? error.message : String(error || '');
  return /largura de banda|bandwidth/i.test(message);
}

function isClickUpRecoverableSyncError_(error) {
  var message = error && error.message ? error.message : String(error || '');
  return isClickUpBandwidthError_(error) ||
    /ClickUp API error (400|403|404|408|429|5\d\d)/i.test(message) ||
    /list_id invalido|view_id invalido|folder_id invalido|space_id invalido/i.test(message);
}

function appendQueryParam_(url, key, value) {
  if (!value) return url;
  var sep = String(url).indexOf('?') >= 0 ? '&' : '?';
  return String(url) + sep + encodeURIComponent(key) + '=' + encodeURIComponent(value);
}

function getScriptProperty_(key, fallback) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value !== null && value !== undefined && value !== '' ? value : fallback;
}

function sanitizeText_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function normalizeClickUpId_(value) {
  var text = sanitizeText_(value);
  if (!text) return '';

  if (/^\d+$/.test(text)) return text;

  var numeric = Number(String(text).replace(',', '.'));
  if (isFinite(numeric) && numeric > 0) {
    var rounded = String(Math.round(numeric));
    if (/^\d+$/.test(rounded)) return rounded;
  }

  if (/^[0-9 .,'+\-Ee]+$/.test(text)) {
    var digits = text.replace(/\D/g, '');
    if (digits) return digits;
  }

  return text;
}

function normalizeClickUpNumericId_(value) {
  var id = normalizeClickUpId_(value);
  return /^\d+$/.test(id) ? id : '';
}

function extractClickUpIdFromUrl_(url, kind) {
  var text = sanitizeText_(url);
  if (!text) return '';
  var listMatch = text.match(/\/list\/(\d+)/i) || text.match(/[?&]list_id=(\d+)/i);
  if (kind === 'list') return listMatch ? normalizeClickUpId_(listMatch[1]) : '';
  var folderMatch =
    text.match(/\/v\/o\/f\/(\d+)/i) ||
    text.match(/\/v\/l\/f\/(\d+)/i) ||
    text.match(/\/v\/f\/(\d+)/i) ||
    text.match(/\/folder\/(\d+)/i) ||
    text.match(/[?&]folder_id=(\d+)/i);
  if (kind === 'folder') return folderMatch ? normalizeClickUpId_(folderMatch[1]) : '';
  var spaceMatch =
    text.match(/\/v\/s\/(\d+)/i) ||
    text.match(/\/space\/(\d+)/i) ||
    text.match(/[?&]space_id=(\d+)/i);
  if (kind === 'space') return spaceMatch ? normalizeClickUpId_(spaceMatch[1]) : '';
  var viewMatch =
    text.match(/\/v\/l\/(?!f\/)([^/?#]+)/i) ||
    text.match(/\/v\/o\/(?!f\/)([^/?#]+)/i) ||
    text.match(/[?&]view_id=([^&#]+)/i);
  if (kind === 'view') return viewMatch ? normalizeClickUpId_(decodeURIComponent(viewMatch[1])) : '';
  return '';
}

function sanitizeMonth_(value) {
  return String(value || '').toUpperCase().trim();
}

function extractLeadingNumber_(value) {
  var text = sanitizeText_(value);
  var match = text.match(/^(\d+)/) || text.match(/^fase\s*(\d+(?:\.\d+)?)/i);
  return match ? Number(match[1]) : 9999;
}

function normalizeKey_(value) {
  var key = sanitizeText_(value)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (key === 'WAND' || key.indexOf('WANDERLEY') >= 0 || key.indexOf('WANDERLEI') >= 0) return 'WANDERLEY CABRAL';
  return key;
}

function buildProjectKey_(month, cliente) {
  return sanitizeMonth_(month) + '|' + normalizeKey_(cliente);
}

function normalizeProjectKey_(value) {
  var text = sanitizeText_(value);
  if (!text) return '';
  var parts = text.split('|');
  if (parts.length < 2) return normalizeKey_(text);
  var mes = sanitizeMonth_(parts.shift());
  var cliente = normalizeKey_(parts.join('|'));
  return mes + '|' + cliente;
}

function normalizeBoolean_(value, fallback) {
  if (value === true || value === false) return value;
  if (value === null || value === undefined || value === '') return fallback;
  var norm = String(value).toLowerCase().trim();
  return ['1', 'true', 'sim', 'yes', 'y'].indexOf(norm) >= 0;
}

function toInt_(value, fallback) {
  var n = Number(value);
  return isNaN(n) ? fallback : n;
}

function fromMillisIso_(value) {
  var n = Number(value);
  if (!n) return '';
  return new Date(n).toISOString();
}

function diffDaysFromIso_(iso) {
  if (!iso) return '';
  var date = new Date(iso);
  if (isNaN(date.getTime())) return '';
  var now = new Date();
  var start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var end = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var diff = end.getTime() - start.getTime();
  return Math.max(0, Math.floor(diff / 86400000));
}

function dedupeTasks_(tasks) {
  var seen = {};
  return tasks.filter(function(task) {
    var id = String(task.id || '');
    if (!id || seen[id]) return false;
    seen[id] = true;
    return true;
  });
}

function isClosedStatus_(status) {
  var norm = sanitizeText_(status).toLowerCase();
  return /(done|closed|complete|conclu|finaliz|feito|resolved|encerrad|sucesso|finalizado|aprovado gest)/.test(norm);
}

function isMilestoneTask_(task) {
  if (task && task._confirmed_milestone === true) return true;
  var typeName = sanitizeText_(
    task.custom_item && task.custom_item.name ||
    task.custom_task_type && task.custom_task_type.name ||
    task.task_type ||
    task.type
  ).toLowerCase();
  if (String(task.custom_item_id || '') === '1') return true;
  if (typeName.indexOf('milestone') >= 0 || typeName.indexOf('marco') >= 0) return true;
  return false;
}

var CLICKUP_CUSTOM_ITEM_TYPES_MEMORY_ = null;
function fetchClickUpCustomItemTypes_() {
  if (CLICKUP_CUSTOM_ITEM_TYPES_MEMORY_) return CLICKUP_CUSTOM_ITEM_TYPES_MEMORY_;
  var types = {};
  var workspaceId = getClickUpWorkspaceId_();
  if (!workspaceId) return types;
  try {
    var response = clickupRequest_('get', '/team/' + workspaceId + '/custom_item');
    var items = Array.isArray(response)
      ? response
      : (response.custom_items || response.custom_task_types || response.items || []);
    (items || []).forEach(function(item) {
      var id = String(item && (item.id || item.custom_item_id) || '').trim();
      var name = sanitizeText_(item && (item.name || item.label || item.type));
      if (id && name) types[id] = name;
    });
  } catch (error) {
    // Mantem a sincronizacao atual funcionando quando o workspace nao libera este endpoint.
  }
  CLICKUP_CUSTOM_ITEM_TYPES_MEMORY_ = types;
  return types;
}

function enrichClickUpCustomItemNames_(tasks, types) {
  types = types || {};
  return (tasks || []).map(function(task) {
    var id = String(task && task.custom_item_id || '').trim();
    if (task && !task.custom_item_name && id && types[id]) task.custom_item_name = types[id];
    return task;
  });
}

function clickUpTaskCustomItemName_(task) {
  return sanitizeText_(
    task && (
      task.custom_item_name ||
      task.item_tipo_clickup ||
      task.custom_item && task.custom_item.name ||
      task.custom_item && task.custom_item.label ||
      task.custom_item && task.custom_item.type ||
      task.custom_item_type && task.custom_item_type.name ||
      task.custom_type && task.custom_type.name ||
      task.custom_task_type && task.custom_task_type.name ||
      task.task_type ||
      task.type
    )
  );
}

function isClickUpProjectDeliveryCustomItemId_(task) {
  var id = String(task && task.custom_item_id || '').trim();
  return id && CLICKUP_PROJECT_DELIVERY_CUSTOM_ITEM_IDS.indexOf(id) >= 0;
}

function isProjectDeliveryTask_(task) {
  var key = normalizeKey_(
    (task && (task.custom_item_name || task.item_tipo_clickup)) ||
    clickUpTaskCustomItemName_(task)
  );
  return isClickUpProjectDeliveryCustomItemId_(task) ||
    key === 'ENTREGA' ||
    key === 'ENTREGA DE PROJETO';
}

function isProjectDeliveryClosingCandidate_(task) {
  return isProjectDeliveryTask_(task) || String(task && task.marcador_entrega || '').toLowerCase() === 'sim';
}

function isProjectBreakOffText_(value) {
  var key = normalizeKey_(value);
  return key.indexOf('FASE 8') >= 0 || key.indexOf('BREAK OFF') >= 0;
}

function isProjectClosingDeliveryItem_(item, phaseName, phase) {
  var isBreakOff = isProjectBreakOffText_(phaseName) ||
    isProjectBreakOffText_(item && (item.fase_nome || item.fase || item.phase)) ||
    isProjectBreakOffText_(item && (item.name || item.nome));
  var status = sanitizeText_(item && (item.status_original || item.status && (item.status.status || item.status.type || item.status.label)));
  return isBreakOff && isProjectDeliveryClosingCandidate_(item) &&
    hasProjectClosingApprovalSignalByTaskOrPhase_(item, status, phase);
}

function isClosingTrackedTask_(task) {
  return isMilestoneTask_(task) || isProjectDeliveryTask_(task);
}

function isPhaseTask_(task, byId) {
  var name = sanitizeText_(task.name).toLowerCase();
  if (String(task.parent || '').trim()) return false;
  if (/^fase\s*(\d+(?:\.\d+)?|x)\b/.test(name)) return true;
  if (/^fase\s*final/.test(name)) return true;
  return false;
}

function countDirectChildren_(task, byId) {
  var taskId = String(task.id || '');
  var count = 0;
  Object.keys(byId).forEach(function(id) {
    if (String(byId[id].parent || '') === taskId) count += 1;
  });
  return count;
}

function resolvePhaseForTask_(task, byId, phaseMap) {
  var current = task;
  var visited = {};
  var depth = 0;
  while (current && current.parent && !visited[current.parent]) {
    visited[current.parent] = true;
    depth += 1;
    if (phaseMap[String(current.parent)]) {
      return {
        phase: phaseMap[String(current.parent)],
        depth: depth
      };
    }
    current = byId[String(current.parent)];
  }

  if (task.list && task.list.name) {
    var listName = sanitizeText_(task.list.name);
    var phaseByName = Object.keys(phaseMap).map(function(id) { return phaseMap[id]; }).filter(function(phase) {
      return phase.nome === listName;
    })[0];
    if (phaseByName) {
      return {
        phase: phaseByName,
        depth: 0
      };
    }
  }
  return {
    phase: null,
    depth: 0
  };
}

function shouldCountTaskInSummary_(task, phaseInfo, byId) {
  if (phaseInfo.phase) {
    if (countDirectChildren_(task, byId) > 0) {
      return false;
    }
    return true;
  }
  if (task.parent) return false;
  return countDirectChildren_(task, byId) === 0;
}

function isRecoverableClickUpHttpError_(code, text) {
  var message = String(text || '');
  return code === 408 ||
    code === 409 ||
    code === 423 ||
    code === 425 ||
    code === 429 ||
    (code >= 500 && code < 600) ||
    /largura de banda|bandwidth/i.test(message);
}

function isRecoverableClickUpTransportError_(error) {
  var message = error && error.message ? error.message : String(error || '');
  return /timeout|timed out|temporar|temporari|service unavailable|bandwidth|largura de banda|internal error|exception/i.test(message);
}

function getLatestUpdate_(tasks) {
  var latest = 0;
  tasks.forEach(function(task) {
    var updated = Number(task.date_updated || 0);
    if (updated > latest) latest = updated;
  });
  return latest ? new Date(latest).toISOString() : '';
}

function getLatestUpdateItem_(items) {
  var latest = null;
  (items || []).forEach(function(item) {
    var date = item && item.updated_at ? new Date(item.updated_at) : null;
    if (!date || isNaN(date.getTime())) return;
    if (!latest || date.getTime() > latest.ts) {
      latest = {
        ts: date.getTime(),
        item: {
          id: String(item.id || ''),
          tipo: item.tipo || 'task',
          nome: sanitizeText_(item.nome || ''),
          fase_nome: sanitizeText_(item.fase_nome || ''),
          status_original: sanitizeText_(item.status_original || ''),
          updated_at: date.toISOString()
        }
      };
    }
  });
  return latest ? latest.item : null;
}

function getLatestUpdateItemFromRawTasks_(tasks, byId, phaseMap) {
  var latest = null;
  (tasks || []).forEach(function(task) {
    var updated = Number(task && task.date_updated || 0);
    if (!updated) return;
    if (!latest || updated > latest.ts) {
      var phaseInfo = resolvePhaseForTask_(task, byId || {}, phaseMap || {});
      var status = task.status && (task.status.status || task.status.type || task.status.label) || '';
      latest = {
        ts: updated,
        item: {
          id: String(task.id || ''),
          tipo: isMilestoneTask_(task) ? 'marco' : (isPhaseTask_(task, byId || {}) ? 'fase' : 'task'),
          nome: sanitizeText_(task.name || ''),
          fase_nome: phaseInfo && phaseInfo.phase ? sanitizeText_(phaseInfo.phase.nome || '') : '',
          status_original: sanitizeText_(status),
          updated_at: new Date(updated).toISOString()
        }
      };
    }
  });
  return latest ? latest.item : null;
}

function sincronizarDiariasCmaxMesAtual() {
  return syncCmaxDailyEvents_({
    month: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM')
  });
}

function sincronizarHistoricoDiariasCmax() {
  return startCmaxDailyHistoryBackground_({});
}

function reconstruirVisaoDiariasCmax() {
  try {
    var sheet = getCmaxDailySheet_();
    var headers = getCmaxDailyHeaders_();
    var values = sheet.getDataRange().getValues();
    rebuildCmaxDailyMaterializedView_(values.slice(1), headers);
    return { ok: true, rows: Math.max(0, values.length - 1), sheet: CMAX_DAILY_VIEW_SHEET };
  } finally {
    PropertiesService.getScriptProperties().deleteProperty('CMAX_VIEW_BUILD_SCHEDULED');
  }
}

var CMAX_DAILY_SHEET = 'CMAX_DIARIAS';
var CMAX_DAILY_VIEW_SHEET = 'CMAX_DIARIAS_VIEW';
var CMAX_AGENDA_URL = 'https://www.multbovinos.com/servicos/eventoscontato/obtenha-agenda/?format=json';
var CMAX_TOKEN_AUTH_URL = 'https://www.multbovinos.com/servicos/login/';
var CMAX_HISTORY_DEFAULT_START_MONTH = '2024-01';
var CMAX_VIEW_CACHE_SECONDS = 21600;

function getCmaxDailyHeaders_() {
  return [
    'event_key', 'event_id', 'data', 'mes', 'ano', 'consultor', 'cliente',
    'tipo', 'resultado', 'descricao', 'hora_inicio', 'hora_fim',
    'sincronizado_em', 'raw_json'
  ];
}

function getCmaxDailyViewHeaders_() {
  return [
    'event_key', 'event_id', 'data', 'mes', 'ano', 'consultor', 'cliente',
    'tipo', 'resultado', 'descricao', 'hora_inicio', 'hora_fim',
    'sincronizado_em', 'modalidade', 'contabiliza_diaria'
  ];
}

function getCmaxDailySheet_() {
  var sheet = getOrCreateSheet_(CMAX_DAILY_SHEET);
  ensureHeaders_(sheet, getCmaxDailyHeaders_());
  return sheet;
}

function getCmaxDailyEvents_(params) {
  params = params || {};
  var props = PropertiesService.getScriptProperties();
  var month = sanitizeCmaxMonth_(params.month || params.mes);
  var consultant = sanitizeText_(params.consultant || params.consultor).toUpperCase();
  var cached = readCmaxMaterializedCache_(month, consultant);
  if (cached) {
    cached.cached = true;
    cached.history_sync = getCmaxDailyHistoryStatus_().history_sync;
    cached.consultant_compensation = getConsultantCompensationData_().consultants;
    cached.consultant_daily_rates = CONSULTANT_SENIORITY_RATES;
    return cached;
  }
  var meta;
  try { meta = JSON.parse(props.getProperty('CMAX_VIEW_META_JSON') || 'null'); } catch (ignored) { meta = null; }
  if (!meta || !meta.ranges) {
    scheduleCmaxDailyViewBuild_();
    return {
      ok: true,
      events: [],
      total: 0,
      month: month,
      history_months: cmaxHistoryMonths_(),
      history_loaded_months: [],
      consultant_compensation: getConsultantCompensationData_().consultants,
      consultant_daily_rates: CONSULTANT_SENIORITY_RATES,
      history_sync: getCmaxDailyHistoryStatus_().history_sync,
      building_view: true,
      message: 'Preparando visão rápida CMAX em segundo plano.'
    };
  }
  var events = [];
  var snapshotEvents = readCmaxDailyMonthSnapshot_(month, consultant, meta);
  if (snapshotEvents) {
    events = snapshotEvents;
  }
  var rangeInfo = month ? meta.ranges[month] : meta.all;
  if (!snapshotEvents && rangeInfo && rangeInfo.count > 0) {
    var sheet = getOrCreateSheet_(CMAX_DAILY_VIEW_SHEET);
    var headers = getCmaxDailyViewHeaders_();
    var values = sheet.getRange(rangeInfo.start, 1, rangeInfo.count, headers.length).getValues();
    events = values.map(function(row) {
      var item = {};
      headers.forEach(function(header, index) { item[header] = row[index]; });
      item.contabiliza_diaria = item.contabiliza_diaria === true || String(item.contabiliza_diaria).toLowerCase() === 'true';
      return item;
    }).filter(function(item) {
      return !consultant || sanitizeText_(item.consultor).toUpperCase() === consultant;
    });
  }
  var result = {
    ok: true,
    events: events,
    total: events.length,
    month: month,
      available_months: cmaxRelevantMonths_(meta.available_months),
    history_months: cmaxHistoryMonths_(),
      history_loaded_months: cmaxSnapshotMonths_(meta),
    training_team_cutoff: meta.training_team_cutoff || '',
    training_team: meta.training_team || [],
    available_activities: meta.available_activities || [],
    history_sync: getCmaxDailyHistoryStatus_().history_sync,
    synced_at: meta.synced_at || '',
    materialized: true,
    sheet: CMAX_DAILY_VIEW_SHEET
  };
  writeCompressedScriptCache_(cmaxViewCacheKey_({ month: month, consultant: consultant }), result, CMAX_VIEW_CACHE_SECONDS);
  return result;
}

function cmaxDailyMonthSnapshotSheetName_(month) {
  return 'CMAX_M_' + sanitizeCmaxMonth_(month).replace('-', '_');
}

function readCmaxDailyMonthSnapshot_(month, consultant, meta) {
  if (!month || !meta || !meta.month_sheets || !meta.month_sheets[month]) return null;
  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(meta.month_sheets[month]);
  if (!sheet) return null;
  var headers = getCmaxDailyViewHeaders_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues().map(function(row) {
    var item = {};
    headers.forEach(function(header, index) { item[header] = row[index]; });
    item.contabiliza_diaria = item.contabiliza_diaria === true || String(item.contabiliza_diaria).toLowerCase() === 'true';
    return item;
  }).filter(function(item) {
    return !consultant || sanitizeText_(item.consultor).toUpperCase() === consultant;
  });
}

function readCmaxMaterializedCache_(month, consultant) {
  if (consultant) return readCompressedScriptCache_(cmaxViewCacheKey_({ month: month, consultant: consultant }));
  if (month) return readCompressedScriptCache_(cmaxViewCacheKey_({ month: month }));
  var props = PropertiesService.getScriptProperties();
  var meta;
  try { meta = JSON.parse(props.getProperty('CMAX_VIEW_META_JSON') || 'null'); } catch (ignored) { meta = null; }
  if (!meta || !Array.isArray(meta.available_months)) return null;
  var availableMonths = cmaxRelevantMonths_(meta.available_months);
  var snapshots = availableMonths.map(function(itemMonth) {
    return readCompressedScriptCache_(cmaxViewCacheKey_({ month: itemMonth }));
  });
  if (snapshots.some(function(snapshot) { return !snapshot; })) return null;
  var first = snapshots[0] || {};
  var events = [];
  snapshots.forEach(function(snapshot) { events = events.concat(snapshot.events || []); });
  return {
    ok: true,
    events: events,
    total: events.length,
    month: '',
    available_months: cmaxRelevantMonths_(meta.available_months),
    history_months: cmaxHistoryMonths_(),
    history_loaded_months: cmaxSnapshotMonths_(meta),
    training_team_cutoff: meta.training_team_cutoff || '',
    training_team: meta.training_team || [],
    consultant_compensation: getConsultantCompensationData_().consultants,
    consultant_daily_rates: CONSULTANT_SENIORITY_RATES,
    available_activities: meta.available_activities || [],
    history_sync: first.history_sync || {},
    synced_at: meta.synced_at || '',
    materialized: true,
    cached: true,
    sheet: CMAX_DAILY_VIEW_SHEET
  };
}

function scheduleCmaxDailyViewBuild_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CMAX_VIEW_BUILD_SCHEDULED') === '1') return;
  props.setProperty('CMAX_VIEW_BUILD_SCHEDULED', '1');
  ScriptApp.newTrigger('reconstruirVisaoDiariasCmax').timeBased().after(1000).create();
}

function getCmaxDailyEventsLegacy_(params) {
  params = params || {};
  var cacheKey = cmaxViewCacheKey_(params);
  var cached = readCompressedScriptCache_(cacheKey);
  if (cached) {
    cached.cached = true;
    return cached;
  }
  var sheet = getCmaxDailySheet_();
  var headersExpected = getCmaxDailyHeaders_();
  var rawJsonIndex = headersExpected.indexOf('raw_json');
  var readColumns = rawJsonIndex > 0 ? rawJsonIndex : headersExpected.length;
  var range = sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), readColumns);
  var values = range.getValues();
  var headers = values.length ? values[0].map(function(value) { return String(value || ''); }) : headersExpected.slice(0, readColumns);
  var startTimeIndex = headers.indexOf('hora_inicio');
  var endTimeIndex = headers.indexOf('hora_fim');
  var timeDisplayValues = startTimeIndex >= 0 && endTimeIndex === startTimeIndex + 1
    ? sheet.getRange(1, startTimeIndex + 1, Math.max(1, sheet.getLastRow()), 2).getDisplayValues()
    : [];
  var needsRawJson = values.slice(1).some(function(row, rowIndex) {
    var start = normalizeCmaxSheetTime_((timeDisplayValues[rowIndex + 1] || [])[0] || row[startTimeIndex]);
    var end = normalizeCmaxSheetTime_((timeDisplayValues[rowIndex + 1] || [])[1] || row[endTimeIndex]);
    return !isCmaxValidTime_(start) || !isCmaxValidTime_(end);
  });
  var rawJsonValues = rawJsonIndex >= 0 && needsRawJson
    ? sheet.getRange(1, rawJsonIndex + 1, Math.max(1, sheet.getLastRow()), 1).getValues()
    : [];
  var month = sanitizeCmaxMonth_(params.month || params.mes);
  var consultant = sanitizeText_(params.consultant || params.consultor).toUpperCase();
  var allEvents = values.slice(1).map(function(row, rowIndex) {
    var item = {};
    headers.forEach(function(header, index) { item[header] = row[index]; });
    item.data = normalizeCmaxSheetDate_(item.data);
    item.mes = normalizeCmaxSheetMonth_(item.mes || item.data);
    item.ano = item.mes ? item.mes.slice(0, 4) : String(item.ano || '');
    item.hora_inicio = normalizeCmaxSheetTime_((timeDisplayValues[rowIndex + 1] || [])[0] || item.hora_inicio);
    item.hora_fim = normalizeCmaxSheetTime_((timeDisplayValues[rowIndex + 1] || [])[1] || item.hora_fim);
    if (!isCmaxValidTime_(item.hora_inicio) || !isCmaxValidTime_(item.hora_fim)) {
      var raw = parseCmaxRawJson_((rawJsonValues[rowIndex + 1] || [])[0]);
      if (raw) {
        item.hora_inicio = cmaxEventTime_(raw, [
          'hora_inicio', 'hora_inicial', 'hora_de', 'horario_inicio', 'horario_de', 'inicio_hora', 'hr_inicio', 'start_time'
        ]);
        item.hora_fim = cmaxEventTime_(raw, [
          'hora_fim', 'hora_final', 'hora_ate', 'horario_fim', 'horario_ate', 'fim_hora', 'hr_fim', 'end_time'
        ]);
      }
    }
    item.modalidade = normalizeCmaxModality_(item.descricao || item.tipo);
    item.contabiliza_diaria = isCmaxDailyModality_(item.modalidade);
    delete item.raw_json;
    return item;
  });
  var trainingTeam = {};
  var trainingCutoff = cmaxTrainingTeamCutoff_();
  allEvents.forEach(function(item) {
    if (item.contabiliza_diaria && item.data >= trainingCutoff && isCmaxTrainingConsultant_(item.consultor)) {
      trainingTeam[sanitizeText_(item.consultor).toUpperCase()] = true;
    }
  });
  var teamEvents = allEvents.filter(function(item) {
    return !!trainingTeam[sanitizeText_(item.consultor).toUpperCase()];
  });
  var events = teamEvents.filter(function(item) {
    if (month && item.mes !== month) return false;
    if (consultant && sanitizeText_(item.consultor).toUpperCase() !== consultant) return false;
    return true;
  });
  var availableMonths = {};
  var availableConsultants = {};
  var availableActivities = {};
  teamEvents.forEach(function(item) {
    if (item.mes) availableMonths[item.mes] = true;
    if (item.consultor) availableConsultants[item.consultor] = true;
    if (item.descricao || item.modalidade || item.tipo) availableActivities[item.descricao || item.modalidade || item.tipo] = true;
  });

  var result = {
    ok: true,
    events: events,
    total: events.length,
    month: month,
    available_months: cmaxRelevantMonths_(Object.keys(availableMonths)),
    history_months: cmaxHistoryMonths_(),
    history_loaded_months: cmaxSnapshotMonths_(),
    training_team_cutoff: trainingCutoff,
    training_team: Object.keys(availableConsultants).sort(),
    available_activities: Object.keys(availableActivities).sort(),
    history_sync: getCmaxDailyHistoryBackgroundStatus_(),
    synced_at: events.reduce(function(latest, item) {
      var value = String(item.sincronizado_em || '');
      return value > latest ? value : latest;
    }, ''),
    sheet: CMAX_DAILY_SHEET
  };
  writeCompressedScriptCache_(cacheKey, result, CMAX_VIEW_CACHE_SECONDS);
  return result;
}

function cmaxViewCacheKey_(params) {
  params = params || {};
  var month = sanitizeCmaxMonth_(params.month || params.mes) || 'all';
  var consultant = sanitizeText_(params.consultant || params.consultor).toUpperCase() || 'all';
  return 'cmax:view:v3:' + month + ':' + consultant;
}

function readCompressedScriptCache_(key) {
  try {
    var encoded = CacheService.getScriptCache().get(key);
    if (!encoded) return null;
    var bytes = Utilities.base64Decode(encoded);
    var json = Utilities.ungzip(Utilities.newBlob(bytes)).getDataAsString('UTF-8');
    return JSON.parse(json);
  } catch (ignored) {
    return null;
  }
}

function writeCompressedScriptCache_(key, value, seconds) {
  try {
    var compressed = Utilities.gzip(Utilities.newBlob(JSON.stringify(value), 'application/json'));
    var encoded = Utilities.base64Encode(compressed.getBytes());
    CacheService.getScriptCache().put(key, encoded, Math.max(60, Math.min(Number(seconds || 3600), 21600)));
  } catch (ignored) {}
}

function clearCmaxViewCacheForMonths_(months) {
  var keys = (months || []).map(function(month) { return cmaxViewCacheKey_({ month: month }); });
  if (keys.length) CacheService.getScriptCache().removeAll(keys);
}

function cmaxTrainingTeamCutoff_() {
  var props = PropertiesService.getScriptProperties();
  var months = Math.max(1, Math.min(toInt_(props.getProperty('CMAX_TRAINING_TEAM_MONTHS'), 6), 60));
  var now = new Date();
  var cutoff = new Date(now.getFullYear(), now.getMonth() - months, 1);
  return Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function startCmaxDailyHistoryBackground_(params) {
  params = params || {};
  var props = PropertiesService.getScriptProperties();
  var currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var startMonth = cmaxHistoryStartMonth_(params.start_month);
  if (cmaxSnapshotMonths_().indexOf(currentMonth) < 0) {
    addCmaxForcedHistoryMonths_([currentMonth]);
  } else {
    removeCmaxForcedHistoryMonths_([currentMonth]);
  }
  props.setProperty('CMAX_HISTORY_BACKGROUND_ACTIVE', '1');
  props.setProperty('CMAX_HISTORY_BACKGROUND_START_MONTH', startMonth);
  props.setProperty('CMAX_HISTORY_BACKGROUND_FAILURES', '0');
  props.setProperty('CMAX_HISTORY_BACKGROUND_UPDATED_AT', new Date().toISOString());
  props.deleteProperty('CMAX_HISTORY_BACKGROUND_ERROR');
  props.setProperty('CMAX_HISTORY_BACKGROUND_LAST_MONTH', currentMonth);
  props.setProperty('CMAX_HISTORY_BACKGROUND_STARTED_AT', new Date().toISOString());
  props.deleteProperty('CMAX_HISTORY_BACKGROUND_COMPLETED_AT');
  updateCmaxHistoryProgressProperties_();
  scheduleCmaxDailyHistoryBackground_(5000);
  return {
    ok: true,
    scheduled: true,
    current_month: currentMonth,
    start_month: startMonth,
    history_sync: getCmaxDailyHistoryBackgroundStatus_()
  };
}

function continueCmaxDailyHistoryBackgroundTrigger() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CMAX_HISTORY_BACKGROUND_ACTIVE') !== '1') {
    clearCmaxDailyHistoryBackgroundTriggers_();
    return;
  }
  try {
    var result = continueCmaxDailyHistoryBatch_({ batch_size: props.getProperty('CMAX_HISTORY_BATCH_SIZE') || '1' });
    if (result.done) return;
    scheduleCmaxDailyHistoryBackground_(5000);
  } catch (error) {
    var failures = toInt_(props.getProperty('CMAX_HISTORY_BACKGROUND_FAILURES'), 0) + 1;
    props.setProperty('CMAX_HISTORY_BACKGROUND_FAILURES', String(failures));
    props.setProperty('CMAX_HISTORY_BACKGROUND_ERROR', simplifyErrorMessage_(error));
    if (failures >= 10) {
      props.setProperty('CMAX_HISTORY_BACKGROUND_ACTIVE', '0');
      clearCmaxDailyHistoryBackgroundTriggers_();
      return;
    }
    scheduleCmaxDailyHistoryBackground_(60000);
  }
}

function continueCmaxDailyHistoryBatch_(params) {
  params = params || {};
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CMAX_HISTORY_BACKGROUND_ACTIVE') !== '1') {
    return { ok: true, done: true, history_sync: getCmaxDailyHistoryBackgroundStatus_() };
  }
  var batchSize = Math.max(1, Math.min(toInt_(params.batch_size, 1), 2));
  var pendingMonths = cmaxPendingHistoryMonths_();
  var batchMonths = pendingMonths.slice(0, batchSize);
  var attemptedMonths = batchMonths.slice();
  var eventsByMonth = {};
  batchMonths.forEach(function(month) {
    eventsByMonth[month] = fetchCmaxDailyEventsForMonth_(month).events;
  });
  if (batchMonths.length) {
    var unchangedMonths = batchMonths.filter(function(month) {
      return cmaxDailyMonthMatches_(month, eventsByMonth[month]);
    });
    if (unchangedMonths.length) markCmaxHistoryMonthsCompleted_(unchangedMonths);
    unchangedMonths.forEach(function(month) { delete eventsByMonth[month]; });
    batchMonths = batchMonths.filter(function(month) { return unchangedMonths.indexOf(month) < 0; });
  }
  if (batchMonths.length) {
    var writeLock = LockService.getScriptLock();
    if (!writeLock.tryLock(1000)) {
      return { ok: true, done: false, busy: true, history_sync: getCmaxDailyHistoryBackgroundStatus_() };
    }
    try {
      Object.keys(eventsByMonth).forEach(function(month) {
        writeCmaxDailyMonthSnapshot_(month, eventsByMonth[month]);
      });
      markCmaxHistoryMonthsCompleted_(batchMonths);
    } finally {
      writeLock.releaseLock();
    }
  }
  if (attemptedMonths.length) props.setProperty('CMAX_HISTORY_BACKGROUND_LAST_MONTH', attemptedMonths[attemptedMonths.length - 1]);
  var remainingMonths = cmaxPendingHistoryMonths_();
  props.setProperty('CMAX_HISTORY_BACKGROUND_FAILURES', '0');
  props.setProperty('CMAX_HISTORY_BACKGROUND_UPDATED_AT', new Date().toISOString());
  props.deleteProperty('CMAX_HISTORY_BACKGROUND_ERROR');
  var done = remainingMonths.length === 0;
  if (done) {
    props.setProperty('CMAX_HISTORY_BACKGROUND_ACTIVE', '0');
    props.setProperty('CMAX_HISTORY_BACKGROUND_COMPLETED_AT', new Date().toISOString());
    clearCmaxDailyHistoryBackgroundTriggers_();
  }
  updateCmaxHistoryProgressProperties_();
  return {
    ok: true,
    done: done,
    batch_months: batchMonths,
    history_sync: getCmaxDailyHistoryBackgroundStatus_()
  };
}

function getCmaxDailyHistoryBackgroundStatus_() {
  var props = PropertiesService.getScriptProperties();
  var pending = cmaxPendingHistoryMonths_();
  var completed = cmaxProcessedHistoryMonths_();
  return {
    active: props.getProperty('CMAX_HISTORY_BACKGROUND_ACTIVE') === '1',
    start_month: cmaxHistoryStartMonth_(),
    next_month: pending[0] || '',
    last_month: props.getProperty('CMAX_HISTORY_BACKGROUND_LAST_MONTH') || '',
    processed_months: completed.length,
    started_at: props.getProperty('CMAX_HISTORY_BACKGROUND_STARTED_AT') || '',
    updated_at: props.getProperty('CMAX_HISTORY_BACKGROUND_UPDATED_AT') || '',
    completed_at: props.getProperty('CMAX_HISTORY_BACKGROUND_COMPLETED_AT') || '',
    error: props.getProperty('CMAX_HISTORY_BACKGROUND_ERROR') || ''
  };
}

function getCmaxDailyHistoryStatus_() {
  var props = PropertiesService.getScriptProperties();
  var loaded = cmaxSnapshotMonths_();
  var history = cmaxHistoryMonths_();
  var loadedMap = {};
  loaded.forEach(function(month) {
    month = sanitizeCmaxMonth_(month);
    if (month) loadedMap[month] = true;
  });
  var pending = history.filter(function(month) { return !loadedMap[month]; });
  return {
    ok: true,
    history_months: history,
    history_loaded_months: history.filter(function(month) { return !!loadedMap[month]; }),
    history_sync: {
      active: props.getProperty('CMAX_HISTORY_BACKGROUND_ACTIVE') === '1',
      start_month: cmaxHistoryStartMonth_(),
      next_month: pending[0] || '',
      last_month: props.getProperty('CMAX_HISTORY_BACKGROUND_LAST_MONTH') || '',
      processed_months: history.length - pending.length,
      started_at: props.getProperty('CMAX_HISTORY_BACKGROUND_STARTED_AT') || '',
      updated_at: props.getProperty('CMAX_HISTORY_BACKGROUND_UPDATED_AT') || '',
      completed_at: props.getProperty('CMAX_HISTORY_BACKGROUND_COMPLETED_AT') || '',
      error: props.getProperty('CMAX_HISTORY_BACKGROUND_ERROR') || ''
    }
  };
}

function cmaxSnapshotMonths_(meta) {
  if (!meta) {
    try { meta = JSON.parse(PropertiesService.getScriptProperties().getProperty('CMAX_VIEW_META_JSON') || 'null'); } catch (ignored) { meta = null; }
  }
  return cmaxRelevantMonths_(Object.keys(meta && meta.month_sheets || {}));
}

function scheduleCmaxDailyHistoryBackground_(delayMs) {
  clearCmaxDailyHistoryBackgroundTriggers_();
  ScriptApp.newTrigger('continueCmaxDailyHistoryBackgroundTrigger')
    .timeBased()
    .after(Math.max(1000, Number(delayMs || 5000)))
    .create();
}

function clearCmaxDailyHistoryBackgroundTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'continueCmaxDailyHistoryBackgroundTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function cmaxShiftMonth_(month, offset) {
  var parts = sanitizeCmaxMonth_(month).split('-');
  if (parts.length !== 2) return '';
  var absoluteMonth = Number(parts[0]) * 12 + Number(parts[1]) - 1 + Number(offset || 0);
  var year = Math.floor(absoluteMonth / 12);
  var monthNumber = absoluteMonth - year * 12 + 1;
  return year + '-' + String(monthNumber).padStart(2, '0');
}

function cmaxHistoryMonths_() {
  var current = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var start = cmaxHistoryStartMonth_();
  var months = [];
  while (current && current >= start && months.length < 120) {
    months.push(current);
    current = cmaxShiftMonth_(current, -1);
  }
  return months;
}

function cmaxHistoryStartMonth_(requested) {
  var props = PropertiesService.getScriptProperties();
  var configured = sanitizeCmaxMonth_(requested || props.getProperty('CMAX_HISTORY_START_MONTH')) || CMAX_HISTORY_DEFAULT_START_MONTH;
  return configured < CMAX_HISTORY_DEFAULT_START_MONTH ? CMAX_HISTORY_DEFAULT_START_MONTH : configured;
}

function cmaxRelevantMonths_(months) {
  var start = cmaxHistoryStartMonth_();
  return (months || []).map(sanitizeCmaxMonth_).filter(function(month) {
    return !!month && month >= start;
  }).filter(function(month, index, all) {
    return all.indexOf(month) === index;
  }).sort().reverse();
}

function cmaxProcessedHistoryMonths_() {
  var allMonths = cmaxHistoryMonths_();
  var completed = cmaxHistoryCompletionMap_();
  return allMonths.filter(function(month) { return !!completed[month]; });
}

var CMAX_HISTORY_COMPLETION_CACHE_ = null;

function cmaxHistoryCompletionMap_() {
  if (CMAX_HISTORY_COMPLETION_CACHE_) return CMAX_HISTORY_COMPLETION_CACHE_;
  var props = PropertiesService.getScriptProperties();
  var completed = {};
  try {
    JSON.parse(props.getProperty('CMAX_HISTORY_COMPLETED_MONTHS_JSON') || '[]').forEach(function(month) {
      month = sanitizeCmaxMonth_(month);
      if (month) completed[month] = true;
    });
  } catch (ignored) {}
  try {
    var meta = JSON.parse(props.getProperty('CMAX_VIEW_META_JSON') || 'null');
    (meta && meta.history_loaded_months || []).forEach(function(month) {
      month = sanitizeCmaxMonth_(month);
      if (month) completed[month] = true;
    });
  } catch (ignoredMeta) {}
  CMAX_HISTORY_COMPLETION_CACHE_ = completed;
  return completed;
}

function cmaxMonthsPresentInSheet_() {
  var sheet = getCmaxDailySheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var monthColumn = getCmaxDailyHeaders_().indexOf('mes') + 1;
  var found = {};
  sheet.getRange(2, monthColumn, lastRow - 1, 1).getValues().forEach(function(row) {
    var month = normalizeCmaxSheetMonth_(row[0]);
    if (month) found[month] = true;
  });
  return Object.keys(found).sort().reverse();
}

function cmaxPendingHistoryMonths_() {
  var completed = {};
  cmaxProcessedHistoryMonths_().forEach(function(month) { completed[month] = true; });
  var forced = cmaxForcedHistoryMonths_();
  var props = PropertiesService.getScriptProperties();
  var meta;
  try { meta = JSON.parse(props.getProperty('CMAX_VIEW_META_JSON') || 'null'); } catch (ignored) { meta = null; }
  var monthSheets = meta && meta.month_sheets || {};
  var missingSnapshots = cmaxHistoryMonths_().filter(function(month) {
    return !monthSheets[month] && forced.indexOf(month) < 0;
  });
  return forced.concat(missingSnapshots).concat(cmaxHistoryMonths_().filter(function(month) {
    return !completed[month] && forced.indexOf(month) < 0 && missingSnapshots.indexOf(month) < 0;
  }));
}

function cmaxDailyMonthMatches_(month, events) {
  var props = PropertiesService.getScriptProperties();
  var meta;
  try { meta = JSON.parse(props.getProperty('CMAX_VIEW_META_JSON') || 'null'); } catch (ignored) { meta = null; }
  if (!meta || !meta.month_sheets || !meta.month_sheets[month]) return false;
  var previousHash = meta && meta.month_hashes && meta.month_hashes[month];
  return !!previousHash && previousHash === cmaxEventsHash_(events || []);
}

function cmaxEventsHash_(events) {
  var fields = ['event_key', 'data', 'consultor', 'cliente', 'tipo', 'resultado', 'descricao', 'hora_inicio', 'hora_fim'];
  var signatures = (events || []).map(function(event) {
    return fields.map(function(field) {
      var value = event && event[field];
      if (field === 'data') return normalizeCmaxSheetDate_(value);
      if (field === 'hora_inicio' || field === 'hora_fim') return normalizeCmaxSheetTime_(value);
      return sanitizeText_(value);
    }).join('|');
  }).sort();
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    signatures.join('\n'),
    Utilities.Charset.UTF_8
  ));
}

function markCmaxHistoryMonthsCompleted_(months) {
  var props = PropertiesService.getScriptProperties();
  var completed = {};
  cmaxProcessedHistoryMonths_().forEach(function(month) { completed[month] = true; });
  (months || []).forEach(function(month) {
    month = sanitizeCmaxMonth_(month);
    if (month) completed[month] = true;
  });
  props.setProperty('CMAX_HISTORY_COMPLETED_MONTHS_JSON', JSON.stringify(Object.keys(completed).sort().reverse()));
  removeCmaxForcedHistoryMonths_(months);
  CMAX_HISTORY_COMPLETION_CACHE_ = completed;
}

function cmaxForcedHistoryMonths_() {
  var props = PropertiesService.getScriptProperties();
  try {
    return JSON.parse(props.getProperty('CMAX_HISTORY_FORCED_MONTHS_JSON') || '[]').map(sanitizeCmaxMonth_).filter(function(month) {
      return !!month;
    });
  } catch (ignored) {
    return [];
  }
}

function addCmaxForcedHistoryMonths_(months) {
  var props = PropertiesService.getScriptProperties();
  var forced = cmaxForcedHistoryMonths_();
  (months || []).forEach(function(month) {
    month = sanitizeCmaxMonth_(month);
    if (month && forced.indexOf(month) < 0) forced.push(month);
  });
  props.setProperty('CMAX_HISTORY_FORCED_MONTHS_JSON', JSON.stringify(forced));
}

function removeCmaxForcedHistoryMonths_(months) {
  var props = PropertiesService.getScriptProperties();
  var remove = {};
  (months || []).forEach(function(month) { remove[sanitizeCmaxMonth_(month)] = true; });
  props.setProperty('CMAX_HISTORY_FORCED_MONTHS_JSON', JSON.stringify(cmaxForcedHistoryMonths_().filter(function(month) {
    return !remove[month];
  })));
}

function updateCmaxHistoryProgressProperties_() {
  var props = PropertiesService.getScriptProperties();
  var pending = cmaxPendingHistoryMonths_();
  props.setProperty('CMAX_HISTORY_BACKGROUND_CURSOR', pending[0] || '');
  props.setProperty('CMAX_HISTORY_BACKGROUND_PROCESSED', String(cmaxProcessedHistoryMonths_().length));
}

function syncCmaxDailyEvents_(params) {
  params = params || {};
  var month = sanitizeCmaxMonth_(params.month || params.mes) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var fetched = fetchCmaxDailyEventsForMonth_(month);
  var writeLock = LockService.getScriptLock();
  if (!writeLock.tryLock(60000)) throw new Error('Outra atualização CMAX está finalizando uma gravação. Aguarde alguns segundos.');
  try {
    writeCmaxDailyMonthSnapshot_(month, fetched.events);
    markCmaxHistoryMonthsCompleted_([month]);
    updateCmaxHistoryProgressProperties_();
  } finally {
    writeLock.releaseLock();
  }
  return fetched.result;
}

function writeCmaxDailyMonthSnapshot_(month, events) {
  month = sanitizeCmaxMonth_(month);
  if (!month) return;
  var props = PropertiesService.getScriptProperties();
  var meta;
  try { meta = JSON.parse(props.getProperty('CMAX_VIEW_META_JSON') || '{}'); } catch (ignored) { meta = {}; }
  meta.month_sheets = meta.month_sheets || {};
  meta.month_hashes = meta.month_hashes || {};
  meta.available_months = meta.available_months || [];
  meta.history_loaded_months = meta.history_loaded_months || [];
  meta.training_team = meta.training_team || [];
  meta.available_activities = meta.available_activities || [];

  var trainingTeam = {};
  meta.training_team.forEach(function(name) { trainingTeam[sanitizeText_(name).toUpperCase()] = true; });
  (events || []).forEach(function(event) {
    var modality = normalizeCmaxModality_(event.descricao || event.tipo);
    if (isCmaxDailyModality_(modality) && isCmaxTrainingConsultant_(event.consultor)) trainingTeam[sanitizeText_(event.consultor).toUpperCase()] = true;
  });
  var activities = {};
  meta.available_activities.forEach(function(activity) { activities[activity] = true; });
  var items = (events || []).map(function(event) {
    var item = {};
    getCmaxDailyViewHeaders_().forEach(function(header) { item[header] = event[header] === undefined ? '' : event[header]; });
    item.data = normalizeCmaxSheetDate_(event.data);
    item.mes = month;
    item.ano = month.slice(0, 4);
    item.hora_inicio = normalizeCmaxSheetTime_(event.hora_inicio);
    item.hora_fim = normalizeCmaxSheetTime_(event.hora_fim);
    item.modalidade = normalizeCmaxModality_(event.descricao || event.tipo);
    item.contabiliza_diaria = isCmaxDailyModality_(item.modalidade);
    if (event.descricao || item.modalidade || event.tipo) activities[event.descricao || item.modalidade || event.tipo] = true;
    return item;
  }).filter(function(item) {
    return !!trainingTeam[sanitizeText_(item.consultor).toUpperCase()];
  }).sort(function(a, b) {
    return String(a.consultor).localeCompare(String(b.consultor), 'pt-BR') ||
      String(a.data).localeCompare(String(b.data)) ||
      String(a.hora_inicio).localeCompare(String(b.hora_inicio));
  });

  var ss = SpreadsheetApp.openById(getScriptProperty_('SHEET_ID'));
  var sheetName = cmaxDailyMonthSnapshotSheetName_(month);
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var headers = getCmaxDailyViewHeaders_();
  var values = items.map(function(item) {
    return headers.map(function(header) { return item[header] === undefined ? '' : item[header]; });
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (values.length) sheet.getRange(2, 1, values.length, headers.length).setValues(values);

  meta.month_sheets[month] = sheetName;
  meta.month_hashes[month] = cmaxEventsHash_(events || []);
  if (meta.available_months.indexOf(month) < 0) meta.available_months.push(month);
  if (meta.history_loaded_months.indexOf(month) < 0) meta.history_loaded_months.push(month);
  meta.available_months = cmaxRelevantMonths_(meta.available_months);
  meta.history_loaded_months = cmaxRelevantMonths_(meta.history_loaded_months);
  meta.training_team = Object.keys(trainingTeam).sort();
  meta.available_activities = Object.keys(activities).sort();
  meta.synced_at = new Date().toISOString();
  props.setProperty('CMAX_VIEW_META_JSON', JSON.stringify(meta));
  writeCmaxMaterializedCaches_(items, meta);
}

function fetchCmaxDailyEventsForMonth_(month) {
  var range = cmaxMonthRange_(month);
  var response = fetchCmaxAgendaWithAutomaticToken_(range);
  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code === 401 || code === 403) throw new Error('Autenticação CMAX recusada mesmo após renovação automática.');
  if (code < 200 || code >= 300) throw new Error('CMAX respondeu HTTP ' + code + ': ' + text.slice(0, 240));

  var payload;
  try { payload = JSON.parse(text); } catch (error) { throw new Error('CMAX retornou uma resposta que não é JSON.'); }
  var candidates = collectCmaxAgendaCandidates_(payload);
  var syncedAt = new Date().toISOString();
  var events = candidates.map(function(item) {
    return normalizeCmaxAgendaEvent_(item, syncedAt);
  }).filter(function(item) {
    return item && item.mes === month && isCmaxPositiveResult_(item.resultado);
  });

  var unique = {};
  events.forEach(function(item) { unique[item.event_key] = item; });
  events = Object.keys(unique).map(function(key) { return unique[key]; });

  return {
    events: events,
    result: {
      ok: true,
      month: month,
      imported: events.length,
      candidates: candidates.length,
      ignored: Math.max(0, candidates.length - events.length),
      synced_at: syncedAt
    }
  };
}

function fetchCmaxAgendaWithAutomaticToken_(range) {
  var props = PropertiesService.getScriptProperties();
  var token = String(props.getProperty('CMAX_JWT_TOKEN') || '').trim();
  if (!token) token = renewCmaxJwtToken_('');
  var response = fetchCmaxAgendaWithToken_(range, token);
  if (response.getResponseCode() !== 401 && response.getResponseCode() !== 403) return response;
  token = renewCmaxJwtToken_(token);
  return fetchCmaxAgendaWithToken_(range, token);
}

function fetchCmaxAgendaWithToken_(range, token) {
  return UrlFetchApp.fetch(CMAX_AGENDA_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'JWT ' + token,
      Accept: 'application/json, text/plain, */*',
      'Django-Timezone': 'America/Sao_Paulo',
      Referer: 'https://www.multbovinos.com/'
    },
    payload: JSON.stringify({
      origem: 'tela',
      data_de: range.start,
      data_ate: range.end
    }),
    muteHttpExceptions: true
  });
}

function renewCmaxJwtToken_(currentToken) {
  var props = PropertiesService.getScriptProperties();
  var refreshUrl = String(props.getProperty('CMAX_TOKEN_REFRESH_URL') || '').trim();
  if (currentToken && refreshUrl) {
    var refreshed = requestCmaxToken_(refreshUrl, { token: currentToken });
    if (refreshed) {
      props.setProperty('CMAX_JWT_TOKEN', refreshed);
      return refreshed;
    }
  }
  var username = String(props.getProperty('CMAX_EMAIL') || props.getProperty('CMAX_USERNAME') || '').trim();
  var password = String(props.getProperty('CMAX_PASSWORD') || '');
  if (username && password) {
    var authUrl = String(props.getProperty('CMAX_TOKEN_AUTH_URL') || CMAX_TOKEN_AUTH_URL).trim();
    var authenticated = requestCmaxToken_(authUrl, { email: username, password: password });
    if (authenticated) {
      props.setProperty('CMAX_JWT_TOKEN', authenticated);
      return authenticated;
    }
  }
  throw new Error('Não foi possível renovar a autenticação CMAX automaticamente. Configure CMAX_EMAIL e CMAX_PASSWORD nas propriedades do Apps Script.');
}

function requestCmaxToken_(url, payload) {
  if (!url) return '';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Accept: 'application/json, text/plain, */*' },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return '';
  try {
    var parsed = JSON.parse(response.getContentText());
    return String(parsed.token || parsed.access || parsed.jwt || (parsed.dados_acesso || {}).token || '').trim();
  } catch (ignored) {
    return '';
  }
}

function collectCmaxAgendaCandidates_(value, depth, output, seen, inheritedDate) {
  depth = depth || 0;
  output = output || [];
  seen = seen || {};
  if (value === null || value === undefined || depth > 10) return output;
  if (Array.isArray(value)) {
    value.forEach(function(item) { collectCmaxAgendaCandidates_(item, depth + 1, output, seen, inheritedDate); });
    return output;
  }
  if (typeof value !== 'object') return output;
  var ownDate = cmaxOwnValue_(value, ['data', 'data_evento', 'data_inicio', 'start', 'date']);
  var eventDate = ownDate || inheritedDate;
  var result = cmaxOwnValue_(value, ['resultado_texto', 'resultado', 'status_texto', 'status']);
  var id = cmaxOwnValue_(value, ['id', 'evento_id', 'event_id', 'uid']);
  var eventSignal = cmaxOwnValue_(value, [
    'contato', 'contato_texto', 'cliente', 'cliente_nome', 'responsavel', 'responsavel_texto',
    'grupo_evento', 'grupo_evento_texto', 'tipo_evento_texto', 'tipo_comunicacao'
  ]);
  if (eventDate && result !== '' && eventSignal !== '') {
    if (!ownDate) {
      value = cmaxCloneWithParentDate_(value, eventDate);
    }
    var signature = String(id || '') + '|' + String(eventDate || '') + '|' + JSON.stringify(value).slice(0, 180);
    if (!seen[signature]) {
      seen[signature] = true;
      output.push(value);
    }
  }
  Object.keys(value).forEach(function(key) {
    collectCmaxAgendaCandidates_(value[key], depth + 1, output, seen, eventDate);
  });
  return output;
}

function normalizeCmaxAgendaEvent_(raw, syncedAt) {
  var rawDate = deepFindFirst_(raw, ['data', 'data_evento', 'data_inicio', 'start', 'date', '_cmax_parent_date']);
  var date = parseCmaxDate_(rawDate);
  if (!date) return null;
  var eventId = sanitizeText_(deepFindFirst_(raw, ['id', 'evento_id', 'event_id', 'uid']));
  var consultant = cmaxScalarText_(deepFindFirst_(raw, [
    'responsavel_texto', 'responsavel_nome', 'consultor', 'usuario_texto', 'responsavel'
  ]));
  var client = cmaxScalarText_(deepFindFirst_(raw, [
    'contato_texto', 'contato_nome', 'cliente', 'cliente_nome', 'contato'
  ]));
  var type = cmaxScalarText_(deepFindFirst_(raw, [
    'grupo_evento_texto', 'tipo_evento_texto', 'tipo_texto', 'atividade_texto', 'descricao_tipo', 'titulo', 'grupo_evento'
  ]));
  var result = cmaxScalarText_(deepFindFirst_(raw, ['resultado_texto', 'resultado', 'status_texto', 'status']));
  if (/^1(?:\.0)?$/.test(result)) result = 'Positivo';
  var description = sanitizeText_(deepFindFirst_(raw, ['descricao', 'observacao', 'obs', 'assunto', 'titulo']));
  var isoDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var month = isoDate.slice(0, 7);
  var key = eventId || [isoDate, consultant, client, type, result].join('|').toUpperCase();
  return {
    event_key: key,
    event_id: eventId,
    data: isoDate,
    mes: month,
    ano: month.slice(0, 4),
    consultor: consultant || 'Sem consultor',
    cliente: client || 'Cliente não identificado',
    tipo: type || 'Agenda CMAX',
    resultado: result,
    descricao: description,
    hora_inicio: cmaxEventTime_(raw, [
      'hora_inicio', 'hora_inicial', 'hora_de', 'horario_inicio', 'horario_de', 'inicio_hora', 'hr_inicio', 'start_time'
    ]),
    hora_fim: cmaxEventTime_(raw, [
      'hora_fim', 'hora_final', 'hora_ate', 'horario_fim', 'horario_ate', 'fim_hora', 'hr_fim', 'end_time'
    ]),
    sincronizado_em: syncedAt,
    raw_json: JSON.stringify(raw)
  };
}

function cmaxScalarText_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value !== 'object') return sanitizeText_(value);
  return sanitizeText_(deepFindFirst_(value, ['nome', 'name', 'texto', 'descricao', 'label', 'status', 'resultado', 'id']));
}

function cmaxOwnValue_(value, aliases) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return '';
  var aliasMap = {};
  (aliases || []).forEach(function(alias) { aliasMap[String(alias).toLowerCase()] = true; });
  var keys = Object.keys(value);
  for (var i = 0; i < keys.length; i++) {
    if (aliasMap[String(keys[i]).toLowerCase()]) {
      var found = value[keys[i]];
      return found === null || found === undefined ? '' : found;
    }
  }
  return '';
}

function cmaxCloneWithParentDate_(value, date) {
  var clone = {};
  Object.keys(value || {}).forEach(function(key) { clone[key] = value[key]; });
  clone._cmax_parent_date = date;
  return clone;
}

function cmaxEventTime_(raw, aliases) {
  var own = cmaxOwnValue_(raw, aliases);
  var normalizedOwn = normalizeCmaxSheetTime_(own);
  if (/^\d{2}:\d{2}$/.test(normalizedOwn)) return normalizedOwn;
  var nested = deepFindFirst_(raw, aliases);
  var normalizedNested = normalizeCmaxSheetTime_(nested);
  if (/^\d{2}:\d{2}$/.test(normalizedNested)) return normalizedNested;
  var wantsEnd = (aliases || []).some(function(alias) {
    return /fim|final|ate|end|termino/i.test(normalizeCmaxKey_(alias));
  });
  return findCmaxSemanticTime_(raw, wantsEnd ? 'end' : 'start');
}

function normalizeCmaxKey_(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '_');
}

function findCmaxSemanticTime_(value, mode, depth) {
  depth = depth || 0;
  if (value === null || value === undefined || depth > 10) return '';
  if (Array.isArray(value)) {
    for (var arrayIndex = 0; arrayIndex < value.length; arrayIndex++) {
      var arrayFound = findCmaxSemanticTime_(value[arrayIndex], mode, depth + 1);
      if (arrayFound) return arrayFound;
    }
    return '';
  }
  if (typeof value !== 'object') return '';
  var keys = Object.keys(value);
  var keyPattern = mode === 'end'
    ? /(^|_)(fim|final|ate|end|termino|encerramento)($|_)/
    : /(^|_)(inicio|inicial|start|de)($|_)/;
  for (var index = 0; index < keys.length; index++) {
    var key = normalizeCmaxKey_(keys[index]);
    if (!keyPattern.test(key)) continue;
    var direct = normalizeCmaxSheetTime_(value[keys[index]]);
    if (isCmaxValidTime_(direct)) return direct;
  }
  for (var intervalIndex = 0; intervalIndex < keys.length; intervalIndex++) {
    var text = typeof value[keys[intervalIndex]] === 'string' ? value[keys[intervalIndex]] : '';
    var interval = text.match(/(?:^|\D)(\d{1,2}:\d{2})\s*(?:-|a|ate|até)\s*(\d{1,2}:\d{2})(?:\D|$)/i);
    if (interval) return normalizeCmaxSheetTime_(mode === 'end' ? interval[2] : interval[1]);
  }
  for (var childIndex = 0; childIndex < keys.length; childIndex++) {
    var found = findCmaxSemanticTime_(value[keys[childIndex]], mode, depth + 1);
    if (found) return found;
  }
  return '';
}

function parseCmaxRawJson_(value) {
  if (!value) return null;
  if (typeof value === 'object') return value;
  try { return JSON.parse(String(value)); } catch (error) { return null; }
}

function isCmaxValidTime_(value) {
  return /^\d{2}:\d{2}$/.test(String(value || ''));
}

function replaceCmaxDailyMonth_(month, events) {
  var eventsByMonth = {};
  eventsByMonth[month] = events;
  replaceCmaxDailyMonths_(eventsByMonth);
}

function replaceCmaxDailyMonths_(eventsByMonth) {
  var sheet = getCmaxDailySheet_();
  var headers = getCmaxDailyHeaders_();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var displayValues = range.getDisplayValues();
  var monthIndex = headers.indexOf('mes');
  var startTimeIndex = headers.indexOf('hora_inicio');
  var endTimeIndex = headers.indexOf('hora_fim');
  var rawJsonIndex = headers.indexOf('raw_json');
  var replacedMonths = {};
  Object.keys(eventsByMonth || {}).forEach(function(month) { replacedMonths[month] = true; });
  var retained = values.slice(1).map(function(row, rowIndex) {
    var copy = row.slice();
    copy[startTimeIndex] = normalizeCmaxSheetTime_(displayValues[rowIndex + 1][startTimeIndex] || copy[startTimeIndex]);
    copy[endTimeIndex] = normalizeCmaxSheetTime_(displayValues[rowIndex + 1][endTimeIndex] || copy[endTimeIndex]);
    if (!isCmaxValidTime_(copy[startTimeIndex]) || !isCmaxValidTime_(copy[endTimeIndex])) {
      var raw = parseCmaxRawJson_(copy[rawJsonIndex]);
      if (raw) {
        copy[startTimeIndex] = cmaxEventTime_(raw, [
          'hora_inicio', 'hora_inicial', 'hora_de', 'horario_inicio', 'horario_de', 'inicio_hora', 'hr_inicio', 'start_time'
        ]);
        copy[endTimeIndex] = cmaxEventTime_(raw, [
          'hora_fim', 'hora_final', 'hora_ate', 'horario_fim', 'horario_ate', 'fim_hora', 'hr_fim', 'end_time'
        ]);
      }
    }
    return copy;
  }).filter(function(row) { return !replacedMonths[normalizeCmaxSheetMonth_(row[monthIndex])]; });
  var added = [];
  Object.keys(eventsByMonth || {}).forEach(function(month) {
    (eventsByMonth[month] || []).forEach(function(item) {
      added.push(headers.map(function(header) { return item[header] === undefined ? '' : item[header]; }));
    });
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, startTimeIndex + 1, Math.max(1, retained.length + added.length), 2).setNumberFormat('@');
  if (retained.length + added.length) {
    sheet.getRange(2, 1, retained.length + added.length, headers.length).setValues(retained.concat(added));
  }
  rebuildCmaxDailyMaterializedView_(retained.concat(added), headers);
}

function rebuildCmaxDailyMaterializedView_(rows, sourceHeaders) {
  var viewHeaders = getCmaxDailyViewHeaders_();
  var trainingCutoff = cmaxTrainingTeamCutoff_();
  var items = (rows || []).map(function(row) {
    var item = {};
    sourceHeaders.forEach(function(header, index) { item[header] = row[index]; });
    item.data = normalizeCmaxSheetDate_(item.data);
    item.mes = normalizeCmaxSheetMonth_(item.mes || item.data);
    item.ano = item.mes ? item.mes.slice(0, 4) : String(item.ano || '');
    item.hora_inicio = normalizeCmaxSheetTime_(item.hora_inicio);
    item.hora_fim = normalizeCmaxSheetTime_(item.hora_fim);
    item.modalidade = normalizeCmaxModality_(item.descricao || item.tipo);
    item.contabiliza_diaria = isCmaxDailyModality_(item.modalidade);
    return item;
  });
  var trainingTeamMap = {};
  items.forEach(function(item) {
    if (item.contabiliza_diaria && item.data >= trainingCutoff && isCmaxTrainingConsultant_(item.consultor)) {
      trainingTeamMap[sanitizeText_(item.consultor).toUpperCase()] = true;
    }
  });
  var monthHashes = {};
  var rawMonthItems = {};
  items.forEach(function(item) {
    (rawMonthItems[item.mes] || (rawMonthItems[item.mes] = [])).push(item);
  });
  Object.keys(rawMonthItems).forEach(function(month) { monthHashes[month] = cmaxEventsHash_(rawMonthItems[month]); });
  items = items.filter(function(item) {
    return !!trainingTeamMap[sanitizeText_(item.consultor).toUpperCase()];
  }).sort(function(a, b) {
    return String(b.mes).localeCompare(String(a.mes)) ||
      String(a.consultor).localeCompare(String(b.consultor), 'pt-BR') ||
      String(a.data).localeCompare(String(b.data)) ||
      String(a.hora_inicio).localeCompare(String(b.hora_inicio));
  });

  var values = items.map(function(item) {
    return viewHeaders.map(function(header) { return item[header] === undefined ? '' : item[header]; });
  });
  var viewSheet = getOrCreateSheet_(CMAX_DAILY_VIEW_SHEET);
  viewSheet.clearContents();
  viewSheet.getRange(1, 1, 1, viewHeaders.length).setValues([viewHeaders]);
  if (values.length) viewSheet.getRange(2, 1, values.length, viewHeaders.length).setValues(values);

  var ranges = {};
  var availableMonths = {};
  var availableActivities = {};
  var syncedAt = '';
  items.forEach(function(item, index) {
    if (!ranges[item.mes]) ranges[item.mes] = { start: index + 2, count: 0 };
    ranges[item.mes].count++;
    availableMonths[item.mes] = true;
    availableActivities[item.descricao || item.modalidade || item.tipo] = true;
    if (String(item.sincronizado_em || '') > syncedAt) syncedAt = String(item.sincronizado_em || '');
  });
  var meta = {
    ranges: ranges,
    all: { start: 2, count: items.length },
    available_months: cmaxRelevantMonths_(Object.keys(availableMonths)),
    history_loaded_months: cmaxSnapshotMonths_(),
    training_team_cutoff: trainingCutoff,
    training_team: Object.keys(trainingTeamMap).sort(),
    available_activities: Object.keys(availableActivities).sort(),
    month_hashes: monthHashes,
    synced_at: syncedAt,
    rebuilt_at: new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty('CMAX_VIEW_META_JSON', JSON.stringify(meta));
  writeCmaxMaterializedCaches_(items, meta);
}

function writeCmaxMaterializedCaches_(items, meta) {
  var byMonth = {};
  (items || []).forEach(function(item) {
    if (!item.mes) return;
    (byMonth[item.mes] || (byMonth[item.mes] = [])).push(item);
  });
  Object.keys(byMonth).forEach(function(month) {
    var events = byMonth[month];
    writeCompressedScriptCache_(cmaxViewCacheKey_({ month: month }), {
      ok: true,
      events: events,
      total: events.length,
      month: month,
      available_months: cmaxRelevantMonths_(meta.available_months),
      history_months: cmaxHistoryMonths_(),
      history_loaded_months: cmaxSnapshotMonths_(meta),
      training_team_cutoff: meta.training_team_cutoff || '',
      training_team: meta.training_team || [],
      available_activities: meta.available_activities || [],
      history_sync: getCmaxDailyHistoryStatus_().history_sync,
      synced_at: meta.synced_at || '',
      materialized: true,
      sheet: CMAX_DAILY_VIEW_SHEET
    }, CMAX_VIEW_CACHE_SECONDS);
  });
}

function sanitizeCmaxMonth_(value) {
  var match = String(value || '').trim().match(/^(\d{4})-(0[1-9]|1[0-2])$/);
  return match ? match[0] : '';
}

function normalizeCmaxSheetMonth_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var text = String(value || '').trim();
  var match = text.match(/^(\d{4})-(0[1-9]|1[0-2])/);
  return match ? match[1] + '-' + match[2] : '';
}

function normalizeCmaxSheetDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var parsed = parseCmaxDate_(value);
  return parsed ? Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(value || '');
}

function normalizeCmaxSheetTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    if (value.getUTCFullYear() <= 1900) return cmaxLegacyTimeFromHours_(value.getUTCHours(), value.getUTCMinutes());
    return Utilities.formatDate(value, 'America/Sao_Paulo', 'HH:mm');
  }
  var text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
    var legacy = text.match(/^(1899|1900)-\d{2}-\d{2}T(\d{2}):(\d{2})/);
    if (legacy) return cmaxLegacyTimeFromHours_(Number(legacy[2]), Number(legacy[3]));
    var parsed = new Date(text);
    if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, 'America/Sao_Paulo', 'HH:mm');
  }
  var iso = text.match(/T(\d{2}):(\d{2})/);
  if (iso) return iso[1] + ':' + iso[2];
  var time = text.match(/^(\d{1,2}):(\d{2})/);
  if (time) return String(Number(time[1])).padStart(2, '0') + ':' + time[2];
  return text;
}

function cmaxLegacyTimeFromHours_(hours, minutes) {
  var adjusted = (Number(hours || 0) - 8 + 24) % 24;
  return String(adjusted).padStart(2, '0') + ':' + String(Number(minutes || 0)).padStart(2, '0');
}

function normalizeCmaxModality_(value) {
  return sanitizeText_(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
}

function isCmaxDailyModality_(value) {
  var modality = normalizeCmaxModality_(value);
  return modality === 'TREINAMENTO ON LINE' ||
    modality === 'TREINAMENTO IN LOCO' ||
    modality === 'TREINAMENTO INTERNO' ||
    modality === 'TREINAMENTO ON LINE AVULSO' ||
    modality === 'TREINAMENTO IN LOCO AVULSO';
}

function isCmaxTrainingConsultant_(name) {
  var key = normalizeKey_(name);
  return key.indexOf('LAIS') < 0 && key.indexOf('EVELYN') < 0;
}

function cmaxMonthRange_(month) {
  var parts = month.split('-');
  var year = Number(parts[0]);
  var monthIndex = Number(parts[1]) - 1;
  var lastDay = new Date(year, monthIndex + 1, 0).getDate();
  return {
    start: year + '-' + String(monthIndex + 1).padStart(2, '0') + '-01',
    end: year + '-' + String(monthIndex + 1).padStart(2, '0') + '-' + String(lastDay).padStart(2, '0')
  };
}

function parseCmaxDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var text = String(value || '').trim();
  var br = text.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
  var iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var date = new Date(text);
  return isNaN(date.getTime()) ? null : date;
}

function isCmaxPositiveResult_(value) {
  var normalized = sanitizeText_(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
  return normalized === 'POSITIVO' || normalized === 'POSITIVE' || normalized === '1';
}

function buildProjectUrl_(mapping, tasks) {
  if (mapping.project_url) return mapping.project_url;
  if (mapping.view_id) return 'https://app.clickup.com/v/o/' + mapping.view_id;
  var firstTask = tasks && tasks[0];
  if (firstTask && firstTask.url) return firstTask.url;
  return '';
}
