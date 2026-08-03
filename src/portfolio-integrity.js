(function(root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.PortfolioIntegrity = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';

  var DEFAULT_MONTHS = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

  function normalizeText(value) {
    return String(value || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]+/g, ' ')
      .replace(/\s+/g, ' ').trim();
  }

  function normalizeMonth(value) {
    return normalizeText(value).slice(0, 3);
  }

  function projectIdentity(project) {
    if (!project || typeof project !== 'object') return '';
    var explicit = String(project.project_key || project.projectKey || project._carteira_key || '').trim();
    if (explicit) return explicit.toUpperCase();
    var month = normalizeMonth(project.mes || project.mes_origem);
    var client = normalizeText(project.cliente || project.nome_cliente);
    var row = String(project._sheet_row || project.row_number || '').trim();
    return month && client && row ? [month, client, row].join('|') : '';
  }

  function buildManifest(projects, months) {
    var list = Array.isArray(projects) ? projects : [];
    var monthList = Array.isArray(months) && months.length ? months.map(normalizeMonth) : DEFAULT_MONTHS.slice();
    var byMonth = {};
    var identities = {};
    var duplicateKeys = [];
    var missingIdentity = 0;
    monthList.forEach(function(month) { byMonth[month] = 0; });
    list.forEach(function(project) {
      var month = normalizeMonth(project && (project.mes || project.mes_origem));
      if (month) byMonth[month] = Number(byMonth[month] || 0) + 1;
      var identity = projectIdentity(project);
      if (!identity) {
        missingIdentity += 1;
      } else if (identities[identity]) {
        if (duplicateKeys.indexOf(identity) < 0) duplicateKeys.push(identity);
      } else {
        identities[identity] = true;
      }
    });
    return {
      total: list.length,
      byMonth: byMonth,
      duplicateKeys: duplicateKeys,
      missingIdentity: missingIdentity
    };
  }

  function validateSnapshot(payload, options) {
    payload = payload || {};
    options = options || {};
    var projects = Array.isArray(payload.projects) ? payload.projects : payload.projetos;
    var months = options.months || DEFAULT_MONTHS;
    var manifest = buildManifest(projects, months);
    var errors = [];
    var declaredTotal = Number(payload.total);
    var declaredByMonth = payload.byMonth || payload.projetos_por_mes || {};
    var minimumTotal = Number(options.minimumTotal || 0);

    if (!Array.isArray(projects)) errors.push('projects_not_array');
    if (manifest.total < minimumTotal) errors.push('below_minimum_total');
    if (options.requireDeclaredTotal && !Number.isFinite(declaredTotal)) errors.push('declared_total_missing');
    if (Number.isFinite(declaredTotal) && declaredTotal !== manifest.total) errors.push('declared_total_mismatch');

    Object.keys(declaredByMonth || {}).forEach(function(rawMonth) {
      var month = normalizeMonth(rawMonth);
      var declared = Number(declaredByMonth[rawMonth]);
      if (Number.isFinite(declared) && Number(manifest.byMonth[month] || 0) !== declared) {
        errors.push('month_total_mismatch:' + month);
      }
    });

    if (manifest.duplicateKeys.length) errors.push('duplicate_project_keys');
    if (Array.isArray(payload.failedMonths) && payload.failedMonths.length) errors.push('failed_months');

    return {
      ok: errors.length === 0,
      errors: errors,
      manifest: manifest
    };
  }

  function shouldReplaceSnapshot(candidatePayload, currentProjects, options) {
    var validation = validateSnapshot(candidatePayload, options);
    var currentTotal = Array.isArray(currentProjects) ? currentProjects.length : 0;
    if (!validation.ok) return Object.assign({ replace: false, reason: 'invalid_candidate' }, validation);
    if (currentTotal && validation.manifest.total < currentTotal) {
      return Object.assign({ replace: false, reason: 'candidate_smaller_than_current' }, validation);
    }
    return Object.assign({ replace: true, reason: 'validated_candidate' }, validation);
  }

  return {
    DEFAULT_MONTHS: DEFAULT_MONTHS,
    normalizeMonth: normalizeMonth,
    projectIdentity: projectIdentity,
    buildManifest: buildManifest,
    validateSnapshot: validateSnapshot,
    shouldReplaceSnapshot: shouldReplaceSnapshot
  };
});
