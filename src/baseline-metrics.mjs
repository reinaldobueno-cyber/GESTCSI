import { createHash } from 'node:crypto';

export function sha256(value) {
  return createHash('sha256').update(String(value ?? ''), 'utf8').digest('hex');
}

export function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const input = String(text || '');
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index];
    if (quoted) {
      if (char === '"' && input[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ',') {
      row.push(field);
      field = '';
    } else if (char === '\n') {
      row.push(field.replace(/\r$/, ''));
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }
  if (field || row.length) {
    row.push(field.replace(/\r$/, ''));
    rows.push(row);
  }
  if (!rows.length) return [];
  const headers = rows.shift().map((header) => header.trim());
  while (headers.length && !headers.at(-1)) headers.pop();
  return rows
    .filter((cells) => cells.some((cell) => cell !== ''))
    .map((cells) => Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ''])));
}

export function normalizeKey(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

export function canonicalConsultantName(value) {
  const original = String(value || '').trim();
  if (original.includes(',')) {
    const parts = original.split(/\s*,\s*/).filter(Boolean);
    const selected = parts.find((part) => !/ADMINISTRATIVO|ADMINISTRADOR|MULTSOFT|SUPORTE/.test(normalizeKey(part)));
    return canonicalConsultantName(selected || parts[0] || original);
  }
  const key = normalizeKey(original);
  if (key === 'YASMIN.SOUSA@MULTIBOVINOS.COM.BR') return 'Yasmin Sousa';
  if (key === 'ANITA' || key === 'ANITA CRISTINA RODRIGUES TAVARES') return 'Anita Cristina Rodrigues Tavares';
  if (key === 'WAND' || key.includes('WANDERLEY') || key.includes('WANDERLEI')) return 'Wanderley Cabral';
  if (key === 'LUCAS' || key === 'LUCAS PEREIRA DA SILVA') return 'Lucas Pereira da Silva';
  if (key === 'SERGIO' || key === 'SERGIO CASTRO') return 'SÉRGIO';
  if (original.includes('@')) {
    return original.split('@')[0].split(/[._-]+/).filter(Boolean)
      .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()).join(' ');
  }
  return original;
}

export function consultantKey(value) {
  return normalizeKey(canonicalConsultantName(value));
}

function consultantTokens(value) {
  return consultantKey(value).split(/\s+/).filter((token) => token.length > 2);
}

function compensationFor(name, compensationRows) {
  const key = consultantKey(name);
  const exact = compensationRows.find((row) => consultantKey(row.consultant_key || row.consultant_name) === key);
  if (exact?.seniority) return exact;
  const tokens = consultantTokens(name);
  const firstMatches = compensationRows.filter((row) => {
    const rowTokens = consultantTokens(row.consultant_key || row.consultant_name);
    return tokens.length && rowTokens.length && tokens[0] === rowTokens[0];
  });
  let best = null;
  let bestScore = 0;
  for (const row of firstMatches) {
    const rowTokens = consultantTokens(row.consultant_key || row.consultant_name);
    const score = tokens.filter((token) => rowTokens.includes(token)).length;
    if (score >= 2 && score > bestScore) {
      best = row;
      bestScore = score;
    }
  }
  if (best?.seniority) return best;
  if (firstMatches.length === 1 && firstMatches[0].seniority) return firstMatches[0];
  if (tokens[0] === 'GUILHERME') return { seniority: 'master', daily_value: 220 };
  return { seniority: '', daily_value: 0 };
}

function cmaxTime(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const legacy = text.match(/^(1899|1900)-\d{2}-\d{2}T(\d{2}):(\d{2})/);
  if (legacy) return `${String((Number(legacy[2]) - 8 + 24) % 24).padStart(2, '0')}:${legacy[3]}`;
  if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
      return new Intl.DateTimeFormat('pt-BR', {
        timeZone: 'America/Sao_Paulo', hour: '2-digit', minute: '2-digit', hour12: false
      }).format(parsed);
    }
  }
  const iso = text.match(/T(\d{2}):(\d{2})/);
  if (iso) return `${iso[1]}:${iso[2]}`;
  const time = text.match(/^(\d{1,2}):(\d{2})/);
  return time ? `${String(Number(time[1])).padStart(2, '0')}:${time[2]}` : '';
}

function timeMinutes(value) {
  const parts = cmaxTime(value).split(':');
  return parts.length === 2 ? Number(parts[0]) * 60 + Number(parts[1]) : null;
}

function isPositive(item) {
  const result = normalizeKey(item?.resultado || 'SEM RESULTADO');
  return result === 'POSITIVO' || result === 'POSITIVE' || result === '1';
}

function countsDaily(item) {
  if (!isPositive(item) || consultantKey(item?.consultor).includes('LAIS')) return false;
  const mode = normalizeKey(item?.modalidade || item?.descricao || item?.tipo);
  return [
    'TREINAMENTO ON LINE', 'TREINAMENTO IN LOCO', 'TREINAMENTO INTERNO',
    'TREINAMENTO ON LINE AVULSO', 'TREINAMENTO IN LOCO AVULSO'
  ].includes(mode);
}

function eventEquivalent(item) {
  if (!countsDaily(item)) return 0;
  const start = timeMinutes(item.hora_inicio);
  const end = timeMinutes(item.hora_fim);
  if (start === null || end === null) return 1;
  let duration = end - start;
  if (duration <= 0) duration += 24 * 60;
  if (duration > 240) return 1;
  if (duration >= 180) return 0.5;
  return Math.round((duration / 360) * 10000) / 10000;
}

function isoDate(year, month, day) {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function easterDate(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(year, month - 1, day));
}

function addUtcDays(date, days) {
  const next = new Date(date);
  next.setUTCDate(next.getUTCDate() + days);
  return next;
}

function specialMultiplier(value) {
  const match = String(value || '').slice(0, 10).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return 1;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCDay() === 0 || date.getUTCDay() === 6) return 1.5;
  const monthDay = `${match[2]}-${match[3]}`;
  if (['01-01', '04-21', '05-01', '09-07', '10-12', '11-02', '11-15', '11-20', '12-25', '05-24', '10-24'].includes(monthDay)) return 1.5;
  const easter = easterDate(year);
  const dynamicDates = [addUtcDays(easter, -2), addUtcDays(easter, 60)]
    .map((item) => isoDate(item.getUTCFullYear(), item.getUTCMonth() + 1, item.getUTCDate()));
  return dynamicDates.includes(`${match[1]}-${match[2]}-${match[3]}`) ? 1.5 : 1;
}

export function cmaxMetrics(events, compensationRows, month, carryoverEvents = []) {
  const carryoverKeys = new Set(['29870', '29871', '29872']);
  const pool = [...events, ...carryoverEvents].filter((item) => {
    const key = String(item.event_key || item.event_id || '');
    const eventMonth = String(item.data || item.mes || '').slice(0, 7);
    return eventMonth === month || (month === '2026-07' && carryoverKeys.has(key));
  });
  const unique = new Map();
  for (const item of pool) {
    const key = String(item.event_key || item.event_id || [item.data, item.consultor, item.cliente, item.hora_inicio].join('|'));
    if (!unique.has(key)) unique.set(key, item);
  }
  const filtered = [...unique.values()];
  const eligible = filtered.filter(countsDaily).sort((a, b) => (
    String(a.consultor || '').localeCompare(String(b.consultor || ''), 'pt-BR') ||
    String(a.data || '').localeCompare(String(b.data || '')) ||
    String(a.hora_inicio || '').localeCompare(String(b.hora_inicio || ''))
  ));
  const used = new Map();
  const consultantTotals = new Map();
  let equivalent = 0;
  let specialEquivalent = 0;
  for (const item of eligible) {
    const person = consultantKey(item.consultor);
    if (person === 'THIAGO RIBEIRO') continue;
    const date = String(item.data || '').slice(0, 10);
    const dayKey = `${person}|${date}`;
    const available = Math.max(0, 1 - (used.get(dayKey) || 0));
    const allocated = Math.min(available, eventEquivalent(item));
    used.set(dayKey, (used.get(dayKey) || 0) + allocated);
    equivalent += allocated;
    const multiplier = specialMultiplier(date);
    if (multiplier > 1) specialEquivalent += allocated;
    const current = consultantTotals.get(person) || { equivalent: 0, value: 0, specialEquivalent: 0, name: item.consultor };
    const comp = compensationFor(item.consultor, compensationRows);
    current.equivalent += allocated;
    current.specialEquivalent += multiplier > 1 ? allocated : 0;
    current.value += allocated * Number(comp.daily_value || 0) * multiplier;
    current.rate = Number(comp.daily_value || 0);
    consultantTotals.set(person, current);
  }
  const consultants = [...consultantTotals.values()];
  return {
    source_events: events.length,
    unique_events_in_payment_pool: filtered.length,
    eligible_events: eligible.length,
    payable_equivalent: round(equivalent),
    special_equivalent: round(specialEquivalent),
    consultants_with_eligible_activity: consultants.length,
    consultants_without_rate: consultants.filter((item) => item.equivalent > 0 && !item.rate).length,
    daily_value: round(consultants.reduce((sum, item) => sum + item.value, 0))
  };
}

function round(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 10000) / 10000;
}

function statusCounts(rows) {
  return rows.reduce((counts, row) => {
    const key = String(row.situacao || 'outro').toLowerCase() || 'outro';
    counts[key] = (counts[key] || 0) + 1;
    return counts;
  }, {});
}

function monthFromDate(value, previousIfFirstDay = false) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return '';
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  if (previousIfFirstDay && Number(match[3]) === 1) date.setUTCMonth(date.getUTCMonth() - 1);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

function isProjectClosing(item) {
  if (normalizeKey(item.fechamento_projeto) === 'SIM') return true;
  const type = normalizeKey(item.item_tipo);
  if (type === 'FECHAMENTO DE PROJETO' || type === 'ENTREGA DE PROJETO') return true;
  const marker = normalizeKey(item.marcador_entrega || item.item_tipo);
  const name = normalizeKey(item.marco);
  const phase = normalizeKey(item.fase);
  const isBreakOff = phase.includes('FASE 8') || phase.includes('BREAK OFF') || name.includes('FASE 8') || name.includes('BREAK OFF');
  const hasDelivery = marker.includes('ENTREGA') || name.includes('ENTREGA DO PROJETO') || name.includes('KICKOFF DE ENTREGA');
  return isBreakOff && hasDelivery;
}

function rowKey(item) {
  return String(item.task_id || item.link || [item.project_key, item.projeto, item.marco, item.closed_at].join('|'));
}

export function milestoneMetrics(rows, month) {
  const milestones = rows.filter((row) => !isProjectClosing(row));
  const closed = milestones.filter((row) => String(row.mes_fechamento || '').slice(0, 7) === month);
  const approvedSeen = new Set();
  const approvedForPayment = milestones.filter((row) => {
    if (String(row.situacao || '').toLowerCase() !== 'aprovado') return false;
    const effectiveMonth = monthFromDate(row.validation_at, true) || String(row.mes_validacao || row.mes_fechamento || '').slice(0, 7);
    if (effectiveMonth !== month) return false;
    const key = rowKey(row);
    if (approvedSeen.has(key)) return false;
    approvedSeen.add(key);
    return true;
  });
  const rejectedForPayment = milestones.filter((row) => {
    if (String(row.situacao || '').toLowerCase() !== 'reprovado') return false;
    const effectiveMonth = String(row.mes_validacao || '').slice(0, 7) || monthFromDate(row.validation_at) || String(row.mes_fechamento || '').slice(0, 7);
    return effectiveMonth === month;
  });
  return {
    closed_rows: closed.length,
    closed_status: statusCounts(closed),
    approved_for_payment: approvedForPayment.length,
    rejected_for_payment: rejectedForPayment.length,
    milestone_value: round(approvedForPayment.reduce((sum, row) => sum + Number(row.valor_bonus || 30), 0))
  };
}

export function projectClosingMetrics(rows, month) {
  const selected = rows.filter((row) => String(row.month || '').slice(0, 7) === month);
  const approved = selected.filter((row) => String(row.decision || '').toLowerCase() === 'approved');
  const rejected = selected.filter((row) => String(row.decision || '').toLowerCase() === 'rejected');
  return {
    total: selected.length,
    approved: approved.length,
    rejected: rejected.length,
    project_value: round(approved.reduce((sum, row) => sum + Number(row.bonus_value || 80), 0))
  };
}

export function indicationMetrics(rows, month) {
  const selected = rows.filter((row) => {
    const saleMonth = monthFromDate(row.sale_date) || String(row.month || '').slice(0, 7);
    return saleMonth === month;
  });
  return {
    total: selected.length,
    indication_value: round(selected.reduce((sum, row) => sum + Number(row.value || 0), 0))
  };
}

export function portfolioMetrics(rows, monthCode) {
  const selected = rows.filter((row) => normalizeKey(row.mes) === normalizeKey(monthCode));
  const declared = [...new Set(rows.map((row) => Number(row.snapshot_total || 0)).filter(Boolean))];
  return {
    total_rows: rows.length,
    declared_snapshot_totals: declared,
    competence_rows: selected.length,
    distinct_source_row_numbers: new Set(rows.map((row) => String(row.source_sheet_row || '')).filter(Boolean)).size
  };
}

export function reconciliationFor(competence) {
  const bonus = competence.financial;
  const recomputed = round(bonus.daily_value + bonus.milestone_value + bonus.project_value + bonus.indication_value);
  return {
    component_sum_matches_total: recomputed === bonus.total_value,
    component_sum: recomputed,
    portfolio_declared_total_matches_rows: competence.portfolio.declared_snapshot_totals.length === 1 && competence.portfolio.declared_snapshot_totals[0] === competence.portfolio.total_rows,
    cmax_has_no_missing_rates: competence.cmax.consultants_without_rate === 0
  };
}
