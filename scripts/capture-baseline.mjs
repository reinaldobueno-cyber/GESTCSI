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

const SHEET_ID = '1fqvDJ6Xh_POzWGyap9UH2rHulF-2wSe_G80m3rYnBIU';
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxtbpwBZEHDHCtU7lovmFHQ6R-MDgff-CB-yAyH6DfRwo0SjR9WXU6B4EYrgCcza6Kj/exec';
const APP_VERSION = '2026-09-11-clickup-activity-v4';
const TARGETS = [
  { month: '2026-07', code: 'JUL' },
  { month: '2026-08', code: 'AGO' }
];

async function fetchText(url, attempts = 1) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      let requestUrl = url;
      for (let redirect = 0; redirect <= 5; redirect += 1) {
        const response = await fetch(requestUrl, {
          redirect: 'manual',
          headers: {
            accept: 'text/csv,application/javascript,application/json,text/plain,*/*',
            'user-agent': 'GESTCSI-Baseline/1.0'
          }
        });
        if (response.status >= 300 && response.status < 400 && response.headers.get('location')) {
          requestUrl = new URL(response.headers.get('location'), requestUrl).toString();
          continue;
        }
        if (!response.ok) throw new Error(`HTTP ${response.status} em ${new URL(requestUrl).hostname}`);
        return { text: await response.text(), status: response.status, attempts: attempt, redirects: redirect };
      }
      throw new Error(`Excesso de redirecionamentos em ${new URL(url).hostname}`);
    } catch (error) {
      lastError = error;
      if (attempt < attempts) await new Promise((resolve) => setTimeout(resolve, attempt * 750));
    }
  }
  throw lastError;
}

async function fetchCsv(name, query) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&${query}`;
  const response = await fetchText(url, 3);
  const rows = parseCsv(response.text);
  return {
    name,
    rows,
    evidence: {
      status: response.status,
      bytes: Buffer.byteLength(response.text, 'utf8'),
      rows: rows.length,
      columns: rows.length ? Object.keys(rows[0]).length : 0,
      sha256: sha256(response.text)
    }
  };
}

async function fetchCmax(month) {
  const callback = 'baselineProbe';
  const url = `${APPS_SCRIPT_URL}?action=getCmaxDailyEvents&month=${month}&callback=${callback}`;
  const response = await fetchText(url);
  const json = response.text.replace(new RegExp(`^${callback}\\(`), '').replace(/\);?\s*$/, '');
  const payload = JSON.parse(json);
  if (!payload.ok || !Array.isArray(payload.events)) throw new Error(`CMAX ${month} retornou resposta inválida`);
  return {
    month,
    payload,
    evidence: {
      status: response.status,
      bytes: Buffer.byteLength(response.text, 'utf8'),
      events: payload.events.length,
      sha256: sha256(response.text),
      synced_at: String(payload.synced_at || ''),
      materialized: Boolean(payload.materialized),
      attempts: response.attempts,
      redirects: response.redirects
    }
  };
}

const [portfolio, milestones, decisions, indications, compensation] = await Promise.all([
  fetchCsv('portfolio', 'gid=1667939534'),
  fetchCsv('milestones', 'gid=756684793'),
  fetchCsv('closing_decisions', 'gid=533338656'),
  fetchCsv('bonus_indications', 'sheet=BONUS_INDICACOES'),
  fetchCsv('consultant_compensation', 'sheet=CONSULTORES_REMUNERACAO')
]);
// O Apps Script é consultado em sequência: chamadas paralelas já produziram 404
// intermitente na captura, comportamento que a baseline precisa detectar.
const cmaxJune = await fetchCmax('2026-06');
const cmaxJuly = await fetchCmax('2026-07');
const cmaxAugust = await fetchCmax('2026-08');

const cmaxByMonth = new Map([[cmaxJuly.month, cmaxJuly], [cmaxAugust.month, cmaxAugust]]);
const competences = Object.fromEntries(TARGETS.map(({ month, code }) => {
  const cmaxSource = cmaxByMonth.get(month);
  const cmax = cmaxMetrics(cmaxSource.payload.events, compensation.rows, month, cmaxJune.payload.events);
  const milestone = milestoneMetrics(milestones.rows, month);
  const projectClosing = projectClosingMetrics(decisions.rows, month);
  const indication = indicationMetrics(indications.rows, month);
  const competence = {
    portfolio: portfolioMetrics(portfolio.rows, code),
    cmax,
    milestones: milestone,
    project_closing: projectClosing,
    indications: indication,
    financial: {
      daily_value: cmax.daily_value,
      milestone_value: milestone.milestone_value,
      project_value: projectClosing.project_value,
      indication_value: indication.indication_value,
      total_value: Number((cmax.daily_value + milestone.milestone_value + projectClosing.project_value + indication.indication_value).toFixed(4))
    }
  };
  competence.reconciliation = reconciliationFor(competence);
  return [month, competence];
}));

const sources = Object.fromEntries([
  portfolio, milestones, decisions, indications, compensation
].map((source) => [source.name, source.evidence]));
for (const source of [cmaxJune, cmaxJuly, cmaxAugust]) sources[`cmax_${source.month}`] = source.evidence;

const output = {
  schema_version: 1,
  captured_at: new Date().toISOString(),
  production_base: {
    git_sha: '11e7c2fdc9af2b1f36f1217ae90a061773f19490',
    app_version: APP_VERSION,
    branch: 'gestcsi/main'
  },
  privacy: 'Somente agregados, metadados técnicos e hashes; sem nomes, clientes ou registros individuais.',
  sources,
  competences,
  known_divergences: [
    'A suíte preserva o piso histórico de 201 projetos, enquanto a fonte materializada atual declara e entrega 234 linhas.',
    'Os hashes representam fontes mutáveis e mudarão quando a operação atualizar as planilhas ou o CMAX.',
    'A baseline usa as regras financeiras presentes no frontend do SHA-base; observações locais do navegador não fazem parte do cálculo.',
    'A leitura valida disponibilidade e coerência dos dados, mas não executa comandos de escrita nem sincronizações administrativas.',
    'Chamadas CMAX paralelas e uma repetição sequencial produziram HTTP 404 intermitente; o capturador usa redirecionamento explícito, sequência e retentativa controlada, mas a origem ainda precisa de monitoramento.'
  ]
};
output.aggregate_sha256 = sha256(JSON.stringify({ sources: output.sources, competences: output.competences }));
process.stdout.write(`${JSON.stringify(output, null, 2)}\n`);
