import { pathToFileURL } from 'node:url';

export const DEFAULT_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxtbpwBZEHDHCtU7lovmFHQ6R-MDgff-CB-yAyH6DfRwo0SjR9WXU6B4EYrgCcza6Kj/exec';
export const EXPECTED_APPS_SCRIPT_VERSION = '2026-09-15-health-v1';

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function checkHealthOnce(options = {}) {
  const endpoint = options.endpoint || DEFAULT_APPS_SCRIPT_URL;
  const expectedVersion = options.expectedVersion || EXPECTED_APPS_SCRIPT_VERSION;
  const timeoutMs = options.timeoutMs ?? 10_000;
  const fetchImpl = options.fetchImpl || fetch;
  const url = new URL(endpoint);
  url.searchParams.set('action', 'health');

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetchImpl(url, {
      redirect: 'follow',
      signal: controller.signal,
      headers: { accept: 'application/json' }
    });
    if (!response.ok) throw new Error(`http_${response.status}`);

    let payload;
    try {
      payload = await response.json();
    } catch {
      throw new Error('invalid_json');
    }

    if (payload.ok !== true) throw new Error('service_not_ok');
    if (payload.service !== 'gestcsi-apps-script') throw new Error('unexpected_service');
    if (payload.version !== expectedVersion) throw new Error(`unexpected_version:${payload.version || 'missing'}`);
    if (payload.status !== 'ready') throw new Error(`service_${payload.status || 'invalid_status'}`);

    return {
      ok: true,
      http: response.status,
      service: payload.service,
      status: payload.status,
      version: payload.version,
      checked_at: payload.checked_at,
      visibility: payload.visibility
    };
  } catch (error) {
    if (error?.name === 'AbortError') throw new Error(`timeout_after_${timeoutMs}ms`);
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

export async function monitorHealth(options = {}) {
  const attempts = Math.max(1, options.attempts ?? 3);
  const retryDelayMs = Math.max(0, options.retryDelayMs ?? 5_000);
  const failures = [];

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const result = await checkHealthOnce(options);
      return { ...result, attempt, consecutive_failures: 0 };
    } catch (error) {
      failures.push({ attempt, error: error?.message || String(error) });
      if (attempt < attempts && retryDelayMs) await wait(retryDelayMs);
    }
  }

  return {
    ok: false,
    service: 'gestcsi-apps-script',
    consecutive_failures: failures.length,
    failures,
    detected_at: new Date().toISOString()
  };
}

async function runCli() {
  const result = await monitorHealth({
    endpoint: process.env.GESTCSI_APPS_SCRIPT_URL || DEFAULT_APPS_SCRIPT_URL,
    expectedVersion: process.env.GESTCSI_APPS_SCRIPT_VERSION || EXPECTED_APPS_SCRIPT_VERSION,
    attempts: Number(process.env.GESTCSI_HEALTH_ATTEMPTS || 3),
    timeoutMs: Number(process.env.GESTCSI_HEALTH_TIMEOUT_MS || 10_000),
    retryDelayMs: Number(process.env.GESTCSI_HEALTH_RETRY_MS || 5_000)
  });
  const output = JSON.stringify(result);
  if (result.ok) console.log(output);
  else {
    console.error(output);
    process.exitCode = 1;
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await runCli();
}
