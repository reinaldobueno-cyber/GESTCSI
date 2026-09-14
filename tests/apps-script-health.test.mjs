import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import {
  checkHealthOnce,
  monitorHealth,
  EXPECTED_APPS_SCRIPT_VERSION
} from '../scripts/check-apps-script-health.mjs';

function response(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { 'content-type': 'application/json' }
  });
}

function healthy(overrides = {}) {
  return {
    ok: true,
    service: 'gestcsi-apps-script',
    status: 'ready',
    schema_version: 1,
    version: EXPECTED_APPS_SCRIPT_VERSION,
    checked_at: '2026-09-14T12:00:00.000Z',
    visibility: 'public',
    ...overrides
  };
}

test('accepts the public health contract without credentials or business data', async () => {
  let requested;
  const result = await checkHealthOnce({
    endpoint: 'https://example.test/exec',
    fetchImpl: async (url) => {
      requested = url;
      return response(healthy());
    }
  });
  assert.equal(requested.searchParams.get('action'), 'health');
  assert.equal(requested.searchParams.has('auth_token'), false);
  assert.equal(result.visibility, 'public');
  assert.equal(result.version, EXPECTED_APPS_SCRIPT_VERSION);
});

test('requests protected detail when an authentication token is provided', async () => {
  let requested;
  const result = await checkHealthOnce({
    endpoint: 'https://example.test/exec',
    authToken: 'test-session-token',
    fetchImpl: async (url) => {
      requested = url;
      return response(healthy({ visibility: 'authenticated', dependencies: {} }));
    }
  });
  assert.equal(requested.searchParams.get('detail'), '1');
  assert.equal(requested.searchParams.get('auth_token'), 'test-session-token');
  assert.equal(result.visibility, 'authenticated');
});

test('classifies a request that exceeds the timeout', async () => {
  const fetchImpl = (_url, options) => new Promise((_resolve, reject) => {
    options.signal.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')));
  });
  await assert.rejects(
    checkHealthOnce({ endpoint: 'https://example.test/exec', fetchImpl, timeoutMs: 5 }),
    /timeout_after_5ms/
  );
});

test('rejects a degraded deployment so consecutive probes can raise an incident', async () => {
  await assert.rejects(
    checkHealthOnce({
      endpoint: 'https://example.test/exec',
      fetchImpl: async () => response(healthy({ status: 'degraded' }))
    }),
    /service_degraded/
  );
});

test('alerts once only after three consecutive failures such as a removed deployment', async () => {
  let requests = 0;
  const alerts = [];
  const result = await monitorHealth({
    endpoint: 'https://example.test/removed',
    attempts: 3,
    retryDelayMs: 0,
    fetchImpl: async () => {
      requests += 1;
      return response({ error: 'not_found' }, 404);
    },
    alert: async (incident) => alerts.push(incident)
  });
  assert.equal(requests, 3);
  assert.equal(result.ok, false);
  assert.equal(result.consecutive_failures, 3);
  assert.equal(alerts.length, 1);
});

test('keeps the Apps Script public payload minimal and protects diagnostic detail', async () => {
  const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
  assert.match(source, /action === 'health'/);
  assert.match(source, /service: 'gestcsi-apps-script'/);
  assert.match(source, /if \(!detailed\) return payload;\s+requireAdmin_\(params\);/);
  assert.doesNotMatch(source, /payload\.(?:token|password|credential)\s*=/i);
});

test('schedules an external probe and opens an incident only after the monitor fails', async () => {
  const workflow = await readFile(new URL('../.github/workflows/apps-script-health.yml', import.meta.url), 'utf8');
  assert.match(workflow, /cron: '\*\/15 \* \* \* \*'/);
  assert.match(workflow, /GESTCSI_HEALTH_ATTEMPTS: '3'/);
  assert.match(workflow, /if: steps\.probe\.outcome == 'failure'/);
  assert.match(workflow, /github\.rest\.issues\.create/);
});
