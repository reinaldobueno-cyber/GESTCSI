import test from 'node:test';
import assert from 'node:assert/strict';
import bridge from '../gateway/src/index.js';

const base = 'https://staging.gestcsi.example';
const disabled = { BRIDGE_ENABLED: '0' };

async function call(path, options = {}, env = disabled) {
  const response = await bridge.fetch(new Request(base + path, options), env);
  return { response, payload: await response.json() };
}

test('health is public but does not disclose bindings', async () => {
  const { response, payload } = await call('/api/v1/health');
  assert.equal(response.status, 200);
  assert.equal(payload.enabled, false);
  assert.equal(response.headers.get('Cache-Control'), 'no-store');
  assert.equal(JSON.stringify(payload).includes('APPS_SCRIPT_URL'), false);
});

test('sensitive routes fail closed without bindings', async () => {
  for (const [path, method] of [
    ['/api/v1/auth/login', 'POST'],
    ['/api/v1/auth/logout', 'POST'],
    ['/api/v1/projects/inventory', 'GET']
  ]) {
    const { response, payload } = await call(path, { method });
    assert.equal(response.status, 503);
    assert.equal(payload.error, 'bridge_not_configured');
  }
});

test('method policy and unknown commands are closed', async () => {
  assert.equal((await call('/api/v1/auth/login')).response.status, 405);
  assert.equal((await call('/api/v1/projects/inventory', { method: 'POST' })).response.status, 405);
  assert.equal((await call('/api/v1/syncProject', { method: 'POST' })).response.status, 404);
  assert.equal((await call('/api/v1/health', { method: 'POST' })).response.status, 405);
});

test('enabled bridge rejects missing session and foreign origins before upstream access', async () => {
  const env = {
    BRIDGE_ENABLED: '1', DB: { prepare() { throw new Error('DB must not be touched'); } },
    APPS_SCRIPT_URL: 'https://script.google.com/macros/s/example/exec',
    SESSION_ENCRYPTION_KEY: 'placeholder'
  };
  assert.equal((await call('/api/v1/projects/inventory', {}, env)).response.status, 401);
  const foreign = await call('/api/v1/auth/login', {
    method: 'POST', headers: { Origin: 'https://elsewhere.example', 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: 'tester', password_sha: 'a'.repeat(64) })
  }, env);
  assert.equal(foreign.response.status, 403);
});

test('static requests are delegated to the assets binding', async () => {
  const env = { ...disabled, ASSETS: { fetch(request) { return new Response(request.url, { status: 200 }); } } };
  const response = await bridge.fetch(new Request(base + '/index.html'), env);
  assert.equal(response.status, 200);
  assert.equal(await response.text(), base + '/index.html');
});
