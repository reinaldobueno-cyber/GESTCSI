// E05/E09 staging bridge. Only login and the paged historical inventory are
// implemented; every mutation remains closed until session/CSRF/audit smoke.
const COOKIE_NAME = 'gestcsi_session';
const SESSION_TTL_MS = 5 * 60 * 60 * 1000; // Below Apps Script's six-hour cache TTL.
const LOGIN_WINDOW_MS = 15 * 60 * 1000;
const LOGIN_MAX_ATTEMPTS = 5;
const encoder = new TextEncoder();

function json(data, status = 200, extraHeaders = {}) {
  return Response.json(data, {
    status,
    headers: {
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
      ...extraHeaders
    }
  });
}

function bytesToBase64Url(bytes) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function base64UrlToBytes(value) {
  const normalized = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  const binary = atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, '='));
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

function randomToken() {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return bytesToBase64Url(bytes);
}

async function sha256(value) {
  return bytesToBase64Url(new Uint8Array(await crypto.subtle.digest('SHA-256', encoder.encode(value))));
}

async function encryptionKey(env) {
  const raw = base64UrlToBytes(env.SESSION_ENCRYPTION_KEY);
  if (raw.length !== 32) throw new Error('invalid_encryption_key');
  return crypto.subtle.importKey('raw', raw, 'AES-GCM', false, ['encrypt', 'decrypt']);
}

async function encryptToken(token, env) {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const encrypted = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await encryptionKey(env), encoder.encode(token));
  return bytesToBase64Url(iv) + '.' + bytesToBase64Url(new Uint8Array(encrypted));
}

async function decryptToken(value, env) {
  const parts = String(value || '').split('.');
  if (parts.length !== 2) throw new Error('invalid_session_cipher');
  const iv = base64UrlToBytes(parts[0]);
  if (iv.length !== 12) throw new Error('invalid_session_cipher');
  const clear = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, await encryptionKey(env), base64UrlToBytes(parts[1]));
  return new TextDecoder().decode(clear);
}

async function readBoundedJson(body, limit) {
  if (!body) throw new Error('empty_body');
  const reader = body.getReader();
  const chunks = [];
  let size = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > limit) throw new Error('body_too_large');
      chunks.push(value);
    }
  } catch (error) {
    await reader.cancel().catch(() => {});
    throw error;
  }
  const joined = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    joined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return JSON.parse(new TextDecoder().decode(joined));
}

function upstreamUrl(env) {
  const url = new URL(env.APPS_SCRIPT_URL);
  if (url.protocol !== 'https:' || url.hostname !== 'script.google.com' || !/^\/macros\/s\/[^/]+\/exec$/.test(url.pathname)) {
    throw new Error('invalid_upstream_url');
  }
  url.search = '';
  return url;
}

async function requestAppsScript(env, method, params, maxBytes, timeoutMs) {
  const url = upstreamUrl(env);
  let response;
  if (method === 'GET') {
    Object.keys(params).forEach((key) => url.searchParams.set(key, String(params[key])));
    response = await fetch(url, { method: 'GET', redirect: 'follow', signal: AbortSignal.timeout(timeoutMs) });
  } else {
    response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(params),
      redirect: 'follow',
      signal: AbortSignal.timeout(timeoutMs)
    });
  }
  const finalUrl = new URL(response.url);
  if (finalUrl.protocol !== 'https:' || !['script.google.com', 'script.googleusercontent.com'].includes(finalUrl.hostname)) {
    throw new Error('unexpected_upstream_redirect');
  }
  if (!response.ok) throw new Error('upstream_http_error');
  const payload = await readBoundedJson(response.body, maxBytes);
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('invalid_upstream_payload');
  return payload;
}

function exactOrigin(request) {
  return request.headers.get('Origin') === new URL(request.url).origin;
}

function sessionCookie(request) {
  const pairs = String(request.headers.get('Cookie') || '').split(';');
  const item = pairs.map((part) => part.trim()).find((part) => part.startsWith(COOKIE_NAME + '='));
  const value = item ? item.slice(COOKIE_NAME.length + 1) : '';
  return /^[A-Za-z0-9_-]{43}$/.test(value) ? value : '';
}

async function loadSession(request, env) {
  const id = sessionCookie(request);
  if (!id) return null;
  const key = await sha256(id);
  const row = await env.DB.prepare('SELECT token_cipher, user_json, csrf_hash, expires_at FROM sessions WHERE session_hash = ?')
    .bind(key).first();
  if (!row || Number(row.expires_at) <= Date.now()) return null;
  return { key, row, token: await decryptToken(row.token_cipher, env), user: JSON.parse(row.user_json) };
}

async function checkLoginLimit(request, env, username) {
  const clientIp = request.headers.get('CF-Connecting-IP') || 'unknown';
  // Limit both one account and username spraying from one IP.
  const accountKey = await sha256('account|' + username.toLowerCase() + '|' + clientIp);
  const ipKey = await sha256('ip|' + clientIp);
  const now = Date.now();
  const cutoff = now - LOGIN_WINDOW_MS;
  const keys = [accountKey, ipKey];
  for (const key of keys) {
    await env.DB.prepare(`INSERT INTO login_attempts (attempt_key, window_start, attempts) VALUES (?, ?, 1)
      ON CONFLICT(attempt_key) DO UPDATE SET
        attempts = CASE WHEN window_start <= ? THEN 1 ELSE attempts + 1 END,
        window_start = CASE WHEN window_start <= ? THEN excluded.window_start ELSE window_start END`)
      .bind(key, now, cutoff, cutoff).run();
  }
  const records = await Promise.all(keys.map((key) => env.DB.prepare('SELECT attempts FROM login_attempts WHERE attempt_key = ?').bind(key).first()));
  return { accountKey, allowed: records.every((record) => record && Number(record.attempts) <= LOGIN_MAX_ATTEMPTS) };
}

async function login(request, env) {
  if (!exactOrigin(request)) return json({ ok: false, error: 'origin_rejected' }, 403);
  if (!/^application\/json\b/i.test(request.headers.get('Content-Type') || '')) return json({ ok: false, error: 'json_required' }, 415);
  let body;
  try { body = await readBoundedJson(request.body, 2048); } catch (_) { return json({ ok: false, error: 'invalid_payload' }, 400); }
  const username = String(body && body.username || '').trim().toLowerCase();
  const passwordSha = String(body && body.password_sha || '').trim().toLowerCase();
  if (!/^[a-z0-9._@-]{1,120}$/.test(username) || !/^[a-f0-9]{64}$/.test(passwordSha)) {
    return json({ ok: false, error: 'invalid_credentials_format' }, 400);
  }
  const limit = await checkLoginLimit(request, env, username);
  if (!limit.allowed) return json({ ok: false, error: 'too_many_attempts' }, 429);
  const result = await requestAppsScript(env, 'POST', { action: 'login', username, password_sha: passwordSha }, 8192, 15000);
  if (result.ok !== true || !/^[A-Za-z0-9_-]{20,}$/.test(String(result.token || '')) ||
      !result.user || !result.user.enabled || !['admin', 'coordenador', 'user'].includes(result.user.role)) {
    return json({ ok: false, error: 'login_rejected' }, 401);
  }
  const sessionId = randomToken();
  const csrf = randomToken();
  const now = Date.now();
  await env.DB.prepare(`INSERT INTO sessions (session_hash, token_cipher, user_json, csrf_hash, created_at, expires_at)
    VALUES (?, ?, ?, ?, ?, ?)`)
    .bind(await sha256(sessionId), await encryptToken(result.token, env), JSON.stringify(result.user), await sha256(csrf), now, now + SESSION_TTL_MS)
    .run();
  await env.DB.prepare('DELETE FROM login_attempts WHERE attempt_key = ?').bind(limit.accountKey).run();
  return json({ ok: true, user: result.user, csrf_token: csrf }, 200, {
    'Set-Cookie': `${COOKIE_NAME}=${sessionId}; HttpOnly; Secure; SameSite=Strict; Path=/api/v1; Max-Age=${Math.floor(SESSION_TTL_MS / 1000)}`
  });
}

async function verifyCsrf(request, session) {
  if (!exactOrigin(request)) return false;
  const provided = String(request.headers.get('X-GESTCSI-CSRF') || '');
  if (!/^[A-Za-z0-9_-]{43}$/.test(provided)) return false;
  const [a, b] = [base64UrlToBytes(await sha256(provided)), base64UrlToBytes(session.row.csrf_hash)];
  if (a.length !== 32 || b.length !== 32) return false;
  let difference = 0;
  for (let index = 0; index < 32; index += 1) difference |= a[index] ^ b[index];
  return difference === 0;
}

async function logout(request, env) {
  const session = await loadSession(request, env);
  if (!session) return json({ ok: false, error: 'unauthorized' }, 401);
  if (!(await verifyCsrf(request, session))) return json({ ok: false, error: 'csrf_rejected' }, 403);
  await env.DB.prepare('DELETE FROM sessions WHERE session_hash = ?').bind(session.key).run();
  return json({ ok: true }, 200, {
    'Set-Cookie': `${COOKIE_NAME}=; HttpOnly; Secure; SameSite=Strict; Path=/api/v1; Max-Age=0`
  });
}

async function inventory(request, env) {
  const session = await loadSession(request, env);
  if (!session) return json({ ok: false, error: 'unauthorized' }, 401);
  const url = new URL(request.url);
  const offset = Number(url.searchParams.get('offset') || 0);
  const limit = Number(url.searchParams.get('limit') || 100);
  if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(limit) || limit < 25 || limit > 100) {
    return json({ ok: false, error: 'invalid_pagination' }, 400);
  }
  const result = await requestAppsScript(env, 'GET', {
    action: 'getClickUpInventory', auth_token: session.token, lean: '1', paged: '1', offset, limit
  }, 1024 * 1024, 45000);
  if (result.ok !== true || !Array.isArray(result.projetos) || result.paged !== true ||
      !Number.isSafeInteger(Number(result.next_offset)) || Number(result.next_offset) < offset) {
    return json({ ok: false, source_unavailable: true, error: 'historical_inventory_unavailable' }, 502);
  }
  return json({ ...result, source: 'same-origin-bridge' });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (!url.pathname.startsWith('/api/')) {
      return env.ASSETS ? env.ASSETS.fetch(request) : json({ ok: false, error: 'assets_unavailable' }, 503);
    }
    if (url.pathname === '/api/v1/health') {
      if (request.method !== 'GET') return json({ ok: false, error: 'method_not_allowed' }, 405);
      return json({ ok: true, service: 'gestcsi-e05-bridge', enabled: env.BRIDGE_ENABLED === '1', schema_version: 1 });
    }
    const methods = {
      '/api/v1/auth/login': 'POST',
      '/api/v1/auth/logout': 'POST',
      '/api/v1/projects/inventory': 'GET'
    };
    const required = methods[url.pathname];
    if (!required) return json({ ok: false, error: 'not_found' }, 404);
    if (request.method !== required) return json({ ok: false, error: 'method_not_allowed', required_method: required }, 405);
    if (env.BRIDGE_ENABLED !== '1' || !env.DB || !env.APPS_SCRIPT_URL || !env.SESSION_ENCRYPTION_KEY) {
      return json({ ok: false, error: 'bridge_not_configured' }, 503);
    }
    try {
      if (url.pathname === '/api/v1/auth/login') return await login(request, env);
      if (url.pathname === '/api/v1/auth/logout') return await logout(request, env);
      return await inventory(request, env);
    } catch (error) {
      // Do not log request bodies, credentials, cookies, Apps Script tokens or URLs.
      console.error(JSON.stringify({ event: 'bridge_request_failed', route: url.pathname, category: error instanceof Error ? error.name : 'unknown' }));
      return json({ ok: false, error: 'upstream_unavailable' }, 502);
    }
  }
};
