import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';

test('serves the production entrypoint and integrity module over HTTP', { timeout: 15000 }, async (t) => {
  const port = 4197;
  const server = spawn(process.execPath, ['scripts/serve.mjs'], {
    cwd: new URL('../', import.meta.url).pathname.replace(/^\/(.:)/, '$1'),
    env: { ...process.env, PORT: String(port) },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  t.after(() => server.kill());
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('preview não iniciou')), 5000);
    server.once('error', reject);
    server.stdout.on('data', (chunk) => {
      if (!String(chunk).includes(`127.0.0.1:${port}`)) return;
      clearTimeout(timer);
      resolve();
    });
  });
  const root = await fetch(`http://127.0.0.1:${port}/`);
  const html = await root.text();
  assert.equal(root.status, 200);
  assert.match(html, /PANEL_APP_VERSION\s*=\s*'2026-09-11-clickup-activity-v4'/);
  assert.match(html, /src="src\/portfolio-integrity\.js"/);

  const integrity = await fetch(`http://127.0.0.1:${port}/src/portfolio-integrity.js`);
  assert.equal(integrity.status, 200);
  assert.match(await integrity.text(), /root\.PortfolioIntegrity = api/);

  const missing = await fetch(`http://127.0.0.1:${port}/__missing__`);
  assert.equal(missing.status, 404);
});
