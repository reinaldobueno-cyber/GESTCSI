import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const postStart = source.indexOf('function doPost(e)');
const postEnd = source.indexOf('\nfunction syncAllProjects(', postStart);
const verifyStart = source.indexOf('function verifyWebhookRequest_(e, body)');
const verifyEnd = source.indexOf('\nfunction getMonthSheet_(', verifyStart);
assert.ok(postStart >= 0 && postEnd > postStart && verifyStart >= 0 && verifyEnd > verifyStart);

function makeHarness(configuredToken) {
  let queued = 0;
  const context = {
    getScriptProperty_: () => configuredToken,
    enqueueDirtyEvent_: () => { queued += 1; return { queued: true }; },
    dispatchLegacyPostCommand_: (action) => ({ ok: true, action, channel: 'panel' }),
    simplifyErrorMessage_: (error) => error.message,
    jsonOutput_: (payload) => ({ body: JSON.stringify(payload) }),
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput(body) {
        return { body, setMimeType() { return this; } };
      }
    }
  };
  vm.runInNewContext(source.slice(postStart, postEnd) + source.slice(verifyStart, verifyEnd), context);
  return {
    send(token, body) {
      const output = context.doPost({
        parameter: { webhook_token: token },
        postData: { contents: body }
      });
      return JSON.parse(output.body);
    },
    queued: () => queued
  };
}

test('fails closed when no webhook token is configured', () => {
  const harness = makeHarness('');
  assert.equal(harness.send('', '{"event":"taskUpdated","task_id":"1"}').error, 'invalid_signature');
  assert.equal(harness.queued(), 0);
});

test('rejects wrong token and malformed payload without writing to the queue', () => {
  const harness = makeHarness('shared-secret');
  assert.equal(harness.send('wrong', '{"event":"taskUpdated"}').error, 'invalid_signature');
  assert.equal(harness.send('shared-secret', '{bad json').error, 'invalid_payload');
  assert.equal(harness.send('shared-secret', '{}').error, 'invalid_payload');
  assert.equal(harness.queued(), 0);
});

test('queues a valid ClickUp event with the configured token', () => {
  const harness = makeHarness('shared-secret');
  const response = harness.send('shared-secret', '{"event":"taskUpdated","task_id":"1"}');
  assert.equal(response.ok, true);
  assert.equal(response.event, 'taskUpdated');
  assert.equal(harness.queued(), 1);
});

test('dispatches a panel command separately from the ClickUp webhook queue', () => {
  const harness = makeHarness('shared-secret');
  const response = harness.send('', '{"action":"login","username":"user","password_sha":"hash"}');
  assert.equal(response.channel, 'panel');
  assert.equal(response.action, 'login');
  assert.equal(harness.queued(), 0);
});
