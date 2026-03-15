import assert from 'node:assert/strict';
import test from 'node:test';

import { buildAuthenticatedHealthPayload, healthHandler } from '../src/handlers/healthHandler.js';

test('authenticated health payload exposes cors and rate-limit diagnostics', () => {
  const payload = buildAuthenticatedHealthPayload('2026-03-07T12:00:00.000Z');

  assert.equal(payload.timestamp, '2026-03-07T12:00:00.000Z');
  assert.equal(payload.rateLimit.mode, 'memory');
  assert.equal(typeof payload.rateLimit.perMinute, 'number');
  assert.equal(typeof payload.rateLimit.perDay, 'number');
  assert.ok(payload.cors.mode === 'allowlist' || payload.cors.mode === 'permissive-local');
  assert.equal(typeof payload.cors.allowedOrigins, 'number');
});

test('health handler includes additive duration diagnostics', async () => {
  const response = await healthHandler({
    method: 'GET',
    headers: new Headers()
  });

  assert.equal(response.jsonBody?.status, 'ok');
  assert.equal(typeof response.headers?.['X-Grimoire-Duration-Ms'], 'string');
});
