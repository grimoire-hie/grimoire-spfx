import assert from 'node:assert/strict';
import test from 'node:test';

import { buildCorsHeaders } from '../src/middleware/cors.js';

const ALLOWLIST_CONFIG = {
  allowedOrigins: ['https://contoso.sharepoint.com'],
  allowPermissiveLocalCors: false,
};

test('buildCorsHeaders returns the expected CORS headers for allowed origins', () => {
  const headers = buildCorsHeaders('https://contoso.sharepoint.com', ALLOWLIST_CONFIG);

  assert.equal(headers['Access-Control-Allow-Origin'], 'https://contoso.sharepoint.com');
  assert.equal(headers['Access-Control-Allow-Methods'], 'GET, POST, PUT, DELETE, PATCH, OPTIONS');
  assert.equal(headers['Access-Control-Allow-Headers'], 'Content-Type, api-key, Authorization');
  assert.equal(headers['Access-Control-Max-Age'], '3600');
});

test('buildCorsHeaders omits allow-origin for disallowed origins', () => {
  const headers = buildCorsHeaders('https://evil.example', ALLOWLIST_CONFIG);

  assert.deepEqual(headers, {});
});

test('buildCorsHeaders reflects the origin in permissive local mode', () => {
  const headers = buildCorsHeaders('https://localhost:4321', {
    allowedOrigins: [],
    allowPermissiveLocalCors: true,
  });

  assert.equal(headers['Access-Control-Allow-Origin'], 'https://localhost:4321');
});
