import assert from 'node:assert/strict';
import test from 'node:test';

import { getCorsMode, validateConfiguration } from '../src/utils/config.js';

test('validateConfiguration throws when allowlist mode has no origins', () => {
  assert.throws(
    () => validateConfiguration({ allowedOrigins: [], allowPermissiveLocalCors: false }),
    /ALLOWED_ORIGINS/
  );
});

test('validateConfiguration allows empty origins in permissive local mode', () => {
  assert.doesNotThrow(() => validateConfiguration({ allowedOrigins: [], allowPermissiveLocalCors: true }));
});

test('getCorsMode reports the configured mode', () => {
  assert.equal(getCorsMode({ allowPermissiveLocalCors: false }), 'allowlist');
  assert.equal(getCorsMode({ allowPermissiveLocalCors: true }), 'permissive-local');
});
