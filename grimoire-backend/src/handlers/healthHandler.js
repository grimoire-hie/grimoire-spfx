/**
 * Health check handler.
 * Protected by Azure function key auth — returns full backend details.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import {
  ALLOWED_ORIGINS,
  ALLOW_PERMISSIVE_LOCAL_CORS,
  BACKENDS,
  REQUESTS_PER_MINUTE,
  REQUESTS_PER_DAY,
  getCorsMode,
} from "../utils/config.js";
import { createTimingHeaders } from "../utils/diagnostics.js";

export function buildAuthenticatedHealthPayload(timestamp = new Date().toISOString()) {
  const backendStatus = {};
  for (const [name, config] of Object.entries(BACKENDS)) {
    backendStatus[name] = {
      endpoint: config.endpoint,
      auth: config.key ? "api-key" : "managed-identity",
    };
  }

  return {
    status: "ok",
    service: "grimoire-backend",
    timestamp,
    backends: backendStatus,
    limits: {
      perMinute: REQUESTS_PER_MINUTE,
      perDay: REQUESTS_PER_DAY,
    },
    cors: {
      mode: getCorsMode({ allowPermissiveLocalCors: ALLOW_PERMISSIVE_LOCAL_CORS }),
      allowedOrigins: ALLOWED_ORIGINS.length,
    },
    rateLimit: {
      mode: "memory",
      perMinute: REQUESTS_PER_MINUTE,
      perDay: REQUESTS_PER_DAY,
    },
  };
}

export async function healthHandler(request) {
  const requestStartedAt = Date.now();
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  return {
    headers: createTimingHeaders(corsHeaders, {
      totalDurationMs: Date.now() - requestStartedAt,
    }),
    jsonBody: buildAuthenticatedHealthPayload(),
  };
}
