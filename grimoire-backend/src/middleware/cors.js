/**
 * CORS middleware helpers.
 */

import { ALLOWED_ORIGINS, ALLOW_PERMISSIVE_LOCAL_CORS } from "../utils/config.js";

const ALLOWED_METHODS = "GET, POST, PUT, DELETE, PATCH, OPTIONS";
const ALLOWED_HEADERS = "Content-Type, api-key, Authorization";
const MAX_AGE_SECONDS = "3600";

export function isOriginAllowed(origin, config = {
  allowedOrigins: ALLOWED_ORIGINS,
  allowPermissiveLocalCors: ALLOW_PERMISSIVE_LOCAL_CORS,
}) {
  if (!origin) return false;
  return config.allowPermissiveLocalCors || config.allowedOrigins.includes(origin);
}

export function buildCorsHeaders(origin, config = {
  allowedOrigins: ALLOWED_ORIGINS,
  allowPermissiveLocalCors: ALLOW_PERMISSIVE_LOCAL_CORS,
}) {
  const headers = {};

  if (isOriginAllowed(origin, config)) {
    headers["Access-Control-Allow-Origin"] = origin;
    headers["Access-Control-Allow-Methods"] = ALLOWED_METHODS;
    headers["Access-Control-Allow-Headers"] = ALLOWED_HEADERS;
    headers["Access-Control-Max-Age"] = MAX_AGE_SECONDS;
  }

  return headers;
}

export function getCorsHeaders(request) {
  return buildCorsHeaders(request.headers.get("origin"));
}

/**
 * Handle CORS preflight. Returns a response if OPTIONS, or null if not.
 */
export function handlePreflight(request) {
  if (request.method === "OPTIONS") {
    return { status: 200, body: "", headers: getCorsHeaders(request) };
  }
  return null;
}
