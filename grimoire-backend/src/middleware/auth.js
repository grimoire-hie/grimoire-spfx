/**
 * Authentication middleware.
 * Validates proxy API key from request headers.
 */

import { ALLOWED_KEYS } from "../utils/config.js";

/**
 * Extract API key from request headers.
 */
export function extractApiKey(request) {
  return (
    request.headers.get("api-key") ||
    request.headers.get("authorization")?.replace(/^Bearer\s+/i, "")
  );
}

/**
 * Validate API key. Returns { authenticated, apiKey, errorResponse? }.
 */
export function validateAuth(request, corsHeaders) {
  const apiKey = extractApiKey(request);

  if (!apiKey) {
    return {
      authenticated: false,
      apiKey: null,
      errorResponse: {
        status: 401,
        headers: corsHeaders,
        jsonBody: { error: "Missing API key. Use 'api-key' header." },
      },
    };
  }

  if (!ALLOWED_KEYS.includes(apiKey)) {
    return {
      authenticated: false,
      apiKey,
      errorResponse: {
        status: 403,
        headers: corsHeaders,
        jsonBody: { error: "Invalid API key." },
      },
    };
  }

  return { authenticated: true, apiKey, errorResponse: null };
}
