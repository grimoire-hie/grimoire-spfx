/**
 * Caller identity resolution.
 * Resolves a stable caller ID from Easy Auth headers (preferred)
 * or falls back to "anonymous" for function-key-authenticated callers.
 *
 * Used for MCP session ownership and per-user rate limiting.
 */

import { extractAuthenticatedUser } from "./userIdentity.js";

/**
 * Resolve a stable caller identifier from the request.
 * @returns {string} Caller ID: "easyauth:<objectId>" or "anonymous".
 */
export function resolveCallerId(request) {
  const user = extractAuthenticatedUser(request);
  if (user?.objectId) {
    return `easyauth:${user.objectId}`;
  }
  return "anonymous";
}
