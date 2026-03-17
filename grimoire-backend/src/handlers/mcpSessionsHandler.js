/**
 * MCP Sessions handler.
 * GET /api/mcp/sessions — list active MCP sessions for the caller.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { resolveCallerId } from "../middleware/callerIdentity.js";
import { listSessions } from "../mcp/McpSessionManager.js";

export async function mcpSessionsHandler(request) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  const callerId = resolveCallerId(request);

  return {
    headers: corsHeaders,
    jsonBody: { sessions: listSessions(callerId) },
  };
}
