/**
 * MCP Sessions handler.
 * GET /api/mcp/sessions — list active MCP sessions.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { validateAuth } from "../middleware/auth.js";
import { listSessions } from "../mcp/McpSessionManager.js";

export async function mcpSessionsHandler(request) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  const auth = validateAuth(request, corsHeaders);
  if (!auth.authenticated) return auth.errorResponse;

  return {
    headers: corsHeaders,
    jsonBody: { sessions: listSessions() },
  };
}
