/**
 * MCP Disconnect handler.
 * POST /api/mcp/disconnect — close an MCP session.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { resolveCallerId } from "../middleware/callerIdentity.js";
import { disconnectSession } from "../mcp/McpSessionManager.js";

export async function mcpDisconnectHandler(request, context) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  const callerId = resolveCallerId(request);

  let body;
  try {
    body = await request.json();
  } catch {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Invalid JSON body." } };
  }

  const { sessionId } = body;
  if (!sessionId) {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing required field: sessionId" } };
  }

  context.log(`[mcp-disconnect] Disconnecting session: ${sessionId}`);

  try {
    const result = await disconnectSession(sessionId, callerId);
    return { headers: corsHeaders, jsonBody: result };
  } catch (error) {
    context.error(`[mcp-disconnect] Error: ${error.message}`);
    return {
      status: 500,
      headers: corsHeaders,
      jsonBody: { error: "Failed to disconnect MCP session.", detail: error.message, sessionId },
    };
  }
}
