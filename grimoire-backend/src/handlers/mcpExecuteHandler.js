/**
 * MCP Execute handler.
 * POST /api/mcp/execute — execute a tool on a connected MCP session.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { validateAuth } from "../middleware/auth.js";
import { enforceRateLimit } from "../middleware/rateLimit.js";
import { executeToolOnSession } from "../mcp/McpSessionManager.js";

export async function mcpExecuteHandler(request, context) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  const auth = validateAuth(request, corsHeaders);
  if (!auth.authenticated) return auth.errorResponse;
  const rateLimitError = enforceRateLimit(auth.apiKey, corsHeaders);
  if (rateLimitError) return rateLimitError;

  let body;
  try {
    body = await request.json();
  } catch {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Invalid JSON body." } };
  }

  const { sessionId, toolName, arguments: args } = body;

  if (!sessionId) {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing required field: sessionId" } };
  }
  if (!toolName) {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing required field: toolName" } };
  }

  context.log(`[mcp-execute] Session: ${sessionId}, Tool: ${toolName}`);

  try {
    const result = await executeToolOnSession(sessionId, toolName, args);

    context.log(`[mcp-execute] Tool ${toolName} completed in ${result.durationMs}ms`);

    return {
      headers: corsHeaders,
      jsonBody: result,
    };
  } catch (error) {
    context.error(`[mcp-execute] Error: ${error.message}`);
    return {
      status: error.message.includes("not found") ? 404 : 500,
      headers: corsHeaders,
      jsonBody: { error: error.message, sessionId, toolName },
    };
  }
}
