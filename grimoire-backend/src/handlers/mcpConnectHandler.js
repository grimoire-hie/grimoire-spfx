/**
 * MCP Connect handler.
 * POST /api/mcp/connect — initialize a session-based MCP connection.
 * Returns sessionId + discovered tools.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { validateAuth } from "../middleware/auth.js";
import { enforceRateLimit } from "../middleware/rateLimit.js";
import { connectSession } from "../mcp/McpSessionManager.js";
import { validateMcpTargetUrl } from "../utils/mcpUrlPolicy.js";

export async function mcpConnectHandler(request, context) {
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

  const { serverUrl, serverName, bearerToken } = body;

  if (!serverUrl || typeof serverUrl !== "string") {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing required field: serverUrl" } };
  }

  const targetValidation = validateMcpTargetUrl(serverUrl, {
    requiresTokenForwarding: !!bearerToken,
  });
  if (!targetValidation.allowed) {
    return {
      status: 400,
      headers: corsHeaders,
      jsonBody: { error: targetValidation.error, serverUrl },
    };
  }

  context.log(`[mcp-connect] Connecting to: ${serverName || serverUrl}`);

  try {
    const result = await connectSession(serverUrl, serverName || serverUrl, bearerToken);

    context.log(`[mcp-connect] Connected. Session: ${result.sessionId}, Tools: ${result.tools.length}`);

    return {
      headers: corsHeaders,
      jsonBody: {
        ...result,
        connected: true,
      },
    };
  } catch (error) {
    context.error(`[mcp-connect] Error: ${error.message}`);
    return {
      status: 502,
      headers: corsHeaders,
      jsonBody: { error: error.message, serverUrl, connected: false },
    };
  }
}
