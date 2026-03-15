/**
 * MCP Discovery handler (legacy stateless).
 * Calls an MCP server's tools/list endpoint via JSON-RPC 2.0.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { validateAuth } from "../middleware/auth.js";
import { enforceRateLimit } from "../middleware/rateLimit.js";
import { validateMcpTargetUrl } from "../utils/mcpUrlPolicy.js";

export async function mcpDiscoverHandler(request, context) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  // Auth + rate limit
  const auth = validateAuth(request, corsHeaders);
  if (!auth.authenticated) return auth.errorResponse;
  const rateLimitError = enforceRateLimit(auth.apiKey, corsHeaders);
  if (rateLimitError) return rateLimitError;

  // Parse body
  let body;
  try {
    body = await request.json();
  } catch {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Invalid JSON body." } };
  }

  const { serverUrl } = body;
  if (!serverUrl || typeof serverUrl !== "string") {
    return {
      status: 400,
      headers: corsHeaders,
      jsonBody: { error: "Missing required field: serverUrl" },
    };
  }

  const targetValidation = validateMcpTargetUrl(serverUrl);
  if (!targetValidation.allowed) {
    return { status: 400, headers: corsHeaders, jsonBody: { error: targetValidation.error, serverUrl } };
  }

  context.log(`[mcp-discover] Discovering tools from: ${serverUrl}`);

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000);

    const response = await fetch(serverUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "tools/list",
        params: {},
      }),
      signal: controller.signal,
    });

    clearTimeout(timeout);

    if (!response.ok) {
      return {
        status: 502,
        headers: corsHeaders,
        jsonBody: { error: `MCP server returned HTTP ${response.status}`, serverUrl, discovered: false },
      };
    }

    const data = await response.json();
    const tools = data?.result?.tools || data?.result || [];

    context.log(`[mcp-discover] Discovered ${tools.length} tool(s)`);

    return {
      headers: corsHeaders,
      jsonBody: { tools, serverUrl, discovered: true },
    };
  } catch (error) {
    const message = error.name === "AbortError"
      ? "MCP server did not respond within 10 seconds"
      : error.message;

    context.error(`[mcp-discover] Error: ${message}`);
    return {
      status: 502,
      headers: corsHeaders,
      jsonBody: { error: message, serverUrl, discovered: false },
    };
  }
}
