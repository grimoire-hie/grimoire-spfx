/**
 * MCP Discover All handler.
 * POST /api/mcp/discover-all — connect to all 10 Agent 365 servers sequentially,
 * collect real tool schemas, and return them in a single response.
 *
 * Body: { bearerToken: string, envId: string }
 * Response: { servers: [{ id, name, tools: [{ name, description, inputSchema }] }] }
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { resolveCallerId } from "../middleware/callerIdentity.js";
import { enforceRateLimit } from "../middleware/rateLimit.js";
import { connectSession, disconnectSession } from "../mcp/McpSessionManager.js";

const GATEWAY_BASE = "https://agent365.svc.cloud.microsoft/mcp/environments";

const SERVERS = [
  { id: "mcp_ODSPRemoteServer", name: "SharePoint & OneDrive" },
  { id: "mcp_SharePointListsTools", name: "SharePoint Lists" },
  { id: "mcp_MailTools", name: "Outlook Mail" },
  { id: "mcp_CalendarTools", name: "Outlook Calendar" },
  { id: "mcp_TeamsServer", name: "Microsoft Teams" },
  { id: "mcp_MeServer", name: "User Profile" },
  { id: "mcp_M365Copilot", name: "Copilot Search" },
  { id: "mcp_DataverseServer", name: "Dataverse" },
  { id: "mcp_WordServer", name: "Word Documents" },
  { id: "mcp_McpManagement", name: "MCP Management" },
];

const SERVER_TIMEOUT_MS = 30000;

export async function mcpDiscoverAllHandler(request, context) {
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  const callerId = resolveCallerId(request);
  const rateLimitError = enforceRateLimit(callerId, corsHeaders);
  if (rateLimitError) return rateLimitError;

  let body;
  try {
    body = await request.json();
  } catch {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Invalid JSON body." } };
  }

  const { bearerToken, envId } = body;
  if (!envId || typeof envId !== "string") {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing required field: envId" } };
  }
  if (!bearerToken || typeof bearerToken !== "string") {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Missing required field: bearerToken" } };
  }

  context.log(`[mcp-discover-all] Starting discovery for ${SERVERS.length} servers (envId: ${envId})`);

  const results = [];

  // Sequential to avoid token exhaustion
  for (const server of SERVERS) {
    const serverUrl = `${GATEWAY_BASE}/${envId}/servers/${server.id}`;
    context.log(`[mcp-discover-all] Connecting to ${server.name} (${server.id})...`);

    try {
      const result = await withTimeoutAndLateCleanup(
        connectSession(serverUrl, server.name, bearerToken, callerId),
        SERVER_TIMEOUT_MS,
        async (lateResult) => {
          if (!lateResult?.sessionId) return;
          context.warn(`[mcp-discover-all] Late session detected after timeout for ${server.name}. Cleaning up: ${lateResult.sessionId}`);
          try {
            await disconnectSession(lateResult.sessionId);
          } catch {
            // Best-effort cleanup
          }
        }
      );

      results.push({
        id: server.id,
        name: server.name,
        tools: result.tools,
        toolCount: result.tools.length,
        error: null,
      });

      context.log(`[mcp-discover-all] ${server.name}: ${result.tools.length} tools`);

      // Disconnect immediately — we only needed the tool list
      try {
        await disconnectSession(result.sessionId);
      } catch {
        // Best-effort disconnect
      }
    } catch (error) {
      context.warn(`[mcp-discover-all] ${server.name} failed: ${error.message}`);
      results.push({
        id: server.id,
        name: server.name,
        tools: [],
        toolCount: 0,
        error: error.message,
      });
    }
  }

  const totalTools = results.reduce((sum, r) => sum + r.toolCount, 0);
  const successCount = results.filter((r) => !r.error).length;
  context.log(`[mcp-discover-all] Done: ${successCount}/${SERVERS.length} servers, ${totalTools} total tools`);

  return {
    headers: corsHeaders,
    jsonBody: {
      servers: results,
      totalTools,
      successCount,
      serverCount: SERVERS.length,
    },
  };
}

function withTimeoutAndLateCleanup(promise, ms, onLateResolve) {
  let timer;
  let timedOut = false;

  const wrapped = Promise.resolve(promise)
    .then(async (value) => {
      if (timedOut && onLateResolve) {
        await onLateResolve(value);
      }
      return value;
    })
    .finally(() => {
      clearTimeout(timer);
    });

  return new Promise((resolve, reject) => {
    timer = setTimeout(() => {
      timedOut = true;
      reject(new Error(`Timed out after ${ms}ms`));
    }, ms);

    wrapped
      .then((value) => {
        if (!timedOut) resolve(value);
      })
      .catch((error) => {
        if (!timedOut) reject(error);
      });
  });
}
