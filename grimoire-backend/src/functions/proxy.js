/**
 * Route registrations for grimoire-backend.
 * All business logic lives in handlers/ — this file only wires routes.
 *
 * Routes:
 *   GET  /api/health                         → healthHandler
 *   POST /api/realtime/token                 → realtimeTokenHandler
 *   POST /api/mcp/discover                   → mcpDiscoverHandler
 *   POST /api/mcp/discover-all               → mcpDiscoverAllHandler
 *   POST /api/mcp/connect                    → mcpConnectHandler
 *   POST /api/mcp/execute                    → mcpExecuteHandler
 *   POST /api/mcp/disconnect                 → mcpDisconnectHandler
 *   GET  /api/mcp/sessions                   → mcpSessionsHandler
 *   *    /api/{backend}/openai/{*path}        → llmProxyHandler
 */

import { app } from "@azure/functions";
import { validateConfiguration } from "../utils/config.js";

import { healthHandler } from "../handlers/healthHandler.js";
import { llmProxyHandler } from "../handlers/llmProxyHandler.js";
import { realtimeTokenHandler } from "../handlers/realtimeTokenHandler.js";
import { mcpDiscoverHandler } from "../handlers/mcpDiscoverHandler.js";
import { mcpDiscoverAllHandler } from "../handlers/mcpDiscoverAllHandler.js";
import { mcpConnectHandler } from "../handlers/mcpConnectHandler.js";
import { mcpExecuteHandler } from "../handlers/mcpExecuteHandler.js";
import { mcpDisconnectHandler } from "../handlers/mcpDisconnectHandler.js";
import { mcpSessionsHandler } from "../handlers/mcpSessionsHandler.js";
import { userNotesHandler, userPreferencesHandler } from "../handlers/userContextHandler.js";

validateConfiguration();

// ─── Health ───────────────────────────────────────────────────
app.http("health", {
  methods: ["GET", "OPTIONS"],
  authLevel: "function",
  route: "health",
  handler: healthHandler,
});

// ─── Realtime Token ───────────────────────────────────────────
app.http("realtimeToken", {
  methods: ["POST", "OPTIONS"],
  authLevel: "function",
  route: "realtime/token",
  handler: realtimeTokenHandler,
});

// ─── MCP Discovery (legacy stateless) ─────────────────────────
app.http("mcpDiscover", {
  methods: ["POST", "OPTIONS"],
  authLevel: "function",
  route: "mcp/discover",
  handler: mcpDiscoverHandler,
});

// ─── MCP Discover All ────────────────────────────────────────
app.http("mcpDiscoverAll", {
  methods: ["POST", "OPTIONS"],
  authLevel: "function",
  route: "mcp/discover-all",
  handler: mcpDiscoverAllHandler,
});

// ─── MCP Session-Based ────────────────────────────────────────
app.http("mcpConnect", {
  methods: ["POST", "OPTIONS"],
  authLevel: "function",
  route: "mcp/connect",
  handler: mcpConnectHandler,
});

app.http("mcpExecute", {
  methods: ["POST", "OPTIONS"],
  authLevel: "function",
  route: "mcp/execute",
  handler: mcpExecuteHandler,
});

app.http("mcpDisconnect", {
  methods: ["POST", "OPTIONS"],
  authLevel: "function",
  route: "mcp/disconnect",
  handler: mcpDisconnectHandler,
});

app.http("mcpSessions", {
  methods: ["GET", "OPTIONS"],
  authLevel: "function",
  route: "mcp/sessions",
  handler: mcpSessionsHandler,
});

// ─── User Context (notes & preferences) ──────────────────────
app.http("userNotes", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  route: "user/notes",
  handler: userNotesHandler,
});

app.http("userPreferences", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  route: "user/preferences",
  handler: userPreferencesHandler,
});

// ─── LLM Proxy (catch-all for OpenAI routes) ─────────────────
app.http("proxy", {
  methods: ["GET", "POST", "PUT", "DELETE", "PATCH", "OPTIONS"],
  authLevel: "function",
  route: "{backend}/openai/{*path}",
  handler: llmProxyHandler,
});
