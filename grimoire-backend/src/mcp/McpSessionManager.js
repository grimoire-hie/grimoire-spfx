/**
 * McpSessionManager
 * In-memory MCP session lifecycle: connect, execute, disconnect.
 * Sessions auto-expire after configurable timeout (default 30 min).
 * Uses @modelcontextprotocol/sdk for full protocol compliance.
 */

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { MCP_SESSION_TIMEOUT_MINUTES, MAX_MCP_SESSIONS } from "../utils/config.js";

const sessions = new Map();
let sessionCounter = 0;

/**
 * Connect to an MCP server and return session info + available tools.
 */
export async function connectSession(serverUrl, serverName, bearerToken) {
  // Enforce max sessions
  if (sessions.size >= MAX_MCP_SESSIONS) {
    // Clean up expired sessions first
    cleanupExpired();
    if (sessions.size >= MAX_MCP_SESSIONS) {
      throw new Error(`Maximum sessions (${MAX_MCP_SESSIONS}) reached. Disconnect a session first.`);
    }
  }

  const transport = new StreamableHTTPClientTransport(
    new URL(serverUrl),
    bearerToken ? { requestInit: { headers: { Authorization: `Bearer ${bearerToken}` } } } : undefined
  );

  const client = new Client({ name: "grimoire-backend", version: "1.0.0" });
  let connected = false;

  try {
    await client.connect(transport);
    connected = true;

    // List tools
    const toolsResult = await client.listTools();
    const tools = toolsResult.tools || [];

    sessionCounter++;
    const sessionId = `mcp-${sessionCounter}-${Date.now()}`;

    sessions.set(sessionId, {
      client,
      transport,
      serverUrl,
      serverName,
      tools,
      connectedAt: Date.now(),
      lastUsed: Date.now(),
    });

    return {
      sessionId,
      serverUrl,
      serverName,
      tools: tools.map((t) => ({
        name: t.name,
        description: t.description,
        inputSchema: t.inputSchema,
      })),
    };
  } catch (error) {
    if (connected) {
      try {
        await client.close();
      } catch {
        // Best-effort close
      }
    }
    const statusCode = error.code ? ` (HTTP ${error.code})` : "";
    throw new Error(`MCP session setup failed: ${error.message}${statusCode}`);
  }
}

/**
 * Execute a tool on a connected MCP session.
 */
export async function executeToolOnSession(sessionId, toolName, args) {
  const session = sessions.get(sessionId);
  if (!session) {
    throw new Error(`Session '${sessionId}' not found or expired.`);
  }

  session.lastUsed = Date.now();

  const startTime = performance.now();
  const result = await session.client.callTool({ name: toolName, arguments: args || {} });
  const durationMs = Math.round(performance.now() - startTime);

  return {
    success: !result.isError,
    content: result.content || [],
    error: result.isError ? (result.content?.[0]?.text || 'Tool execution failed') : undefined,
    durationMs,
  };
}

/**
 * Disconnect and clean up an MCP session.
 */
export async function disconnectSession(sessionId) {
  const session = sessions.get(sessionId);
  if (!session) {
    return { sessionId, disconnected: false, reason: "Session not found" };
  }

  try {
    await session.client.close();
  } catch {
    // Best-effort close
  }

  sessions.delete(sessionId);
  return { sessionId, disconnected: true };
}

/**
 * List all active sessions.
 */
export function listSessions() {
  const result = [];
  sessions.forEach((session, sessionId) => {
    result.push({
      sessionId,
      serverUrl: session.serverUrl,
      serverName: session.serverName,
      toolCount: session.tools.length,
      connectedAt: new Date(session.connectedAt).toISOString(),
      lastUsed: new Date(session.lastUsed).toISOString(),
    });
  });
  return result;
}

/**
 * Clean up expired sessions.
 */
function cleanupExpired() {
  const now = Date.now();
  const timeoutMs = MCP_SESSION_TIMEOUT_MINUTES * 60 * 1000;

  sessions.forEach((session, sessionId) => {
    if (now - session.lastUsed > timeoutMs) {
      try { session.client.close(); } catch { /* ignore */ }
      sessions.delete(sessionId);
    }
  });
}

// Periodic cleanup every 5 minutes
setInterval(cleanupExpired, 5 * 60 * 1000);
