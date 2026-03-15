/**
 * useMcpServers
 * Hook for managing MCP server connections from the frontend.
 * Handles connect/disconnect/execute lifecycle using McpClientService.
 */

import * as React from 'react';
import { useGrimoireStore } from '../store/useGrimoireStore';
import { McpClientService } from '../services/mcp/McpClientService';
import type { IMcpConnection, IMcpTool } from '../models/IMcpTypes';
import { logService } from '../services/logging/LogService';

export interface IUseMcpServers {
  /** Connect to an MCP server. Returns sessionId if successful. */
  connectServer: (serverUrl: string, serverName: string, bearerToken?: string) => Promise<string | undefined>;
  /** Disconnect an MCP session by sessionId. */
  disconnectServer: (sessionId: string) => Promise<void>;
  /** Execute a tool on a connected session. Returns result as string. */
  executeTool: (sessionId: string, toolName: string, args?: Record<string, unknown>) => Promise<string>;
  /** Get all tools across all connected servers. */
  getAllTools: () => IMcpTool[];
  /** Find which session has a given tool name. */
  findSessionForTool: (toolName: string) => string | undefined;
  /** The McpClientService instance (undefined if no proxy config). */
  client: McpClientService | undefined;
}

export function useMcpServers(): IUseMcpServers {
  const proxyConfig = useGrimoireStore((s) => s.proxyConfig);
  const mcpConnections = useGrimoireStore((s) => s.mcpConnections);
  const addMcpConnection = useGrimoireStore((s) => s.addMcpConnection);
  const updateMcpConnection = useGrimoireStore((s) => s.updateMcpConnection);
  const removeMcpConnection = useGrimoireStore((s) => s.removeMcpConnection);

  const clientRef = React.useRef<McpClientService | undefined>(undefined);

  // Create/update client when proxy config changes
  React.useEffect(() => {
    if (proxyConfig) {
      clientRef.current = new McpClientService(proxyConfig.proxyUrl, proxyConfig.proxyApiKey);
    } else {
      clientRef.current = undefined;
    }
  }, [proxyConfig]);

  const connectServer = React.useCallback(async (
    serverUrl: string,
    serverName: string,
    bearerToken?: string
  ): Promise<string | undefined> => {
    const client = clientRef.current;
    if (!client) {
      logService.error('mcp', 'Cannot connect — no proxy config');
      return undefined;
    }

    // Check if already connected to this server
    const existing = mcpConnections.find(
      (c) => c.serverUrl === serverUrl && c.state === 'connected'
    );
    if (existing) {
      logService.info('mcp', `Already connected to ${serverName} (session: ${existing.sessionId.substring(0, 8)}...)`);
      return existing.sessionId;
    }

    // Add connecting state to store
    const tempConnection: IMcpConnection = {
      sessionId: '', // Will be updated on success
      serverUrl,
      serverName,
      tools: [],
      state: 'connecting',
      connectedAt: new Date()
    };

    // Use a temporary ID for tracking
    const tempId = `temp-${Date.now()}`;
    const connectingEntry: IMcpConnection = { ...tempConnection, sessionId: tempId };
    addMcpConnection(connectingEntry);

    try {
      const result = await client.connect(serverUrl, serverName, bearerToken);

      // Remove temp entry and add real one
      removeMcpConnection(tempId);
      const connection: IMcpConnection = {
        sessionId: result.sessionId,
        serverUrl,
        serverName: result.serverName || serverName,
        tools: result.tools,
        state: 'connected',
        connectedAt: new Date()
      };
      addMcpConnection(connection);

      return result.sessionId;
    } catch (error) {
      // Update temp entry to error state
      updateMcpConnection(tempId, {
        state: 'error',
        error: (error as Error).message
      });
      return undefined;
    }
  }, [mcpConnections, addMcpConnection, updateMcpConnection, removeMcpConnection]);

  const disconnectServer = React.useCallback(async (sessionId: string): Promise<void> => {
    const client = clientRef.current;
    if (!client) return;

    try {
      await client.disconnect(sessionId);
      updateMcpConnection(sessionId, { state: 'disconnected', error: undefined });
      removeMcpConnection(sessionId);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Disconnect failed';
      updateMcpConnection(sessionId, {
        state: 'error',
        error: `Disconnect failed: ${message}`
      });
    }
  }, [updateMcpConnection, removeMcpConnection]);

  const executeTool = React.useCallback(async (
    sessionId: string,
    toolName: string,
    args?: Record<string, unknown>
  ): Promise<string> => {
    const client = clientRef.current;
    if (!client) {
      return JSON.stringify({ success: false, error: 'No proxy config' });
    }

    const result = await client.execute(sessionId, toolName, args);

    if (!result.success) {
      return JSON.stringify({ success: false, error: result.error });
    }

    // Extract text content from MCP response
    const textContent = result.content
      .filter((c) => c.type === 'text' && c.text)
      .map((c) => c.text)
      .join('\n');

    return JSON.stringify({
      success: true,
      result: textContent || '(no text content)',
      durationMs: result.durationMs
    });
  }, []);

  const getAllTools = React.useCallback((): IMcpTool[] => {
    const tools: IMcpTool[] = [];
    mcpConnections.forEach((c) => {
      if (c.state === 'connected') {
        c.tools.forEach((t) => tools.push(t));
      }
    });
    return tools;
  }, [mcpConnections]);

  const findSessionForTool = React.useCallback((toolName: string): string | undefined => {
    let foundSessionId: string | undefined;
    mcpConnections.forEach((c) => {
      if (c.state === 'connected' && !foundSessionId) {
        const hasTool = c.tools.some((t) => t.name === toolName);
        if (hasTool) {
          foundSessionId = c.sessionId;
        }
      }
    });
    return foundSessionId;
  }, [mcpConnections]);

  return {
    connectServer,
    disconnectServer,
    executeTool,
    getAllTools,
    findSessionForTool,
    client: clientRef.current
  };
}
