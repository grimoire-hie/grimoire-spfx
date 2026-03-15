/**
 * Shared MCP utilities — text extraction and session retry logic.
 */

import type { IMcpContent } from '../../models/IMcpTypes';
import { logService } from '../logging/LogService';
import { McpClientService } from './McpClientService';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { connectToM365Server } from '../tools/ToolRuntimeSharedHelpers';

/**
 * Extract the server ID (last path segment) from an MCP server URL.
 * Handles trailing slashes. Returns '' if no segments found.
 */
export function extractServerIdFromUrl(serverUrl: string): string {
  const segments = serverUrl.split('/').filter(Boolean);
  return segments.length > 0 ? segments[segments.length - 1] : '';
}

/**
 * Extract text parts from MCP multi-part content arrays.
 * Returns an array of text strings (filters non-text parts).
 */
export function extractMcpTextParts(content: Array<{ type: string; text?: string }>): string[] {
  return content.filter((c) => c.type === 'text' && c.text).map((c) => c.text || '');
}

/**
 * Extract and join all text from MCP multi-part content.
 */
export function extractMcpText(content: Array<{ type: string; text?: string }>): string {
  return extractMcpTextParts(content).join('\n');
}

/**
 * Execute an MCP tool call with automatic session expiry retry.
 * If the first attempt fails with "not found" or "expired", reconnects and retries once.
 */
export async function withMcpRetry(
  mcpClient: McpClientService,
  sessionId: string,
  toolName: string,
  toolArgs: Record<string, unknown>,
  serverUrl: string,
  serverName: string
): Promise<{ success: boolean; content: IMcpContent[]; error?: string }> {
  let result = await mcpClient.execute(sessionId, toolName, toolArgs);

  if (!result.success && result.error && (result.error.includes('not found') || result.error.includes('expired'))) {
    logService.warning('mcp', `${serverName} session expired, reconnecting...`);
    useGrimoireStore.getState().removeMcpConnection(sessionId);
    const freshSession = await connectToM365Server(
      mcpClient, serverUrl, serverName, useGrimoireStore.getState().getToken
    );
    result = await mcpClient.execute(freshSession, toolName, toolArgs);
  }

  return result;
}
