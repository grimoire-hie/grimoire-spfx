import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IMcpConnection } from '../../models/IMcpTypes';
import { MCP_AUDIENCE } from '../../models/McpServerCatalog';
import { logService } from '../logging/LogService';
import type { McpClientService } from '../mcp/McpClientService';

function extractMcpTextParts(content: Array<{ type: string; text?: string }>): string[] {
  return content.filter((c) => c.type === 'text' && c.text).map((c) => c.text || '');
}

export function extractMcpReply(content: Array<{ type: string; text?: string }>): { reply: string; raw: string } {
  const parts = extractMcpTextParts(content);
  for (let i = 0; i < parts.length; i++) {
    try {
      const parsed = JSON.parse(parts[i]);
      if (parsed && typeof parsed.reply === 'string' && parsed.reply) {
        return { reply: parsed.reply, raw: parts.join('\n') };
      }
    } catch {
      // no-op
    }
  }
  return { reply: '', raw: parts.join('\n') };
}

export function findExistingSession(connections: IMcpConnection[], serverUrl: string): string | undefined {
  let sessionId: string | undefined;
  connections.forEach((c: IMcpConnection) => {
    if (c.serverUrl === serverUrl && c.state === 'connected') {
      sessionId = c.sessionId;
    }
  });
  return sessionId;
}

export async function connectToM365Server(
  mcpClient: McpClientService,
  serverUrl: string,
  serverName: string,
  getToken: ((resource: string) => Promise<string>) | undefined
): Promise<string> {
  let bearerToken: string | undefined;
  if (getToken) {
    try {
      bearerToken = await getToken(MCP_AUDIENCE);
    } catch (err) {
      logService.warning('mcp', `Token acquisition for ${MCP_AUDIENCE} failed: ${(err as Error).message}. Connecting without token.`);
    }
  }

  const connectResult = await mcpClient.connect(serverUrl, serverName, bearerToken);
  const connection: IMcpConnection = {
    sessionId: connectResult.sessionId,
    serverUrl,
    serverName: connectResult.serverName || serverName,
    tools: connectResult.tools,
    state: 'connected',
    connectedAt: new Date()
  };
  useGrimoireStore.getState().addMcpConnection(connection);
  logService.info('mcp', `Auto-connected to ${serverName}: ${connectResult.tools.length} tools`);
  return connectResult.sessionId;
}
