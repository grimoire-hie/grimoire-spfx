import type { IMcpConnection } from '../../models/IMcpTypes';
import { MCP_AUDIENCE } from '../../models/McpServerCatalog';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { connectToM365Server, extractMcpReply, findExistingSession } from './ToolRuntimeSharedHelpers';

describe('ToolRuntimeSharedHelpers', () => {
  beforeEach(() => {
    useGrimoireStore.setState({ mcpConnections: [] });
  });

  it('extracts reply from MCP JSON text payload', () => {
    const result = extractMcpReply([
      { type: 'text', text: '{"reply":"ok"}' },
      { type: 'text', text: 'ignored' }
    ]);
    expect(result.reply).toBe('ok');
    expect(result.raw).toContain('{"reply":"ok"}');
  });

  it('falls back to joined raw text when reply is missing', () => {
    const result = extractMcpReply([
      { type: 'text', text: 'first' },
      { type: 'text', text: 'second' }
    ]);
    expect(result.reply).toBe('');
    expect(result.raw).toBe('first\nsecond');
  });

  it('finds existing connected session for a server URL', () => {
    const connections: IMcpConnection[] = [
      {
        sessionId: 's1',
        serverUrl: 'https://mcp.contoso.com/a',
        serverName: 'A',
        tools: [],
        state: 'connected',
        connectedAt: new Date()
      },
      {
        sessionId: 's2',
        serverUrl: 'https://mcp.contoso.com/b',
        serverName: 'B',
        tools: [],
        state: 'disconnected',
        connectedAt: new Date()
      }
    ];

    expect(findExistingSession(connections, 'https://mcp.contoso.com/a')).toBe('s1');
    expect(findExistingSession(connections, 'https://mcp.contoso.com/b')).toBeUndefined();
  });

  it('connects to M365 server and registers a connection in store', async () => {
    const getToken = jest.fn(async () => 'bearer-token');
    const connect = jest.fn(async () => ({
      sessionId: 'session-123',
      serverName: 'M365 Test',
      tools: [{ name: 'toolA', description: 'A', inputSchema: {} }]
    }));

    const mcpClient = { connect } as unknown as {
      connect: (serverUrl: string, serverName: string, token?: string) => Promise<{
        sessionId: string;
        serverName: string;
        tools: Array<{ name: string; description: string; inputSchema: unknown }>;
      }>;
    };

    const sessionId = await connectToM365Server(
      mcpClient as never,
      'https://mcp.contoso.com/server',
      'M365 Test',
      getToken
    );

    expect(sessionId).toBe('session-123');
    expect(getToken).toHaveBeenCalledWith(MCP_AUDIENCE);
    expect(connect).toHaveBeenCalledWith('https://mcp.contoso.com/server', 'M365 Test', 'bearer-token');
    const conn = useGrimoireStore.getState().mcpConnections.find((c) => c.sessionId === 'session-123');
    expect(conn).toBeDefined();
    expect(conn?.serverName).toBe('M365 Test');
  });
});
