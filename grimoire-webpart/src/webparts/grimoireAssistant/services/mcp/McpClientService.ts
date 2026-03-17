/**
 * McpClientService
 * Frontend service that communicates with grimoire-backend MCP endpoints.
 * All MCP traffic routes through the backend proxy — no direct browser-to-MCP-server calls.
 *
 * Endpoints:
 *   POST /api/mcp/connect     → connect to MCP server, get sessionId + tools
 *   POST /api/mcp/execute     → execute a tool on a connected session
 *   POST /api/mcp/disconnect  → close a session
 *   GET  /api/mcp/sessions    → list active sessions
 */

import type {
  IMcpTool,
  IMcpExecuteResult,
  IMcpContent
} from '../../models/IMcpTypes';
import { logService } from '../logging/LogService';
import { normalizeError } from '../utils/errorUtils';

const MCP_CONNECT_TIMEOUT_MS = 20_000;
const MCP_EXECUTE_TIMEOUT_MS = 60_000;
const MCP_DISCONNECT_TIMEOUT_MS = 10_000;

export interface IMcpConnectResult {
  sessionId: string;
  serverName: string;
  tools: IMcpTool[];
  connected: boolean;
}

export class McpClientService {
  private readonly proxyUrl: string;
  private readonly apiKey: string;

  constructor(proxyUrl: string, apiKey: string) {
    // Strip trailing /api if present to normalize the base URL
    this.proxyUrl = proxyUrl.replace(/\/api\/?$/, '');
    this.apiKey = apiKey;
  }

  /**
   * Connect to an MCP server via the backend proxy.
   * Returns sessionId and discovered tools.
   */
  public async connect(
    serverUrl: string,
    serverName: string,
    bearerToken?: string
  ): Promise<IMcpConnectResult> {
    const startTime = performance.now();
    logService.info('mcp', `Connecting to ${serverName} (${serverUrl})`);

    try {
      const response = await this.fetchWithTimeout(
        `${this.proxyUrl}/api/mcp/connect`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-functions-key': this.apiKey
          },
          body: JSON.stringify({ serverUrl, serverName, bearerToken })
        },
        MCP_CONNECT_TIMEOUT_MS,
        'connect'
      );

      const durationMs = Math.round(performance.now() - startTime);

      if (!response.ok) {
        const errorBody = await response.json().catch(() => ({ error: response.statusText }));
        const errorMsg = (errorBody as { error?: string }).error || `HTTP ${response.status}`;
        logService.error('mcp', `Connect failed: ${errorMsg}`, undefined, durationMs);
        throw new Error(errorMsg);
      }

      const result = await response.json() as IMcpConnectResult;
      logService.info(
        'mcp',
        `Connected to ${serverName}: ${result.tools.length} tools`,
        JSON.stringify(result.tools.map((t: IMcpTool) => t.name)),
        durationMs
      );

      return result;
    } catch (error) {
      const durationMs = Math.round(performance.now() - startTime);
      const normalizedError = normalizeError(error, 'MCP connect failed');
      if (normalizedError.message && !normalizedError.message.startsWith('HTTP')) {
        logService.error('mcp', `Connect error: ${normalizedError.message}`, undefined, durationMs);
      }
      throw error;
    }
  }

  /**
   * Execute a tool on an existing MCP session.
   */
  public async execute(
    sessionId: string,
    toolName: string,
    args?: Record<string, unknown>
  ): Promise<IMcpExecuteResult> {
    const startTime = performance.now();
    logService.info('mcp', `Execute: ${toolName} on session ${sessionId.substring(0, 8)}...`);

    try {
      const response = await this.fetchWithTimeout(
        `${this.proxyUrl}/api/mcp/execute`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-functions-key': this.apiKey
          },
          body: JSON.stringify({ sessionId, toolName, arguments: args || {} })
        },
        MCP_EXECUTE_TIMEOUT_MS,
        'execute'
      );

      const durationMs = Math.round(performance.now() - startTime);

      if (!response.ok) {
        const errorBody = await response.json().catch(() => ({ error: response.statusText }));
        const errorMsg = (errorBody as { error?: string }).error || `HTTP ${response.status}`;
        logService.error('mcp', `Execute failed: ${errorMsg}`, undefined, durationMs);
        return {
          success: false,
          content: [],
          error: errorMsg,
          durationMs
        };
      }

      const result = await response.json() as IMcpExecuteResult;
      logService.info(
        'mcp',
        `Tool ${toolName} completed`,
        this.summarizeContent(result.content),
        durationMs
      );

      return result;
    } catch (error) {
      const durationMs = Math.round(performance.now() - startTime);
      const normalizedError = normalizeError(error, 'MCP execute failed');
      logService.error('mcp', `Execute error: ${normalizedError.message}`, undefined, durationMs);
      return {
        success: false,
        content: [],
        error: normalizedError.message,
        durationMs
      };
    }
  }

  /**
   * Disconnect an MCP session.
   */
  public async disconnect(sessionId: string): Promise<void> {
    logService.info('mcp', `Disconnecting session: ${sessionId.substring(0, 8)}...`);

    try {
      const response = await this.fetchWithTimeout(
        `${this.proxyUrl}/api/mcp/disconnect`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-functions-key': this.apiKey
          },
          body: JSON.stringify({ sessionId })
        },
        MCP_DISCONNECT_TIMEOUT_MS,
        'disconnect'
      );

      if (!response.ok) {
        const errorBody = await response.json().catch(() => ({ error: response.statusText }));
        const errorMsg = (errorBody as { error?: string }).error || `HTTP ${response.status}`;
        throw new Error(errorMsg);
      }

      logService.info('mcp', `Session ${sessionId.substring(0, 8)}... disconnected`);
    } catch (error) {
      const normalizedError = normalizeError(error, 'MCP disconnect failed');
      logService.error('mcp', `Disconnect error: ${normalizedError.message}`);
      throw error;
    }
  }

  /**
   * Summarize MCP content array for logging.
   */
  private summarizeContent(content: IMcpContent[]): string {
    if (!content || content.length === 0) return '(empty)';
    return content.map((c) => {
      if (c.type === 'text' && c.text) {
        return c.text.length > 120 ? `${c.text.substring(0, 120)}...` : c.text;
      }
      return `[${c.type}]`;
    }).join('; ');
  }

  private async fetchWithTimeout(
    url: string,
    init: RequestInit,
    timeoutMs: number,
    operation: string
  ): Promise<Response> {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    try {
      return await fetch(url, { ...init, signal: controller.signal });
    } catch (error) {
      const normalizedError = normalizeError(error, `MCP ${operation} failed`);
      if (normalizedError.name === 'AbortError') {
        throw new Error(`MCP ${operation} timed out after ${Math.round(timeoutMs / 1000)}s`);
      }
      throw error;
    } finally {
      clearTimeout(timeout);
    }
  }
}
