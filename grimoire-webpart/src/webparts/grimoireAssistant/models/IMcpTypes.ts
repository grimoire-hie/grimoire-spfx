/**
 * IMcpTypes — MCP session and tool type definitions
 */

export interface IMcpConnection {
  /** Unique session ID from the backend */
  sessionId: string;
  /** MCP server URL */
  serverUrl: string;
  /** Display name for the server */
  serverName: string;
  /** Available tools from this server */
  tools: IMcpTool[];
  /** Connection state */
  state: 'connecting' | 'connected' | 'error' | 'disconnected';
  /** Error message if state is 'error' */
  error?: string;
  /** When the session was established */
  connectedAt: Date;
}

export interface IMcpTool {
  /** Tool name (e.g., "search_documents") */
  name: string;
  /** Human-readable description */
  description: string;
  /** JSON Schema for input parameters */
  inputSchema: Record<string, unknown>;
}

export interface IMcpExecuteRequest {
  sessionId: string;
  toolName: string;
  arguments: Record<string, unknown>;
}

export interface IMcpExecuteResult {
  /** Whether the tool call succeeded */
  success: boolean;
  /** Tool execution result (content array from MCP spec) */
  content: IMcpContent[];
  /** Error message if not success */
  error?: string;
  /** Execution time in ms */
  durationMs: number;
}

export interface IMcpContent {
  type: 'text' | 'image' | 'resource';
  text?: string;
  data?: string;
  mimeType?: string;
  uri?: string;
}
