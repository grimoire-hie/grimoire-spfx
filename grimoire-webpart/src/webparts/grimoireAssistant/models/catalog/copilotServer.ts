/**
 * Copilot Search MCP Server catalog entry.
 * Generated from discovery-output.json — 1 tools.
 * Source of truth: MCP protocol discovery, not guesses.
 */

import type { IMcpCatalogEntry } from '../McpServerCatalog';
import { GATEWAY_BASE } from './constants';

export const COPILOT_SERVER: IMcpCatalogEntry = {
  id: 'mcp_M365Copilot',
  name: 'Copilot Search',
  description: 'Universal search across all M365 content',
  scope: 'McpServers.CopilotMCP.All',
  urlTemplate: `${GATEWAY_BASE}/mcp_M365Copilot`,
  tools: [
    { name: 'copilot_chat', description: 'Use this tool to search internal Microsoft 365 content (documents, emails, chats, sites, files) when the specific workload is unclear or spans multiple areas, but always prefer workload-specific to...', inputSchema: {"type":"object","properties":{"message":{"type":"string","description":"User message text to send to Copilot."},"conversationId":{"type":"string","description":"Existing conversation id (GUID). Auto-created if missing."},"agentId":{"type":"string","description":"Optional agent id; only sent when provided."},"fileUris":{"type":"array","description":"Optional list of file URIs to ground the response (SharePoint/OneDrive/etc.).","items":{"type":"string"}},"enableWebSearch":{"type":"boolean","description":"Enable or disable web search grounding. Defaults to false (web search disabled). Set to true to allow Copilot to use web search for enhanced respon..."}},"required":["message"]}, blockHint: 'markdown' as const }
  ]
};
