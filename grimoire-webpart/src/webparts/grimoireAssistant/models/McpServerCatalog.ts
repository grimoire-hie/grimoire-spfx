/**
 * M365 MCP Server Catalog
 * Barrel file assembling per-server catalog entries from ./catalog/.
 * 8 Agent 365 servers, 100 tools — generated from MCP protocol discovery.
 */

// ─── Catalog Types ──────────────────────────────────────────────

export interface IMcpCatalogTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
  /** Optional hint to override schema inference for block type mapping */
  blockHint?: import('./IBlock').BlockType;
}

export interface IMcpCatalogEntry {
  id: string;
  name: string;
  description: string;
  scope: string;
  urlTemplate: string;
  tools: IMcpCatalogTool[];
}

export type McpCapabilityFocus =
  | 'sharepoint'
  | 'onedrive'
  | 'outlook'
  | 'mail'
  | 'calendar'
  | 'teams'
  | 'word'
  | 'profile'
  | 'copilot';

// ─── Constants (re-exported from shared file to avoid circular imports) ──

export { GATEWAY_BASE, MCP_AUDIENCE } from './catalog/constants';
import { GATEWAY_BASE } from './catalog/constants';

// ─── Per-Server Imports ─────────────────────────────────────────

import { ODSP_SERVER } from './catalog/odspServer';
import { LISTS_SERVER } from './catalog/listsServer';
import { MAIL_SERVER } from './catalog/mailServer';
import { CALENDAR_SERVER } from './catalog/calendarServer';
import { TEAMS_SERVER } from './catalog/teamsServer';
import { ME_SERVER } from './catalog/meServer';
import { COPILOT_SERVER } from './catalog/copilotServer';
import { WORD_SERVER } from './catalog/wordServer';
import { GRIMOIRE_SERVER } from './catalog/grimoireServer';

// ─── M365 MCP Server Catalog ───────────────────────────────────
// Excludes Dataverse (mcp_DataverseServer) and McpManagement (mcp_McpManagement)
// — both return 403, confirmed not needed.

export const M365_MCP_CATALOG: IMcpCatalogEntry[] = [
  ODSP_SERVER,
  LISTS_SERVER,
  MAIL_SERVER,
  CALENDAR_SERVER,
  TEAMS_SERVER,
  ME_SERVER,
  COPILOT_SERVER,
  WORD_SERVER
];

const CAPABILITY_FOCUS_SERVER_IDS: Readonly<Record<McpCapabilityFocus, readonly string[]>> = {
  sharepoint: ['mcp_ODSPRemoteServer', 'mcp_SharePointListsTools'],
  onedrive: ['mcp_ODSPRemoteServer'],
  outlook: ['mcp_MailTools', 'mcp_CalendarTools'],
  mail: ['mcp_MailTools'],
  calendar: ['mcp_CalendarTools'],
  teams: ['mcp_TeamsServer'],
  word: ['mcp_WordServer'],
  profile: ['mcp_MeServer'],
  copilot: ['mcp_M365Copilot']
};

const CAPABILITY_FOCUS_LABELS: Readonly<Record<McpCapabilityFocus, string>> = {
  sharepoint: 'SharePoint',
  onedrive: 'OneDrive',
  outlook: 'Outlook',
  mail: 'Outlook Mail',
  calendar: 'Outlook Calendar',
  teams: 'Teams',
  word: 'Word',
  profile: 'Profile',
  copilot: 'Copilot'
};

const CAPABILITY_FOCUS_NORMALIZATION: Readonly<Record<string, McpCapabilityFocus>> = {
  sharepoint: 'sharepoint',
  'share point': 'sharepoint',
  onedrive: 'onedrive',
  'one drive': 'onedrive',
  outlook: 'outlook',
  mail: 'mail',
  email: 'mail',
  emails: 'mail',
  inbox: 'mail',
  calendar: 'calendar',
  calendars: 'calendar',
  meeting: 'calendar',
  meetings: 'calendar',
  teams: 'teams',
  team: 'teams',
  chat: 'teams',
  chats: 'teams',
  channel: 'teams',
  channels: 'teams',
  word: 'word',
  profile: 'profile',
  me: 'profile',
  copilot: 'copilot'
};

const CAPABILITY_FOCUS_PATTERNS: ReadonlyArray<readonly [RegExp, McpCapabilityFocus]> = [
  [/\bshare\s*point\b/i, 'sharepoint'],
  [/\bone\s*drive\b/i, 'onedrive'],
  [/\boutlook\b/i, 'outlook'],
  [/\b(?:mail|email|emails|inbox)\b/i, 'mail'],
  [/\b(?:calendar|calendars|meeting|meetings)\b/i, 'calendar'],
  [/\b(?:teams|team|chat|chats|channel|channels)\b/i, 'teams'],
  [/\bword\b/i, 'word'],
  [/\b(?:profile|about me)\b/i, 'profile'],
  [/\bcopilot\b/i, 'copilot']
];

const PERMISSION_READ_VERBS: ReadonlyArray<string> = ['get', 'list', 'read', 'view', 'show', 'inspect', 'check'];
const PERMISSION_WRITE_VERBS: ReadonlyArray<string> = ['share', 'grant', 'invite', 'set', 'update', 'change', 'assign'];
const PERMISSION_READ_PATTERNS: ReadonlyArray<RegExp> = [
  /\b(?:get|list|read|view|show|inspect|check)\w*\b[^.]{0,40}\bpermissions?\b/i,
  /\bpermissions?\b[^.]{0,40}\b(?:get|list|read|view|show|inspect|check)\w*\b/i
];

// ─── Helpers ────────────────────────────────────────────────────

/**
 * Build a full gateway URL for a server by filling in the environment ID.
 */
export function resolveServerUrl(serverId: string, envId: string): string {
  const entry = M365_MCP_CATALOG.find((e) => e.id === serverId);
  if (!entry) {
    return `${GATEWAY_BASE.replace('{envId}', envId)}/${serverId}`;
  }
  return entry.urlTemplate.replace('{envId}', envId);
}

/**
 * Find which catalog server owns a given tool name.
 */
export function findServerForTool(toolName: string): IMcpCatalogEntry | undefined {
  const lower = toolName.toLowerCase();
  return M365_MCP_CATALOG.find((entry) =>
    entry.tools.some((t) => t.name.toLowerCase() === lower)
  );
}

/**
 * Get a catalog entry by ID.
 */
export function getCatalogEntry(id: string): IMcpCatalogEntry | undefined {
  return M365_MCP_CATALOG.find((entry) => entry.id === id);
}

export function normalizeCapabilityFocus(raw?: string): McpCapabilityFocus | undefined {
  if (!raw) {
    return undefined;
  }
  const normalized = raw.trim().toLowerCase().replace(/\s+/g, ' ');
  return CAPABILITY_FOCUS_NORMALIZATION[normalized];
}

export function detectCapabilityFocus(text: string): McpCapabilityFocus | undefined {
  const normalized = text.trim();
  if (!normalized) {
    return undefined;
  }

  for (let i = 0; i < CAPABILITY_FOCUS_PATTERNS.length; i++) {
    const [pattern, focus] = CAPABILITY_FOCUS_PATTERNS[i];
    if (pattern.test(normalized)) {
      return focus;
    }
  }

  return undefined;
}

export function getCapabilityFocusLabel(focus: McpCapabilityFocus): string {
  return CAPABILITY_FOCUS_LABELS[focus];
}

export function getCatalogEntriesForCapabilityFocus(focus: McpCapabilityFocus): IMcpCatalogEntry[] {
  const serverIds = CAPABILITY_FOCUS_SERVER_IDS[focus];
  return serverIds
    .map((serverId) => getCatalogEntry(serverId))
    .filter((entry): entry is IMcpCatalogEntry => entry !== undefined);
}

function looksLikePermissionReader(name: string, description: string): boolean {
  const nameHaystack = name.toLowerCase();
  if (!nameHaystack.includes('permission')) {
    return false;
  }
  const haystack = `${name} ${description}`.toLowerCase();

  const hasReadVerb = PERMISSION_READ_VERBS.some((verb) => haystack.includes(verb));
  if (!hasReadVerb) {
    return false;
  }

  if (PERMISSION_WRITE_VERBS.some((verb) => haystack.includes(verb))) {
    return false;
  }

  return PERMISSION_READ_PATTERNS.some((pattern) => pattern.test(`${name} ${description}`));
}

export function findCapabilityPermissionInspectionTool<T extends { name: string; description?: string }>(
  tools: ReadonlyArray<T>
): T | undefined {
  for (let i = 0; i < tools.length; i++) {
    const tool = tools[i];
    if (looksLikePermissionReader(tool.name, tool.description || '')) {
      return tool;
    }
  }
  return undefined;
}

export function catalogSupportsPermissionInspection(focus?: McpCapabilityFocus): boolean {
  const entries = focus
    ? getCatalogEntriesForCapabilityFocus(focus)
    : M365_MCP_CATALOG;

  return entries.some((entry) => findCapabilityPermissionInspectionTool(entry.tools) !== undefined);
}

/**
 * Get total tool count across all catalog servers.
 */
export function getTotalToolCount(): number {
  return M365_MCP_CATALOG.reduce((sum, entry) => sum + entry.tools.length, 0);
}

/**
 * Get the Grimoire virtual catalog entry (for display alongside M365 servers).
 */
export function getGrimoireServer(): IMcpCatalogEntry {
  return GRIMOIRE_SERVER;
}

/**
 * Get a condensed summary of the catalog for inclusion in the system prompt.
 * Conditional: skips M365 servers if mcpEnvironmentId is not configured.
 */
export function getCatalogSummaryForPrompt(config?: {
  mcpEnvironmentId?: string;
}): string {
  const lines: string[] = [];

  // M365 MCP servers — only include if mcpEnvironmentId is configured
  if (config?.mcpEnvironmentId) {
    M365_MCP_CATALOG.forEach((server) => {
      const toolNames = server.tools.map((t) => t.name).join(', ');
      lines.push(`- **${server.name}** (${server.id}): ${server.description}\n  Tools: ${toolNames}`);
    });
  } else {
    lines.push('*M365 MCP servers are not configured (no environment ID). The user can configure this in the web part settings.*');
  }

  return lines.join('\n');
}
