import { getTools } from '../realtime/ToolRegistry';
import {
  M365_MCP_CATALOG,
  findCapabilityPermissionInspectionTool,
  getCapabilityFocusLabel,
  getCatalogEntriesForCapabilityFocus,
  resolveServerUrl
} from '../../models/McpServerCatalog';
import type {
  IMcpCatalogEntry,
  IMcpCatalogTool,
  McpCapabilityFocus
} from '../../models/McpServerCatalog';
import type { IMcpConnection, IMcpTool } from '../../models/IMcpTypes';
import { getMcpAdapterBehavior } from '../mcp/McpExecutionAdapter';
import { extractServerIdFromUrl } from '../mcp/mcpUtils';
import type { IFunctionCallStore } from './ToolRuntimeContracts';

export interface IUserFacingCapabilityGroup {
  title: string;
  summary: string;
  summaryWithoutPermissionInspection?: string;
  toolNames: string[];
}

export interface IMcpToolDescriptor {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
}

export interface ICapabilityServerView {
  entry: IMcpCatalogEntry;
  tools: IMcpToolDescriptor[];
  source: 'live' | 'catalog';
}

export interface ICapabilityRuntimeHelpSummary {
  families: string[];
  aliasHints: string[];
  targetResolution: string[];
  resultShaping: string[];
}

export const USER_FACING_CAPABILITY_GROUPS: ReadonlyArray<IUserFacingCapabilityGroup> = [
  {
    title: 'Find information',
    summary: 'Search documents, people, sites, recent work, trending files, and emails.',
    toolNames: [
      'search_sharepoint',
      'search_people',
      'search_sites',
      'search_emails',
      'get_recent_documents',
      'get_trending_documents'
    ]
  },
  {
    title: 'Open and inspect content',
    summary: 'Browse libraries and view file details, site details, list items, permissions, and recent activity.',
    summaryWithoutPermissionInspection: 'Browse libraries and view file details, site details, list items, and recent activity.',
    toolNames: [
      'browse_document_library',
      'show_file_details',
      'show_site_info',
      'show_list_items',
      'show_permissions',
      'show_activity_feed'
    ]
  },
  {
    title: 'Read and summarize',
    summary: 'Read files, emails, and Teams conversations in full or answer questions about them.',
    toolNames: [
      'read_file_content',
      'read_email_content',
      'read_teams_messages'
    ]
  },
  {
    title: 'Profile and memory',
    summary: 'Use your profile and save, recall, or delete notes for later.',
    toolNames: [
      'get_my_profile',
      'save_note',
      'recall_notes',
      'delete_note'
    ]
  },
  {
    title: 'Create with guided forms',
    summary: 'Prepare emails, meetings, Teams messages, files, folders, chats, channels, and list items with guided forms.',
    toolNames: ['show_compose_form']
  },
  {
    title: 'Go deeper into Microsoft 365',
    summary: 'Use the broader Microsoft 365 action catalog when a task needs advanced mailbox, calendar, Teams, SharePoint, or Word operations.',
    toolNames: [
      'use_m365_capability',
      'list_m365_servers'
    ]
  }
];

export function dedupeStringList(values: string[]): string[] {
  const seen = new Set<string>();
  const deduped: string[] = [];

  values.forEach((value) => {
    const trimmed = value.trim();
    if (!trimmed || seen.has(trimmed)) {
      return;
    }
    seen.add(trimmed);
    deduped.push(trimmed);
  });

  return deduped;
}

export function getComposePresetCount(tools: ReturnType<typeof getTools>): number {
  const composeTool = tools.find((tool) => tool.name === 'show_compose_form');
  const presets = composeTool?.parameters.properties.preset?.enum;
  return Array.isArray(presets) ? presets.length : 0;
}

export function toMcpToolDescriptor(tool: IMcpCatalogTool | IMcpTool): IMcpToolDescriptor {
  return {
    name: tool.name,
    description: tool.description || '',
    inputSchema: tool.inputSchema || {}
  };
}

export function findConnectedServerSession(
  connections: IMcpConnection[],
  serverId: string,
  mcpEnvironmentId?: string
): IMcpConnection | undefined {
  const expectedUrl = mcpEnvironmentId ? resolveServerUrl(serverId, mcpEnvironmentId) : undefined;
  for (let i = connections.length - 1; i >= 0; i--) {
    const connection = connections[i];
    if (connection.state !== 'connected') {
      continue;
    }
    if ((expectedUrl && connection.serverUrl === expectedUrl) || extractServerIdFromUrl(connection.serverUrl) === serverId) {
      return connection;
    }
  }
  return undefined;
}

export function buildCapabilityServerViews(
  store: IFunctionCallStore,
  focus?: McpCapabilityFocus
): ICapabilityServerView[] {
  const entries = focus ? getCatalogEntriesForCapabilityFocus(focus) : M365_MCP_CATALOG;
  return entries.map((entry) => {
    const liveConnection = findConnectedServerSession(store.mcpConnections, entry.id, store.mcpEnvironmentId);
    const tools = liveConnection
      ? liveConnection.tools.map((tool) => toMcpToolDescriptor(tool))
      : entry.tools.map((tool) => toMcpToolDescriptor(tool));

    return {
      entry,
      tools,
      source: liveConnection ? 'live' : 'catalog'
    };
  });
}

export function normalizeCapabilityDescription(description: string): string {
  return description
    .replace(/\s+/g, ' ')
    .replace(/\.([A-Z])/g, '. $1')
    .replace(/:([A-Z])/g, ': $1')
    .trim();
}

export function truncateCapabilityDescription(description: string, maxLength: number = 220): string {
  const normalized = normalizeCapabilityDescription(description);
  if (normalized.length <= maxLength) {
    return normalized;
  }
  const clipped = normalized.slice(0, maxLength);
  const boundary = clipped.lastIndexOf(' ');
  return `${(boundary > 80 ? clipped.slice(0, boundary) : clipped).trim()}...`;
}

export function summarizeToolInputs(inputSchema: Record<string, unknown>): string | undefined {
  const properties = (inputSchema as { properties?: Record<string, unknown> }).properties || {};
  const requiredFields = (inputSchema as { required?: unknown[] }).required;
  const required = Array.isArray(requiredFields)
    ? requiredFields.filter((value): value is string => typeof value === 'string')
    : [];
  const propertyNames = Object.keys(properties);
  if (propertyNames.length === 0 && required.length === 0) {
    return undefined;
  }

  const parts: string[] = [];
  if (required.length > 0) {
    const renderedRequired = required.slice(0, 3).map((name) => `\`${name}\``).join(', ');
    parts.push(`Required: ${renderedRequired}${required.length > 3 ? ', ...' : ''}.`);
  }
  if (propertyNames.length > 0) {
    const renderedNames = propertyNames.slice(0, 3).map((name) => `\`${name}\``).join(', ');
    parts.push(`Inputs: ${renderedNames}${propertyNames.length > 3 ? ', ...' : ''}.`);
  }
  return parts.join(' ');
}

export function formatCapabilityFieldList(fieldNames: string[], maxItems: number = 8): string | undefined {
  if (fieldNames.length === 0) {
    return undefined;
  }

  const rendered = fieldNames.slice(0, maxItems).map((name) => `\`${name}\``).join(', ');
  return fieldNames.length > maxItems ? `${rendered}, ...` : rendered;
}

export function summarizeCapabilityRequiredInputs(tools: IMcpToolDescriptor[]): string[] {
  const requiredFields = dedupeStringList(
    tools.flatMap((tool) => {
      const required = (tool.inputSchema as { required?: unknown[] }).required;
      return Array.isArray(required)
        ? required.filter((value): value is string => typeof value === 'string')
        : [];
    })
  );

  return requiredFields.sort((left, right) => left.localeCompare(right));
}

export function summarizeCapabilitySchemaFields(tools: IMcpToolDescriptor[]): string[] {
  const fieldNames = dedupeStringList(
    tools.flatMap((tool) => {
      const properties = (tool.inputSchema as { properties?: Record<string, unknown> }).properties || {};
      return Object.keys(properties);
    })
  );

  return fieldNames.sort((left, right) => left.localeCompare(right));
}

export function summarizeCapabilityRuntimeHelp(serverId: string, tools: IMcpToolDescriptor[]): ICapabilityRuntimeHelpSummary {
  const families = dedupeStringList(tools.map((tool) => getMcpAdapterBehavior(serverId, tool.name).family));
  const aliasHints = dedupeStringList(tools.flatMap((tool) => getMcpAdapterBehavior(serverId, tool.name).aliasHints));
  const targetResolution = dedupeStringList(tools.flatMap((tool) => getMcpAdapterBehavior(serverId, tool.name).targetResolution));
  const resultShaping = dedupeStringList(tools.flatMap((tool) => getMcpAdapterBehavior(serverId, tool.name).resultShaping));

  return {
    families,
    aliasHints,
    targetResolution,
    resultShaping
  };
}

export function formatCapabilityToolLine(serverId: string, tool: IMcpToolDescriptor): string {
  const description = truncateCapabilityDescription(tool.description);
  const inputSummary = summarizeToolInputs(tool.inputSchema);
  const adapterBehavior = getMcpAdapterBehavior(serverId, tool.name);
  const runtimeHelpParts = dedupeStringList([
    adapterBehavior.aliasHints[0],
    adapterBehavior.targetResolution[0],
    adapterBehavior.resultShaping[0]
  ]).slice(0, 3);
  const runtimeHelp = runtimeHelpParts.join('; ');
  return `- \`${tool.name}\`: ${description}${inputSummary ? ` ${inputSummary}` : ''}${runtimeHelp ? ` Runtime help: ${runtimeHelp}.` : ''}`;
}

export function getCapabilityGroupToolNames(
  group: IUserFacingCapabilityGroup,
  permissionInspectionSupported: boolean
): string[] {
  return group.toolNames.filter((toolName) => toolName !== 'show_permissions' || permissionInspectionSupported);
}

export function buildCapabilityOverviewMarkdown(
  tools: ReturnType<typeof getTools>,
  hasMcpEnv: boolean,
  serverViews: ICapabilityServerView[]
): { content: string; userFacingCount: number; internalToolCount: number } {
  const availableToolNames = new Set(tools.map((tool) => tool.name));
  const formPresetCount = getComposePresetCount(tools);
  const permissionInspectionSupported = serverViews.some((serverView) =>
    findCapabilityPermissionInspectionTool(serverView.tools) !== undefined
  );

  const builtInLines: string[] = [
    '## What Grimoire Can Help With',
    ''
  ];

  let userFacingCount = 0;
  USER_FACING_CAPABILITY_GROUPS.forEach((group) => {
    const visibleGroupToolNames = getCapabilityGroupToolNames(group, permissionInspectionSupported);
    const availableCount = visibleGroupToolNames.filter((toolName) => availableToolNames.has(toolName)).length;
    userFacingCount += availableCount;

    const baseSummary = permissionInspectionSupported || !group.summaryWithoutPermissionInspection
      ? group.summary
      : group.summaryWithoutPermissionInspection;
    const summary = group.title === 'Create with guided forms' && formPresetCount > 0
      ? `${baseSummary.replace(/\.$/, '')} (${formPresetCount} form presets available).`
      : baseSummary;

    builtInLines.push(`- **${group.title}** (${availableCount} actions): ${summary}`);
  });

  const internalToolCount = Math.max(tools.length - userFacingCount, 0);
  builtInLines.push('');
  builtInLines.push(`*Behind the scenes: ${userFacingCount} user-facing actions plus ${internalToolCount} internal UI and runtime helpers.*`);

  const mcpLines: string[] = [];
  if (hasMcpEnv) {
    const totalMcpTools = serverViews.reduce((sum, server) => sum + server.tools.length, 0);
    mcpLines.push(`\n## Connected Microsoft 365 Services (${serverViews.length} services, ${totalMcpTools} actions behind the scenes)\n`);
    serverViews.forEach((serverView) => {
      const sourceLabel = serverView.source === 'live' ? 'live session' : 'catalog fallback';
      const requiredInputs = summarizeCapabilityRequiredInputs(serverView.tools);
      const renderedRequiredInputs = formatCapabilityFieldList(requiredInputs, 3);
      const summarySuffix = renderedRequiredInputs ? ` Common required inputs: ${renderedRequiredInputs}.` : '';
      mcpLines.push(`- **${serverView.entry.name}**: ${serverView.entry.description} (${serverView.tools.length} actions, ${sourceLabel}).${summarySuffix}`);
    });
  } else {
    mcpLines.push('\n## Connected Microsoft 365 Services\n');
    mcpLines.push('*Not configured yet. Add the MCP Environment ID in the web part settings to connect SharePoint, Outlook, Teams, Calendar, Copilot Search, Word, and profile services.*');
  }

  return {
    content: [...builtInLines, ...mcpLines].join('\n'),
    userFacingCount,
    internalToolCount
  };
}

export function buildFocusedCapabilityMarkdown(
  focus: McpCapabilityFocus,
  serverViews: ICapabilityServerView[]
): { title: string; content: string; focusedToolCount: number; permissionInspectionSupported: boolean } {
  const focusLabel = getCapabilityFocusLabel(focus);
  const permissionInspectionSupported = serverViews.some((serverView) =>
    findCapabilityPermissionInspectionTool(serverView.tools) !== undefined
  );
  const focusedToolCount = serverViews.reduce((sum, serverView) => sum + serverView.tools.length, 0);

  const lines: string[] = [
    `## ${focusLabel} Capabilities`,
    '',
    '*Source of truth: live MCP sessions when connected, otherwise the generated discovery catalog.*',
    ''
  ];

  serverViews.forEach((serverView, index) => {
    const sourceLabel = serverView.source === 'live' ? 'live session' : 'catalog fallback';
    const requiredInputs = summarizeCapabilityRequiredInputs(serverView.tools);
    const schemaFields = summarizeCapabilitySchemaFields(serverView.tools);
    const runtimeHelp = summarizeCapabilityRuntimeHelp(serverView.entry.id, serverView.tools);
    lines.push(`### ${serverView.entry.name} (${serverView.tools.length} actions, ${sourceLabel})`);
    lines.push(serverView.entry.description);
    lines.push('');
    lines.push(`*Schema source: ${sourceLabel === 'live session' ? 'connected MCP session' : 'generated catalog fallback'}.*`);
    const renderedRequiredInputs = formatCapabilityFieldList(requiredInputs);
    if (renderedRequiredInputs) {
      lines.push(`**Common required inputs:** ${renderedRequiredInputs}`);
    } else {
      lines.push('**Common required inputs:** usually resolved from current context, selected items, or hidden form state.');
    }
    const renderedSchemaFields = formatCapabilityFieldList(schemaFields);
    if (renderedSchemaFields) {
      lines.push(`**Common schema fields:** ${renderedSchemaFields}`);
    }
    if (
      runtimeHelp.families.length > 0
      || runtimeHelp.aliasHints.length > 0
      || runtimeHelp.targetResolution.length > 0
      || runtimeHelp.resultShaping.length > 0
    ) {
      lines.push('**Generic runtime help applied by Grimoire**');
      if (runtimeHelp.families.length > 0) {
        lines.push(`- Families: ${runtimeHelp.families.map((family) => `\`${family}\``).join(', ')}`);
      }
      if (runtimeHelp.aliasHints.length > 0) {
        lines.push(`- Aliases: ${runtimeHelp.aliasHints.slice(0, 3).join('; ')}${runtimeHelp.aliasHints.length > 3 ? '; ...' : ''}`);
      }
      if (runtimeHelp.targetResolution.length > 0) {
        lines.push(`- Target resolution: ${runtimeHelp.targetResolution.slice(0, 3).join('; ')}${runtimeHelp.targetResolution.length > 3 ? '; ...' : ''}`);
      }
      if (runtimeHelp.resultShaping.length > 0) {
        lines.push(`- Result shaping: ${runtimeHelp.resultShaping.slice(0, 3).join('; ')}${runtimeHelp.resultShaping.length > 3 ? '; ...' : ''}`);
      }
    }
    lines.push('');
    serverView.tools.forEach((tool) => {
      lines.push(formatCapabilityToolLine(serverView.entry.id, tool));
    });
    if (index < serverViews.length - 1) {
      lines.push('');
    }
  });

  if ((focus === 'sharepoint' || focus === 'onedrive') && !permissionInspectionSupported) {
    lines.push('');
    lines.push('### Current Limitation');
    lines.push('The current SharePoint & OneDrive MCP tool set exposes sharing changes, but it does not expose a permission-read tool yet. Grimoire can modify sharing through MCP, but it should not claim that it can inspect permissions until that tool exists.');
  }

  return {
    title: `${focusLabel} Capabilities`,
    content: lines.join('\n'),
    focusedToolCount,
    permissionInspectionSupported
  };
}
