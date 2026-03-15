import type { BlockType } from '../../models/IBlock';
import type { IMcpConnection, IMcpContent } from '../../models/IMcpTypes';
import { getCatalogEntry } from '../../models/McpServerCatalog';
import type { IHieArtifactRecord, IHieSourceContext, IHieTaskContext } from '../hie/HIETypes';
import { logService } from '../logging/LogService';
import type { McpClientService } from './McpClientService';
import { withMcpRetry } from './mcpUtils';
import {
  deriveMcpTargetContextFromHie,
  deriveMcpTargetContextFromUnknown,
  describeMcpTargetContext,
  mergeMcpTargetContexts,
  type IMcpTargetContext,
  type McpTargetSource
} from './McpTargetContext';
import { pickFirstNonEmptyString } from './OdspLocationResolver';

export interface IMcpAdapterBehavior {
  family: 'profile' | 'sharepoint-onedrive' | 'mail-calendar' | 'teams' | 'copilot' | 'generic';
  aliasHints: string[];
  targetResolution: string[];
  resultShaping: string[];
}

export interface IMcpExecutionTrace {
  serverId?: string;
  serverName?: string;
  toolName: string;
  requiredFields?: string[];
  rawArgs: Record<string, unknown>;
  normalizedArgs?: Record<string, unknown>;
  resolvedArgs?: Record<string, unknown>;
  targetContext?: IMcpTargetContext;
  targetSource: McpTargetSource;
  targetSummary?: string;
  recoverySteps: string[];
  unwrapPath?: string[];
  inferredBlockType?: BlockType;
  finalBlockTitle?: string;
  finalSummary?: string;
}

interface IResolveMcpTargetContextOptions {
  explicitTargetContext?: IMcpTargetContext;
  sourceContext?: IHieSourceContext;
  taskContext?: IHieTaskContext;
  artifacts?: Readonly<Record<string, IHieArtifactRecord>>;
  currentSiteUrl?: string;
}

interface IValidationResult {
  matched: Record<string, unknown>;
  missing: string[];
  extra: Record<string, unknown>;
  typeErrors: string[];
}

interface IResolutionStrategy {
  resolves: string[];
  trigger: (extra: Record<string, unknown>) => string | undefined;
  resolutionTools: string[];
  buildArgs: (triggerValue: string, toolSchema: Record<string, unknown>) => Record<string, unknown>;
  extract: (parsed: Record<string, unknown>) => Record<string, unknown>;
}

export interface IMcpSessionRuntimeHelpers {
  findExistingSession: (connections: IMcpConnection[], serverUrl: string) => string | undefined;
  connectToM365Server: (
    mcpClient: McpClientService,
    serverUrl: string,
    serverName: string,
    getToken: ((resource: string) => Promise<string>) | undefined
  ) => Promise<string>;
}

export interface IMcpResolvedArgsEnricherContext {
  toolName: string;
  serverId: string;
  serverName: string;
  sessionId: string;
  mcpClient: McpClientService;
  targetContext?: IMcpTargetContext;
  targetSource: McpTargetSource;
}

export interface IMcpResolvedArgsEnricherResult {
  args: Record<string, unknown>;
  recoverySteps?: string[];
}

export interface IExecuteCatalogMcpToolOptions {
  serverId: string;
  serverName: string;
  serverUrl: string;
  toolName: string;
  rawArgs: Record<string, unknown>;
  connections: IMcpConnection[];
  mcpClient: McpClientService;
  sessionHelpers: IMcpSessionRuntimeHelpers;
  getToken?: (resource: string) => Promise<string>;
  getConnections?: () => IMcpConnection[];
  explicitTargetContext?: IMcpTargetContext;
  sourceContext?: IHieSourceContext;
  taskContext?: IHieTaskContext;
  artifacts?: Readonly<Record<string, IHieArtifactRecord>>;
  currentSiteUrl?: string;
  webGroundingEnabled?: boolean;
  enrichResolvedArgs?: (context: IMcpResolvedArgsEnricherContext & { args: Record<string, unknown> }) =>
    Promise<IMcpResolvedArgsEnricherResult>;
}

export interface IExecuteCatalogMcpToolResult {
  success: boolean;
  serverId: string;
  serverName: string;
  serverUrl: string;
  sessionId: string;
  realToolName: string;
  connection?: IMcpConnection;
  requiredFields: string[];
  schemaProps: Record<string, { type?: string; description?: string }>;
  normalizedArgs: Record<string, unknown>;
  resolvedArgs: Record<string, unknown>;
  targetContext?: IMcpTargetContext;
  targetSource: McpTargetSource;
  recoverySteps: string[];
  mcpResult?: { success: boolean; content: IMcpContent[]; error?: string };
  error?: string;
  trace: IMcpExecutionTrace;
}

export interface IEnsureCatalogMcpConnectionOptions {
  serverId: string;
  serverName: string;
  serverUrl: string;
  connections: IMcpConnection[];
  mcpClient: McpClientService;
  sessionHelpers: IMcpSessionRuntimeHelpers;
  getToken?: (resource: string) => Promise<string>;
  getConnections?: () => IMcpConnection[];
}

export interface IEnsureCatalogMcpConnectionResult {
  success: boolean;
  sessionId?: string;
  connection?: IMcpConnection;
  liveConnection?: IMcpConnection;
  fallbackConnection?: IMcpConnection;
  recoverySteps: string[];
  error?: string;
}

function buildCatalogConnectionFallback(
  serverId: string,
  serverUrl: string,
  serverName: string,
  sessionId: string
): IMcpConnection | undefined {
  const catalogEntry = getCatalogEntry(serverId);
  if (!catalogEntry) {
    return undefined;
  }

  return {
    sessionId,
    serverUrl,
    serverName,
    state: 'connected',
    connectedAt: new Date(0),
    tools: catalogEntry.tools.map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema
    }))
  };
}

const VALID_LIST_TEMPLATES: Record<string, string> = {
  genericlist: 'genericList',
  documentlibrary: 'documentLibrary',
  events: 'events',
  tasks: 'tasks',
  survey: 'survey',
  links: 'links',
  announcements: 'announcements',
  contacts: 'contacts',
  accessrequest: 'accessRequest',
  issuetracking: 'issueTracking'
};

function normalizeSharePointListTemplate(template: unknown): string | undefined {
  if (typeof template !== 'string') {
    return undefined;
  }
  const trimmed = template.trim();
  if (!trimmed) {
    return undefined;
  }
  const key = trimmed.toLowerCase().replace(/[\s_-]+/g, '');
  const match = VALID_LIST_TEMPLATES[key];
  if (match) {
    return match;
  }
  logService.warning('mcp', `Unrecognized list template "${trimmed}", defaulting to genericList`);
  return 'genericList';
}

function sanitizeSharePointColumnInternalName(name: unknown): string | undefined {
  if (typeof name !== 'string') {
    return undefined;
  }

  const trimmed = name.trim();
  if (!trimmed) {
    return undefined;
  }

  const words = trimmed
    .replace(/[^A-Za-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (words.length === 0) {
    return undefined;
  }

  const collapsed = words
    .map((word, index) => {
      const lower = word.toLowerCase();
      return index === 0
        ? lower
        : lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join('')
    .replace(/[^A-Za-z0-9_]/g, '');

  if (!collapsed) {
    return undefined;
  }

  return /^[A-Za-z_]/.test(collapsed) ? collapsed : `field${collapsed}`;
}

function parseSharePointChoiceValues(value: unknown): string[] | undefined {
  if (Array.isArray(value)) {
    const normalized = value
      .filter((entry): entry is string => typeof entry === 'string')
      .map((entry) => entry.trim())
      .filter(Boolean);
    return normalized.length > 0 ? Array.from(new Set(normalized)) : undefined;
  }

  if (typeof value !== 'string') {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const separator = trimmed.indexOf('\n') !== -1 ? /\r?\n/ : /\s*,\s*/;
  const normalized = trimmed
    .split(separator)
    .map((entry) => entry.trim())
    .filter(Boolean);
  return normalized.length > 0 ? Array.from(new Set(normalized)) : undefined;
}

function normalizeSharePointColumnType(type: unknown): { key: string; value: Record<string, unknown> } | undefined {
  if (typeof type !== 'string') {
    return undefined;
  }

  const trimmed = type.trim();
  if (!trimmed) {
    return undefined;
  }

  const normalized = trimmed.toLowerCase().replace(/[^a-z0-9]+/g, '');

  if (
    normalized === 'text'
    || normalized === 'string'
    || normalized === 'lineoftext'
    || normalized === 'singleline'
    || normalized === 'singlelinetext'
    || normalized === 'singlelineoftext'
    || normalized === 'plaintext'
  ) {
    return { key: 'text', value: {} };
  }

  if (
    normalized === 'notes'
    || normalized === 'note'
    || normalized === 'multiplelinesoftext'
    || normalized === 'multiline'
    || normalized === 'multilinetext'
    || normalized === 'multilineoftext'
    || normalized === 'paragraph'
    || normalized === 'longtext'
  ) {
    return { key: 'text', value: { allowMultipleLines: true } };
  }

  if (
    normalized === 'number'
    || normalized === 'numeric'
    || normalized === 'integer'
    || normalized === 'decimal'
    || normalized === 'currency'
  ) {
    return { key: 'number', value: {} };
  }

  if (
    normalized === 'choice'
    || normalized === 'choices'
    || normalized === 'option'
    || normalized === 'options'
    || normalized === 'dropdown'
  ) {
    return { key: 'choice', value: {} };
  }

  if (
    normalized === 'boolean'
    || normalized === 'bool'
    || normalized === 'yesno'
    || normalized === 'checkbox'
  ) {
    return { key: 'boolean', value: {} };
  }

  if (
    normalized === 'date'
    || normalized === 'datetime'
    || normalized === 'dateandtime'
    || normalized === 'time'
  ) {
    return { key: 'dateTime', value: {} };
  }

  if (
    normalized === 'person'
    || normalized === 'people'
    || normalized === 'user'
    || normalized === 'personorgroup'
    || normalized === 'group'
  ) {
    return { key: 'personOrGroup', value: {} };
  }

  if (
    normalized === 'url'
    || normalized === 'link'
    || normalized === 'linkorpicture'
    || normalized === 'hyperlink'
    || normalized === 'picture'
    || normalized === 'image'
    || normalized === 'hyperlinkorpicture'
  ) {
    return { key: 'hyperlinkOrPicture', value: {} };
  }

  return undefined;
}

function getBrowserSiteContext(): { hostname?: string; siteUrl?: string } {
  if (typeof window === 'undefined' || !window.location) {
    return {};
  }

  const { hostname, origin, pathname } = window.location;
  if (!hostname || hostname === 'localhost' || hostname === '127.0.0.1') {
    return {};
  }

  const siteRootMatch = pathname.match(/^\/(?:sites|teams)\/[^/]+/i);
  const siteRoot = siteRootMatch?.[0];
  return {
    hostname,
    siteUrl: siteRoot ? `${origin}${siteRoot}` : undefined
  };
}

function parseSharePointSiteReference(
  reference: string | undefined,
  fallbackSiteUrl?: string
): { hostname?: string; serverRelativePath?: string } {
  const candidate = reference?.trim();
  if (!candidate) return {};

  const normalizePath = (rawPath: string): string | undefined => {
    const trimmed = rawPath.trim();
    if (!trimmed) return undefined;
    const withoutQuery = trimmed.split('?')[0].split('#')[0];
    const normalized = withoutQuery.replace(/^\/+/, '').replace(/\/+$/, '');
    return normalized || undefined;
  };

  try {
    const url = new URL(candidate);
    return {
      hostname: url.hostname || undefined,
      serverRelativePath: normalizePath(url.pathname)
    };
  } catch {
    // Fall through to path-only handling.
  }

  const fallbackHostname = (() => {
    try {
      if (fallbackSiteUrl) {
        return new URL(fallbackSiteUrl).hostname;
      }
      return getBrowserSiteContext().hostname;
    } catch {
      return getBrowserSiteContext().hostname;
    }
  })();

  return {
    hostname: fallbackHostname,
    serverRelativePath: normalizePath(candidate)
  };
}

function tryParseJson(text: string): unknown | undefined {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function toRecordArray(payload: unknown): Record<string, unknown>[] {
  if (Array.isArray(payload)) {
    return payload.filter((item): item is Record<string, unknown> =>
      typeof item === 'object' && item !== null && !Array.isArray(item)
    );
  }

  if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
    const container = (payload as Record<string, unknown>).value;
    if (Array.isArray(container)) {
      return container.filter((item): item is Record<string, unknown> =>
        typeof item === 'object' && item !== null && !Array.isArray(item)
      );
    }
  }

  return [];
}

function normalizeLookupValue(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }

  return value
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function validateMcpArgs(
  llmArgs: Record<string, unknown>,
  realSchema: Record<string, unknown>
): IValidationResult {
  const schemaProps = (realSchema as { properties?: Record<string, { type?: string }> }).properties || {};
  const required = (realSchema as { required?: string[] }).required || [];
  const realPropNames = Object.keys(schemaProps);

  const matched: Record<string, unknown> = {};
  const extra: Record<string, unknown> = {};
  const typeErrors: string[] = [];

  const lowerToReal = new Map<string, string>();
  realPropNames.forEach((prop) => lowerToReal.set(prop.toLowerCase(), prop));

  const llmKeys = Object.keys(llmArgs);
  llmKeys.forEach((llmKey) => {
    const realKey = lowerToReal.get(llmKey.toLowerCase());
    if (realKey) {
      const expectedType = schemaProps[realKey]?.type;
      const value = llmArgs[llmKey];
      if (expectedType && value !== undefined && value !== null) {
        const actualType = typeof value;
        const typeOk =
          (expectedType === 'string' && actualType === 'string') ||
          (expectedType === 'number' && actualType === 'number') ||
          (expectedType === 'integer' && actualType === 'number') ||
          (expectedType === 'boolean' && actualType === 'boolean') ||
          (expectedType === 'array' && Array.isArray(value)) ||
          (expectedType === 'object' && actualType === 'object' && !Array.isArray(value));
        if (!typeOk) {
          typeErrors.push(`${realKey}: expected ${expectedType}, got ${actualType}`);
          if (expectedType === 'array' && actualType === 'string') {
            matched[realKey] = [value as string];
            return;
          }
        }
      }
      matched[realKey] = llmArgs[llmKey];
    } else {
      extra[llmKey] = llmArgs[llmKey];
    }
  });

  const missing = required.filter((field) =>
    matched[field] === undefined || matched[field] === null || matched[field] === ''
  );
  return { matched, missing, extra, typeErrors };
}

function findUrlInExtra(extra: Record<string, unknown>): string | undefined {
  const values = Object.values(extra);
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (typeof value === 'string' && (value.includes('sharepoint.com') || value.includes('/sites/') || value.startsWith('http'))) {
      return value;
    }
  }
  return undefined;
}

function findSiteLookupTrigger(extra: Record<string, unknown>): string | undefined {
  return pickFirstNonEmptyString(
    extra.siteName,
    extra.site_name,
    extra.siteTitle,
    extra.site,
    extra.search,
    extra.searchQuery,
    extra.query,
    extra.name,
    extra.displayName
  );
}

const ODSP_RESOLUTION_STRATEGIES: IResolutionStrategy[] = [
  {
    resolves: ['documentlibraryid', 'driveid', 'fileorfolderid', 'fileid', 'itemid', 'folderid'],
    trigger: findUrlInExtra,
    resolutionTools: ['getFileOrFolderMetadataByUrl'],
    buildArgs: (url, toolSchema) => {
      const props = (toolSchema as { properties?: Record<string, unknown> }).properties || {};
      const propNames = Object.keys(props);
      const urlParam = propNames.find((prop) => prop.toLowerCase().includes('url')) || propNames[0] || 'url';
      return { [urlParam]: url };
    },
    extract: (parsed) => {
      const parentReference = parsed.parentReference as Record<string, unknown> | undefined;
      const driveId = (parentReference?.driveId || parsed.driveId || parsed.documentLibraryId) as string | undefined;
      const itemId = (parsed.id || parsed.itemId || parsed.fileId) as string | undefined;
      const resolved: Record<string, unknown> = {};
      if (driveId) resolved.documentlibraryid = driveId;
      if (driveId) resolved.driveid = driveId;
      if (itemId) resolved.fileorfolderid = itemId;
      if (itemId) resolved.fileid = itemId;
      if (itemId) resolved.itemid = itemId;
      if (itemId) resolved.folderid = itemId;
      return resolved;
    }
  },
  {
    resolves: ['siteid'],
    trigger: (extra) => pickFirstNonEmptyString(
      extra.siteUrl,
      extra.site_url,
      extra.url,
      extra.webUrl,
      extra.currentSiteUrl,
      extra.sitePath,
      extra.site_path,
      extra.path
    ),
    resolutionTools: ['getSiteByPath'],
    buildArgs: (siteReference, toolSchema) => {
      const props = (toolSchema as { properties?: Record<string, unknown> }).properties || {};
      const propNames = Object.keys(props);
      const hostnameField = propNames.find((prop) => prop.toLowerCase() === 'hostname') || 'hostname';
      const sitePathField = propNames.find((prop) => prop.toLowerCase() === 'serverrelativepath') || 'serverRelativePath';
      const parsed = parseSharePointSiteReference(siteReference, pickFirstNonEmptyString(siteReference));
      return {
        [hostnameField]: parsed.hostname,
        [sitePathField]: parsed.serverRelativePath
      };
    },
    extract: (parsed) => {
      const siteId = typeof parsed.id === 'string' ? parsed.id : typeof parsed.siteId === 'string' ? parsed.siteId : undefined;
      return siteId ? { siteid: siteId } : {};
    }
  },
  {
    resolves: ['siteid'],
    trigger: findSiteLookupTrigger,
    resolutionTools: ['findSite', 'searchSitesByName'],
    buildArgs: (query, toolSchema) => {
      const props = (toolSchema as { properties?: Record<string, unknown> }).properties || {};
      const propNames = Object.keys(props);
      const queryParam = propNames.find((prop) =>
        prop.toLowerCase().includes('query') || prop.toLowerCase().includes('search')
      ) || propNames[0] || 'searchQuery';
      return { [queryParam]: query };
    },
    extract: (parsed) => {
      let siteId: string | undefined;
      if (typeof parsed.id === 'string') siteId = parsed.id;
      if (typeof parsed.siteId === 'string') siteId = parsed.siteId;
      if (Array.isArray(parsed)) {
        const first = parsed[0] as Record<string, unknown> | undefined;
        if (first && typeof first.id === 'string') siteId = first.id;
      }
      if (Array.isArray(parsed.value)) {
        const first = parsed.value[0] as Record<string, unknown> | undefined;
        if (first && typeof first.id === 'string') siteId = first.id;
      }
      return siteId ? { siteid: siteId } : {};
    }
  }
];

async function resolveMissingMcpParameters(
  missing: string[],
  extra: Record<string, unknown>,
  connectionTools: Array<{ name: string; inputSchema: Record<string, unknown> }>,
  mcpClient: McpClientService,
  sessionId: string
): Promise<Record<string, unknown>> {
  const resolved: Record<string, unknown> = {};
  const missingLower = missing.map((field) => field.toLowerCase());

  for (let i = 0; i < ODSP_RESOLUTION_STRATEGIES.length; i++) {
    const strategy = ODSP_RESOLUTION_STRATEGIES[i];
    const canResolve = strategy.resolves.some((field) => missingLower.includes(field));
    if (!canResolve) {
      continue;
    }

    const triggerValue = strategy.trigger(extra);
    if (!triggerValue) {
      continue;
    }

    const resolutionTool = connectionTools.find((tool) =>
      strategy.resolutionTools.some((toolName) => tool.name.toLowerCase().includes(toolName.toLowerCase()))
    );
    if (!resolutionTool) {
      logService.warning('mcp', `Resolution tool "${strategy.resolutionTools.join(', ')}" not found on session`);
      continue;
    }

    logService.info('mcp', `Resolving via ${resolutionTool.name} (trigger: ${triggerValue.substring(0, 80)})`);
    try {
      const resolutionArgs = strategy.buildArgs(triggerValue, resolutionTool.inputSchema);
      const resolutionResult = await mcpClient.execute(sessionId, resolutionTool.name, resolutionArgs);
      if (!resolutionResult.success) {
        logService.warning('mcp', `Resolution tool ${resolutionTool.name} failed: ${resolutionResult.error}`);
        continue;
      }

      const parsed = extractStructuredMcpPayload(resolutionResult.content).payload as Record<string, unknown> | undefined;
      if (!parsed) {
        continue;
      }

      const extracted = strategy.extract(parsed);
      Object.keys(extracted).forEach((key) => {
        resolved[key] = extracted[key];
      });
      logService.info('mcp', `Resolution extracted: ${JSON.stringify(extracted)}`);
    } catch (error) {
      logService.warning('mcp', `Resolution failed: ${(error as Error).message}`);
    }
  }

  return resolved;
}

async function resolveSharePointListId(
  args: {
    rawArgs: Record<string, unknown>;
    resolvedArgs: Record<string, unknown>;
    schemaProps: Record<string, { type?: string; description?: string }>;
    targetContext?: IMcpTargetContext;
  },
  connectionTools: Array<{ name: string; inputSchema: Record<string, unknown> }>,
  mcpClient: McpClientService,
  sessionId: string
): Promise<Record<string, unknown>> {
  const propNames = Object.keys(args.schemaProps);
  const getPropName = (target: string): string | undefined =>
    propNames.find((prop) => prop.toLowerCase() === target.toLowerCase());

  const siteIdField = getPropName('siteId');
  const listIdField = getPropName('listId');
  if (!siteIdField || !listIdField) {
    return {};
  }

  const resolvedListId = pickFirstNonEmptyString(args.resolvedArgs[listIdField]);
  if (resolvedListId) {
    return {};
  }

  const siteId = pickFirstNonEmptyString(
    args.resolvedArgs[siteIdField],
    args.rawArgs.siteId,
    args.rawArgs.site_id,
    args.targetContext?.siteId
  );
  if (!siteId) {
    return {};
  }

  const requestedListUrl = pickFirstNonEmptyString(
    args.rawArgs.listUrl,
    args.rawArgs.list_url,
    args.targetContext?.listUrl
  );
  const requestedListName = pickFirstNonEmptyString(
    args.rawArgs.listName,
    args.rawArgs.list_name,
    args.rawArgs.listTitle,
    args.rawArgs.list_title,
    args.rawArgs.list,
    args.targetContext?.listName
  );
  if (!requestedListUrl && !requestedListName) {
    return {};
  }

  const listListsTool = connectionTools.find((tool) => tool.name.toLowerCase() === 'listlists');
  if (!listListsTool) {
    logService.warning('mcp', 'Resolution tool "listLists" not found on session');
    return {};
  }

  const toolProps = (listListsTool.inputSchema as { properties?: Record<string, unknown> }).properties || {};
  const toolPropNames = Object.keys(toolProps);
  const listListsSiteIdField = toolPropNames.find((prop) => prop.toLowerCase() === 'siteid') || 'siteId';

  logService.info('mcp', `Resolving via ${listListsTool.name} (trigger: ${requestedListUrl || requestedListName || siteId})`);
  try {
    const resolutionResult = await mcpClient.execute(sessionId, listListsTool.name, {
      [listListsSiteIdField]: siteId
    });
    if (!resolutionResult.success) {
      logService.warning('mcp', `Resolution tool ${listListsTool.name} failed: ${resolutionResult.error}`);
      return {};
    }

    const parsed = extractStructuredMcpPayload(resolutionResult.content).payload;
    const items = toRecordArray(parsed);
    if (items.length === 0) {
      return {};
    }

    const normalizedRequestedUrl = requestedListUrl?.trim().toLowerCase();
    const normalizedRequestedName = normalizeLookupValue(requestedListName);
    const matchingList = items.find((item) => {
      const candidateUrl = pickFirstNonEmptyString(item.webUrl, item.listUrl, item.url)?.trim().toLowerCase();
      if (normalizedRequestedUrl && candidateUrl === normalizedRequestedUrl) {
        return true;
      }

      const candidateName = normalizeLookupValue(
        pickFirstNonEmptyString(item.displayName, item.name, item.title)
      );
      return !!(normalizedRequestedName && candidateName && candidateName === normalizedRequestedName);
    });

    if (!matchingList) {
      return {};
    }

    const listId = pickFirstNonEmptyString(matchingList.id, matchingList.listId);
    if (!listId) {
      return {};
    }

    const extracted: Record<string, unknown> = {
      listid: listId
    };
    const listUrl = pickFirstNonEmptyString(matchingList.webUrl, matchingList.listUrl, matchingList.url);
    const listName = pickFirstNonEmptyString(matchingList.displayName, matchingList.name, matchingList.title);
    if (listUrl) {
      extracted.listurl = listUrl;
    }
    if (listName) {
      extracted.listname = listName;
    }

    logService.info('mcp', `Resolution extracted: ${JSON.stringify(extracted)}`);
    return extracted;
  } catch (error) {
    logService.warning('mcp', `Resolution failed: ${(error as Error).message}`);
    return {};
  }
}

function normalizeCopilotChatArgs(
  toolName: string,
  args: Record<string, unknown>,
  rawArgs: Record<string, unknown>,
  webGroundingEnabled: boolean
): Record<string, unknown> {
  if (toolName.toLowerCase() !== 'copilot_chat') {
    return args;
  }

  const normalized = { ...args };
  const currentMessage = pickFirstNonEmptyString(normalized.message);
  if (!currentMessage) {
    const fallbackMessage = pickFirstNonEmptyString(
      normalized.query,
      normalized.prompt,
      normalized.text,
      normalized.request,
      rawArgs.message,
      rawArgs.query,
      rawArgs.prompt,
      rawArgs.text,
      rawArgs.request
    );
    const fallbackUrl = pickFirstNonEmptyString(
      normalized.url,
      rawArgs.url
    );

    if (fallbackMessage) {
      normalized.message = fallbackMessage;
    } else if (fallbackUrl) {
      normalized.message = `Research this URL and summarize the key points: ${fallbackUrl}`;
    }
  }

  if (typeof normalized.enableWebSearch !== 'boolean') {
    normalized.enableWebSearch = webGroundingEnabled;
  }

  return normalized;
}

export function getMcpAdapterBehavior(serverId: string | undefined, toolName: string): IMcpAdapterBehavior {
  const normalizedServerId = (serverId || '').toLowerCase();
  const normalizedToolName = toolName.toLowerCase();

  if (normalizedServerId.includes('me') || normalizedToolName.includes('userdetails') || normalizedToolName.includes('managerdetails')) {
    return {
      family: 'profile',
      aliasHints: ['userEmail -> userIdentifier', 'personEmail -> userIdentifier', 'displayName -> userIdentifier'],
      targetResolution: ['selected person from HIE', 'explicit person target', 'fail before falling back to me'],
      resultShaping: ['profile cards with richer select projection']
    };
  }
  if (normalizedServerId.includes('sharepoint') || normalizedServerId.includes('odsp')) {
    return {
      family: 'sharepoint-onedrive',
      aliasHints: ['sitePath/siteUrl -> hostname + serverRelativePath', 'driveId -> documentLibraryId', 'itemId/fileId/folderId -> fileOrFolderId', 'listItemId -> itemId'],
      targetResolution: ['explicit site/list/library target', 'selected HIE site/list/file target', 'current page site only as fallback'],
      resultShaping: ['wrapped Graph payload unwrapping', 'friendly list/table columns', 'clickable URLs']
    };
  }
  if (normalizedServerId.includes('mail') || normalizedServerId.includes('calendar')) {
    return {
      family: 'mail-calendar',
      aliasHints: ['message/query/prompt -> message', 'email tags -> array fields'],
      targetResolution: ['explicit mail/calendar item target', 'selected HIE item target'],
      resultShaping: ['humanized compose and item summaries']
    };
  }
  if (normalizedServerId.includes('teams')) {
    return {
      family: 'teams',
      aliasHints: ['message -> content', 'team/channel names -> ids when possible'],
      targetResolution: ['explicit Teams destination', 'selected HIE team/channel target'],
      resultShaping: ['share compose fields and clickable destination links']
    };
  }
  if (normalizedToolName.includes('copilot')) {
    return {
      family: 'copilot',
      aliasHints: ['query/prompt/text -> message'],
      targetResolution: ['explicit content target', 'selected HIE content target'],
      resultShaping: ['unwrap nested response payloads before rendering']
    };
  }
  return {
    family: 'generic',
    aliasHints: ['schema-driven alias normalization only'],
    targetResolution: ['explicit target', 'HIE target context', 'current page fallback when safe'],
    resultShaping: ['generic wrapper unwrapping and block inference']
  };
}

export function resolveEffectiveMcpTargetContext(
  options: IResolveMcpTargetContextOptions
): { targetContext?: IMcpTargetContext; targetSource: McpTargetSource } {
  const explicitTargetContext = mergeMcpTargetContexts(
    options.explicitTargetContext,
    deriveMcpTargetContextFromUnknown(options.explicitTargetContext, 'explicit-user')
  );
  if (explicitTargetContext) {
    return {
      targetContext: { ...explicitTargetContext, source: 'explicit-user' },
      targetSource: 'explicit-user'
    };
  }

  const hieTargetContext = deriveMcpTargetContextFromHie(
    options.sourceContext,
    options.taskContext,
    options.artifacts
  );
  if (hieTargetContext) {
    return {
      targetContext: { ...hieTargetContext, source: hieTargetContext.source || 'hie-selection' },
      targetSource: 'hie-selection'
    };
  }

  const currentPageTarget = deriveMcpTargetContextFromUnknown(
    options.currentSiteUrl || getBrowserSiteContext().siteUrl,
    'current-page'
  );
  if (currentPageTarget) {
    return {
      targetContext: { ...currentPageTarget, source: 'current-page' },
      targetSource: 'current-page'
    };
  }

  return { targetContext: undefined, targetSource: 'none' };
}

export function applySchemaArgAliases(
  toolName: string,
  args: Record<string, unknown>,
  rawArgs: Record<string, unknown>,
  schemaProps: Record<string, { type?: string }>,
  currentSiteUrl?: string,
  targetContext?: IMcpTargetContext
): Record<string, unknown> {
  const normalized = { ...args };
  const normalizedToolName = toolName.toLowerCase();
  const propNames = Object.keys(schemaProps);
  const getPropName = (target: string): string | undefined =>
    propNames.find((prop) => prop.toLowerCase() === target.toLowerCase());
  const toRecord = (value: unknown): Record<string, unknown> | undefined =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? value as Record<string, unknown>
      : undefined;

  const userIdentifierField = getPropName('userIdentifier');
  if (userIdentifierField && !pickFirstNonEmptyString(normalized[userIdentifierField])) {
    const identifier = pickFirstNonEmptyString(
      rawArgs[userIdentifierField],
      rawArgs.userEmail,
      rawArgs.email,
      rawArgs.userPrincipalName,
      rawArgs.upn,
      rawArgs.emailAddress,
      rawArgs.userId,
      rawArgs.personEmail,
      rawArgs.personIdentifier,
      rawArgs.person,
      rawArgs.displayName,
      rawArgs.name,
      targetContext?.personIdentifier,
      targetContext?.personEmail,
      targetContext?.personDisplayName
    );
    if (identifier) {
      normalized[userIdentifierField] = identifier;
    }
  }

  const hostnameField = getPropName('hostname');
  const sitePathField = getPropName('serverRelativePath');
  if (hostnameField && sitePathField) {
    const currentHostname = pickFirstNonEmptyString(normalized[hostnameField]);
    const currentPath = pickFirstNonEmptyString(normalized[sitePathField]);
    if (!currentHostname || !currentPath) {
      const browserSiteContext = getBrowserSiteContext();
      const parsed = parseSharePointSiteReference(
        pickFirstNonEmptyString(
          rawArgs.siteUrl,
          rawArgs.site_url,
          rawArgs.url,
          rawArgs.webUrl,
          rawArgs.target_url,
          rawArgs.targetUrl,
          rawArgs.sitePath,
          rawArgs.site_path,
          rawArgs.path,
          targetContext?.siteUrl,
          targetContext?.listUrl,
          targetContext?.documentLibraryUrl,
          targetContext?.fileOrFolderUrl,
          currentSiteUrl,
          browserSiteContext.siteUrl
        ),
        currentSiteUrl || targetContext?.siteUrl
      );
      if (!currentHostname && parsed.hostname) {
        normalized[hostnameField] = parsed.hostname;
      }
      if (!currentPath && parsed.serverRelativePath) {
        normalized[sitePathField] = parsed.serverRelativePath;
      }
    }
  }

  if (normalizedToolName === 'createlist') {
    const displayNameField = getPropName('displayName');
    if (displayNameField && !pickFirstNonEmptyString(normalized[displayNameField])) {
      const legacyListName = pickFirstNonEmptyString(
        rawArgs.displayName,
        rawArgs.display_name,
        rawArgs.listName,
        rawArgs.list_name,
        rawArgs.listTitle,
        rawArgs.title,
        rawArgs.name
      );
      if (legacyListName) {
        normalized[displayNameField] = legacyListName;
      }
    }

    const listField = getPropName('list');
    if (listField) {
      const existingList = toRecord(normalized[listField]) || {};
      const template = normalizeSharePointListTemplate(
        pickFirstNonEmptyString(existingList.template, rawArgs.template)
      );
      const description = pickFirstNonEmptyString(existingList.description, rawArgs.description);
      if (template || description) {
        normalized[listField] = {
          ...existingList,
          ...(template ? { template } : {}),
          ...(description ? { description } : {})
        };
      }
    }
  }

  if (normalizedToolName === 'createlistcolumn') {
    const displayNameField = getPropName('displayName');
    const resolvedDisplayName = pickFirstNonEmptyString(
      normalized[displayNameField || ''],
      rawArgs.columnDisplayName,
      rawArgs.column_display_name,
      rawArgs.displayName,
      rawArgs.display_name,
      rawArgs.columnName,
      rawArgs.column_name,
      rawArgs.name
    );
    if (displayNameField && resolvedDisplayName && !pickFirstNonEmptyString(normalized[displayNameField])) {
      normalized[displayNameField] = resolvedDisplayName;
    }

    const nameField = getPropName('name');
    if (nameField && !pickFirstNonEmptyString(normalized[nameField])) {
      const requestedName = pickFirstNonEmptyString(
        rawArgs.columnName,
        rawArgs.column_name,
        rawArgs.name,
        resolvedDisplayName
      );
      const sanitizedName = sanitizeSharePointColumnInternalName(requestedName);
      if (sanitizedName) {
        normalized[nameField] = sanitizedName;
      }
    }

    const requestedType = pickFirstNonEmptyString(
      rawArgs.columnType,
      rawArgs.column_type,
      rawArgs.type
    );
    const choiceValues = parseSharePointChoiceValues(
      pickFirstNonEmptyString(
        rawArgs.choiceValues,
        rawArgs.choice_values
      ) || rawArgs.choices
    );
    const normalizedType = normalizeSharePointColumnType(requestedType);
    if (normalizedType) {
      propNames.forEach((propertyName) => {
        if ([
          'text',
          'number',
          'choice',
          'boolean',
          'dateTime',
          'personOrGroup',
          'lookup',
          'hyperlinkOrPicture'
        ].includes(propertyName)) {
          delete normalized[propertyName];
        }
      });

      const typeField = getPropName(normalizedType.key);
      if (typeField) {
        normalized[typeField] = normalizedType.key === 'choice' && choiceValues
          ? {
            ...normalizedType.value,
            choices: choiceValues
          }
          : normalizedType.value;
      }
    }
  }

  propNames.forEach((propertyName) => {
    if (pickFirstNonEmptyString(normalized[propertyName])) {
      return;
    }

    const lowered = propertyName.toLowerCase();
    if (lowered === 'siteid' && targetContext?.siteId) {
      normalized[propertyName] = targetContext.siteId;
      return;
    }
    if ((lowered.includes('documentlibraryid') || lowered === 'driveid' || lowered.includes('driveid')) && targetContext?.documentLibraryId) {
      normalized[propertyName] = targetContext.documentLibraryId;
      return;
    }
    if ((lowered === 'documentlibraryurl' || lowered === 'driveurl') && targetContext?.documentLibraryUrl) {
      normalized[propertyName] = targetContext.documentLibraryUrl;
      return;
    }
    if ((lowered.includes('listid') || lowered === 'listid') && targetContext?.listId) {
      normalized[propertyName] = targetContext.listId;
      return;
    }
    if (lowered === 'itemid' && normalizedToolName.includes('listitem') && targetContext?.listItemId) {
      normalized[propertyName] = targetContext.listItemId;
      return;
    }
    if (lowered === 'listurl' && targetContext?.listUrl) {
      normalized[propertyName] = targetContext.listUrl;
      return;
    }
    if (
      (lowered === 'messageid' || (lowered === 'id' && normalizedToolName.includes('message')))
      && targetContext?.mailItemId
    ) {
      normalized[propertyName] = targetContext.mailItemId;
      return;
    }
    if (lowered === 'eventid' && targetContext?.calendarItemId) {
      normalized[propertyName] = targetContext.calendarItemId;
      return;
    }
    if (
      lowered.includes('fileorfolderid')
      || lowered === 'fileid'
      || lowered === 'itemid'
      || lowered === 'folderid'
      || lowered.includes('driveitemid')
    ) {
      if (targetContext?.fileOrFolderId) {
        normalized[propertyName] = targetContext.fileOrFolderId;
        return;
      }
    }
    if (lowered === 'parentfolderid' && targetContext?.fileOrFolderId) {
      normalized[propertyName] = targetContext.fileOrFolderId;
      return;
    }
    if ((lowered === 'fileorfolderurl' || lowered === 'fileurl' || lowered === 'folderurl') && targetContext?.fileOrFolderUrl) {
      normalized[propertyName] = targetContext.fileOrFolderUrl;
      return;
    }
    if (lowered === 'teamid' && targetContext?.teamId) {
      normalized[propertyName] = targetContext.teamId;
      return;
    }
    if (lowered === 'channelid' && targetContext?.channelId) {
      normalized[propertyName] = targetContext.channelId;
      return;
    }
    if ((lowered === 'siteurl' || lowered === 'weburl' || lowered === 'url') && targetContext?.siteUrl) {
      normalized[propertyName] = targetContext.siteUrl;
    }
  });

  const selectField = getPropName('select');
  if (
    selectField
    && !pickFirstNonEmptyString(normalized[selectField])
    && ['getuserdetails', 'getmydetails'].includes(toolName.toLowerCase())
  ) {
    normalized[selectField] = 'displayName,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone,givenName,surname,id';
  }

  return normalized;
}

export function extractStructuredMcpPayload(
  content: Array<{ type: string; text?: string }>
): { payload?: unknown; unwrapPath: string[] } {
  const unwrapPath: string[] = [];
  const textParts = content
    .filter((item) => item.type === 'text' && typeof item.text === 'string' && item.text.trim())
    .map((item) => item.text as string);

  for (let i = 0; i < textParts.length; i++) {
    const parsed = tryParseJson(textParts[i]);
    if (parsed === undefined) {
      continue;
    }

    let current = parsed;
    for (let depth = 0; depth < 4; depth++) {
      if (!current || typeof current !== 'object' || Array.isArray(current)) {
        return { payload: current, unwrapPath };
      }
      const record = current as Record<string, unknown>;
      const unwrapKeys = ['response', 'result', 'payload', 'data'];
      let next: unknown | undefined;

      for (let j = 0; j < unwrapKeys.length; j++) {
        const key = unwrapKeys[j];
        const candidate = record[key];
        if (typeof candidate === 'string') {
          const parsedCandidate = tryParseJson(candidate);
          if (parsedCandidate !== undefined) {
            unwrapPath.push(key);
            next = parsedCandidate;
            break;
          }
        } else if (candidate && typeof candidate === 'object') {
          unwrapPath.push(key);
          next = candidate;
          break;
        }
      }

      if (next === undefined) {
        return { payload: current, unwrapPath };
      }
      current = next;
    }

    return { payload: current, unwrapPath };
  }

  return { unwrapPath };
}

const HIDDEN_ERROR_PATTERNS = [
  /\binvalid\s+(parameter|argument|input|request)/i,
  /\brequired\s+(parameter|field|property)\b/i,
  /\bmissing\s+(required|parameter|field|property)\b/i,
  /\bparameter\s+['"]?\w+['"]?\s+is\s+required/i,
  /\bnot\s+found\b.*\b(resource|item|file|folder|site|list|drive)\b/i,
  /\b(resource|item|file|folder|site|list|drive)\b.*\bnot\s+found\b/i,
  /\bunauthorized\b/i,
  /\bforbidden\b/i,
  /\baccess\s+denied\b/i,
  /\binternal\s+server\s+error\b/i,
  /\btool\s+execution\s+failed\b/i,
  /\bInvalidArgument\b/,
  /\bBadRequest\b/
];

export function extractHiddenMcpError(content: Array<{ type: string; text?: string }>): string | undefined {
  const fullText = content
    .filter((item) => item.type === 'text' && item.text)
    .map((item) => item.text as string)
    .join(' ');
  if (!fullText || fullText.length > 800) {
    return undefined;
  }
  for (let i = 0; i < HIDDEN_ERROR_PATTERNS.length; i++) {
    if (HIDDEN_ERROR_PATTERNS[i].test(fullText)) {
      return fullText.trim();
    }
  }
  return undefined;
}

export function createMcpExecutionTrace(
  toolName: string,
  rawArgs: Record<string, unknown>,
  options: {
    serverId?: string;
    serverName?: string;
    requiredFields?: string[];
    normalizedArgs?: Record<string, unknown>;
    resolvedArgs?: Record<string, unknown>;
    targetContext?: IMcpTargetContext;
    targetSource?: McpTargetSource;
    recoverySteps?: string[];
    unwrapPath?: string[];
    inferredBlockType?: BlockType;
    finalBlockTitle?: string;
    finalSummary?: string;
  } = {}
): IMcpExecutionTrace {
  return {
    serverId: options.serverId,
    serverName: options.serverName,
    toolName,
    requiredFields: options.requiredFields,
    rawArgs,
    normalizedArgs: options.normalizedArgs,
    resolvedArgs: options.resolvedArgs,
    targetContext: options.targetContext,
    targetSource: options.targetSource || options.targetContext?.source || 'none',
    targetSummary: describeMcpTargetContext(options.targetContext),
    recoverySteps: options.recoverySteps || [],
    unwrapPath: options.unwrapPath,
    inferredBlockType: options.inferredBlockType,
    finalBlockTitle: options.finalBlockTitle,
    finalSummary: options.finalSummary
  };
}

export async function executeCatalogMcpTool(
  options: IExecuteCatalogMcpToolOptions
): Promise<IExecuteCatalogMcpToolResult> {
  const {
    serverId,
    serverName,
    serverUrl,
    toolName,
    rawArgs,
    mcpClient,
    sessionHelpers
  } = options;
  const targetResolution = resolveEffectiveMcpTargetContext({
    explicitTargetContext: options.explicitTargetContext,
    sourceContext: options.sourceContext,
    taskContext: options.taskContext,
    artifacts: options.artifacts,
    currentSiteUrl: options.currentSiteUrl
  });
  const connectionResult = await ensureCatalogMcpConnection({
    serverId,
    serverName,
    serverUrl,
    connections: options.connections,
    mcpClient,
    sessionHelpers,
    getToken: options.getToken,
    getConnections: options.getConnections
  });
  const recoverySteps = [...connectionResult.recoverySteps];
  const sessionId = connectionResult.sessionId || '';
  const connection = connectionResult.connection;
  const liveConnection = connectionResult.liveConnection;
  const fallbackConnection = connectionResult.fallbackConnection;

  const buildTrace = (overrides: Partial<IMcpExecutionTrace> & {
    requiredFields?: string[];
    normalizedArgs?: Record<string, unknown>;
    resolvedArgs?: Record<string, unknown>;
  } = {}): IMcpExecutionTrace => createMcpExecutionTrace(
    overrides.toolName || toolName,
    rawArgs,
    {
      serverId,
      serverName,
      requiredFields: overrides.requiredFields,
      normalizedArgs: overrides.normalizedArgs,
      resolvedArgs: overrides.resolvedArgs,
      targetContext: targetResolution.targetContext,
      targetSource: targetResolution.targetSource,
      recoverySteps: overrides.recoverySteps || recoverySteps,
      unwrapPath: overrides.unwrapPath,
      inferredBlockType: overrides.inferredBlockType,
      finalBlockTitle: overrides.finalBlockTitle,
      finalSummary: overrides.finalSummary
    }
  );

  if (!connectionResult.success || !connection) {
    const error = connectionResult.error || `Connected to ${serverName} but no session metadata was available for MCP execution.`;
    return {
      success: false,
      serverId,
      serverName,
      serverUrl,
      sessionId,
      realToolName: toolName,
      requiredFields: [],
      schemaProps: {},
      normalizedArgs: {},
      resolvedArgs: { ...rawArgs },
      targetContext: targetResolution.targetContext,
      targetSource: targetResolution.targetSource,
      recoverySteps,
      error,
      trace: buildTrace({ resolvedArgs: { ...rawArgs }, finalSummary: error })
    };
  }

  const liveMatch = liveConnection?.tools.find((tool) => tool.name.toLowerCase() === toolName.toLowerCase());
  const fallbackMatch = fallbackConnection?.tools.find((tool) => tool.name.toLowerCase() === toolName.toLowerCase());
  const match = liveMatch || fallbackMatch;
  if (!liveMatch && fallbackMatch) {
    recoverySteps.push(`used catalog tool schema for ${toolName}`);
    logService.info('mcp', `Using catalog tool schema fallback for ${toolName} on ${serverName}`);
  }
  if (!match) {
    const available = connection.tools.map((tool) => tool.name).join(', ');
    const error = `Tool "${toolName}" not found on ${serverName}. Available tools: ${available}`;
    return {
      success: false,
      serverId,
      serverName,
      serverUrl,
      sessionId,
      realToolName: toolName,
      connection,
      requiredFields: [],
      schemaProps: {},
      normalizedArgs: {},
      resolvedArgs: { ...rawArgs },
      targetContext: targetResolution.targetContext,
      targetSource: targetResolution.targetSource,
      recoverySteps,
      error,
      trace: buildTrace({ resolvedArgs: { ...rawArgs }, finalSummary: error })
    };
  }

  const realToolName = match.name;
  const validation = validateMcpArgs(rawArgs, match.inputSchema as Record<string, unknown>);
  const schemaProps = (match.inputSchema as { properties?: Record<string, { type?: string; description?: string }>; required?: string[] }).properties || {};
  const requiredFields = (match.inputSchema as { required?: string[] }).required || [];
  const resolvedArgs = applySchemaArgAliases(
    realToolName,
    { ...validation.matched },
    rawArgs,
    schemaProps,
    options.currentSiteUrl,
    targetResolution.targetContext
  );
  let missingFields = requiredFields.filter((field) =>
    resolvedArgs[field] === undefined
    || resolvedArgs[field] === null
    || (typeof resolvedArgs[field] === 'string' && resolvedArgs[field].trim() === '')
  );

  if (validation.typeErrors.length > 0) {
    logService.warning('mcp', `Type errors for ${realToolName}: ${validation.typeErrors.join(', ')}`);
  }

  logService.info('mcp', `Validation: matched=${Object.keys(validation.matched).length}, missing=${missingFields.length}, extra=${Object.keys(validation.extra).length}`);
  if (missingFields.length > 0) {
    logService.debug('mcp', `${realToolName} schema: required=[${requiredFields.join(',')}] props=[${Object.keys(schemaProps).join(',')}] llmArgs=[${Object.keys(rawArgs).join(',')}]`);
  }

  const resolutionExtra: Record<string, unknown> = { ...validation.extra };
  const currentSiteReference = pickFirstNonEmptyString(
    targetResolution.targetContext?.siteUrl,
    targetResolution.targetContext?.listUrl,
    targetResolution.targetContext?.documentLibraryUrl,
    targetResolution.targetContext?.fileOrFolderUrl,
    options.currentSiteUrl,
    getBrowserSiteContext().siteUrl
  );
  if (
    missingFields.some((field) => field.toLowerCase() === 'siteid')
    && Object.keys(resolutionExtra).length === 0
    && currentSiteReference
  ) {
    resolutionExtra.currentSiteUrl = currentSiteReference;
  }

  if (missingFields.length > 0 && Object.keys(resolutionExtra).length > 0) {
    const resolved = await resolveMissingMcpParameters(
      missingFields,
      resolutionExtra,
      connection.tools.map((tool) => ({
        name: tool.name,
        inputSchema: tool.inputSchema as Record<string, unknown>
      })),
      mcpClient,
      sessionId
    );
    missingFields.forEach((missingField) => {
      const resolvedValue = resolved[missingField.toLowerCase()];
      if (resolvedValue !== undefined) {
        resolvedArgs[missingField] = resolvedValue;
        recoverySteps.push(`resolved ${missingField}`);
        logService.info('mcp', `Auto-resolved ${missingField}: ${String(resolvedValue).substring(0, 80)}`);
      }
    });
  }

  if (missingFields.some((field) => field.toLowerCase() === 'listid')) {
    const listResolution = await resolveSharePointListId(
      {
        rawArgs,
        resolvedArgs,
        schemaProps,
        targetContext: targetResolution.targetContext
      },
      connection.tools.map((tool) => ({
        name: tool.name,
        inputSchema: tool.inputSchema as Record<string, unknown>
      })),
      mcpClient,
      sessionId
    );

    missingFields.forEach((missingField) => {
      const resolvedValue = listResolution[missingField.toLowerCase()];
      if (resolvedValue !== undefined) {
        resolvedArgs[missingField] = resolvedValue;
        recoverySteps.push(`resolved ${missingField}`);
        logService.info('mcp', `Auto-resolved ${missingField}: ${String(resolvedValue).substring(0, 80)}`);
      }
    });
  }

  requiredFields.forEach((field) => {
    if (typeof resolvedArgs[field] === 'string' && resolvedArgs[field].trim() === '') {
      resolvedArgs[field] = '*';
      recoverySteps.push(`filled empty ${field}=*`);
      logService.info('mcp', `Filled empty required param "${field}" with "*"`);
    }
  });

  const userIdentifierField = requiredFields.find((field) => field.toLowerCase() === 'useridentifier');
  if (userIdentifierField) {
    const userIdentifierValue = resolvedArgs[userIdentifierField];
    if (
      (userIdentifierValue === undefined
        || userIdentifierValue === null
        || (typeof userIdentifierValue === 'string' && userIdentifierValue.trim() === ''))
      && !pickFirstNonEmptyString(
        targetResolution.targetContext?.personIdentifier,
        targetResolution.targetContext?.personEmail,
        targetResolution.targetContext?.personDisplayName
      )
    ) {
      resolvedArgs[userIdentifierField] = 'me';
      recoverySteps.push(`defaulted ${userIdentifierField}=me`);
      logService.info('mcp', `Auto-filled "${userIdentifierField}" with "me" for ${realToolName}`);
    }
  }

  const normalizedExecutionArgs = normalizeCopilotChatArgs(
    realToolName,
    resolvedArgs,
    rawArgs,
    !!options.webGroundingEnabled
  );

  const enrichResolvedArgs = options.enrichResolvedArgs;
  const finalResolvedArgs = enrichResolvedArgs
    ? await (async (): Promise<Record<string, unknown>> => {
      const enrichment = await enrichResolvedArgs({
        args: normalizedExecutionArgs,
        toolName: realToolName,
        serverId,
        serverName,
        sessionId,
        mcpClient,
        targetContext: targetResolution.targetContext,
        targetSource: targetResolution.targetSource
      });
      if (enrichment.recoverySteps && enrichment.recoverySteps.length > 0) {
        recoverySteps.push(...enrichment.recoverySteps);
      }
      return enrichment.args;
    })()
    : normalizedExecutionArgs;

  missingFields = requiredFields.filter((field) =>
    finalResolvedArgs[field] === undefined
    || finalResolvedArgs[field] === null
    || (typeof finalResolvedArgs[field] === 'string' && finalResolvedArgs[field].trim() === '')
  );

  if (missingFields.length > 0) {
    const missingInfo = missingFields.map((field) => {
      const prop = schemaProps[field];
      return `${field} (${prop?.type || 'unknown'}${prop?.description ? `: ${prop.description}` : ''})`;
    });
    const error = `Missing required parameters: ${missingInfo.join(', ')}`;
    return {
      success: false,
      serverId,
      serverName,
      serverUrl,
      sessionId,
      realToolName,
      connection,
      requiredFields,
      schemaProps,
      normalizedArgs: validation.matched,
      resolvedArgs: finalResolvedArgs,
      targetContext: targetResolution.targetContext,
      targetSource: targetResolution.targetSource,
      recoverySteps,
      error,
      trace: buildTrace({
        toolName: realToolName,
        requiredFields,
        normalizedArgs: validation.matched,
        resolvedArgs: finalResolvedArgs,
        finalSummary: error
      })
    };
  }

  const mcpResult = await withMcpRetry(
    mcpClient,
    sessionId,
    realToolName,
    finalResolvedArgs,
    serverUrl,
    serverName
  );

  if (mcpResult.success) {
    const hiddenError = extractHiddenMcpError(mcpResult.content);
    if (hiddenError) {
      logService.warning('mcp', `Tool ${realToolName}: success response contains error: ${hiddenError.substring(0, 200)}`);
      return {
        success: false,
        serverId,
        serverName,
        serverUrl,
        sessionId,
        realToolName,
        connection,
        requiredFields,
        schemaProps,
        normalizedArgs: validation.matched,
        resolvedArgs: finalResolvedArgs,
        targetContext: targetResolution.targetContext,
        targetSource: targetResolution.targetSource,
        recoverySteps,
        mcpResult: { ...mcpResult, success: false, error: hiddenError },
        error: hiddenError,
        trace: buildTrace({
          toolName: realToolName,
          requiredFields,
          normalizedArgs: validation.matched,
          resolvedArgs: finalResolvedArgs,
          finalSummary: hiddenError
        })
      };
    }
  }

  if (!mcpResult.success) {
    const error = mcpResult.error || 'Tool execution failed';
    return {
      success: false,
      serverId,
      serverName,
      serverUrl,
      sessionId,
      realToolName,
      connection,
      requiredFields,
      schemaProps,
      normalizedArgs: validation.matched,
      resolvedArgs: finalResolvedArgs,
      targetContext: targetResolution.targetContext,
      targetSource: targetResolution.targetSource,
      recoverySteps,
      mcpResult,
      error,
      trace: buildTrace({
        toolName: realToolName,
        requiredFields,
        normalizedArgs: validation.matched,
        resolvedArgs: finalResolvedArgs,
        finalSummary: error
      })
    };
  }

  return {
    success: true,
    serverId,
    serverName,
    serverUrl,
    sessionId,
    realToolName,
    connection,
    requiredFields,
    schemaProps,
    normalizedArgs: validation.matched,
    resolvedArgs: finalResolvedArgs,
    targetContext: targetResolution.targetContext,
    targetSource: targetResolution.targetSource,
    recoverySteps,
    mcpResult,
    trace: buildTrace({
      toolName: realToolName,
      requiredFields,
      normalizedArgs: validation.matched,
      resolvedArgs: finalResolvedArgs
    })
  };
}

export async function ensureCatalogMcpConnection(
  options: IEnsureCatalogMcpConnectionOptions
): Promise<IEnsureCatalogMcpConnectionResult> {
  const currentConnections = options.getConnections ? options.getConnections() : options.connections;
  const recoverySteps: string[] = [];

  try {
    const existingSessionId = options.sessionHelpers.findExistingSession(currentConnections, options.serverUrl);
    const sessionId = existingSessionId || await options.sessionHelpers.connectToM365Server(
      options.mcpClient,
      options.serverUrl,
      options.serverName,
      options.getToken
    );
    const postConnectConnections = options.getConnections ? options.getConnections() : options.connections;
    const liveConnection = postConnectConnections.find((entry) => entry.sessionId === sessionId);
    const fallbackConnection = buildCatalogConnectionFallback(options.serverId, options.serverUrl, options.serverName, sessionId);
    const connection = liveConnection || fallbackConnection;

    if (!liveConnection && fallbackConnection) {
      recoverySteps.push('used catalog session metadata fallback');
      logService.info('mcp', `Using catalog metadata fallback for ${options.serverName} session ${sessionId}`);
    }

    if (!connection) {
      return {
        success: false,
        sessionId,
        liveConnection,
        fallbackConnection,
        recoverySteps,
        error: `Connected to ${options.serverName} but no session metadata was available for MCP execution.`
      };
    }

    return {
      success: true,
      sessionId,
      connection,
      liveConnection,
      fallbackConnection,
      recoverySteps
    };
  } catch (error) {
    return {
      success: false,
      recoverySteps,
      error: error instanceof Error ? error.message : `Failed to connect to ${options.serverName}.`
    };
  }
}
