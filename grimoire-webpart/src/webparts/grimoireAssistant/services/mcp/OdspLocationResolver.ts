import {
  getCatalogEntry,
  resolveServerUrl
} from '../../models/McpServerCatalog';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IFunctionCallStore } from '../tools/ToolRuntimeContracts';
import { findExistingSession, connectToM365Server } from '../tools/ToolRuntimeSharedHelpers';
import { logService } from '../logging/LogService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { McpClientService } from './McpClientService';
import {
  executeCatalogMcpTool,
  extractStructuredMcpPayload
} from './McpExecutionAdapter';
import type { IMcpTargetContext } from './McpTargetContext';

export interface IResolvedPersonalOneDriveLocation {
  siteUrl: string;
  siteId: string;
  documentLibraryId: string;
}

interface IOdspResolutionOptions {
  siteUrl?: string;
  currentSiteUrl?: string;
}

export function pickFirstNonEmptyString(...values: unknown[]): string | undefined {
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (typeof value !== 'string') {
      continue;
    }
    const trimmed = value.trim();
    if (trimmed) {
      return trimmed;
    }
  }
  return undefined;
}

function getBrowserCurrentSiteUrl(): string | undefined {
  if (typeof window === 'undefined' || !window.location) {
    return undefined;
  }

  const { hostname, origin, pathname } = window.location;
  if (!hostname || hostname === 'localhost' || hostname === '127.0.0.1') {
    return undefined;
  }

  const siteRootMatch = pathname.match(/^\/(?:sites|teams|personal)\/[^/]+/i);
  return siteRootMatch?.[0] ? `${origin}${siteRootMatch[0]}` : undefined;
}

export function parseSharePointSiteReference(reference: string): { hostname?: string; serverRelativePath?: string } {
  const candidate = reference.trim();
  if (!candidate) {
    return {};
  }

  const normalizePath = (rawPath: string): string | undefined => {
    const withoutQuery = rawPath.trim().split('?')[0].split('#')[0];
    const normalized = withoutQuery.replace(/^\/+/, '').replace(/\/+$/, '');
    return normalized || undefined;
  };

  try {
    const url = new URL(candidate);
    return {
      hostname: url.hostname || undefined,
      serverRelativePath: normalizePath(decodeURIComponent(url.pathname))
    };
  } catch {
    return {
      serverRelativePath: normalizePath(candidate)
    };
  }
}

export function extractSiteId(payload: unknown): string | undefined {
  if (!payload || typeof payload !== 'object') {
    return undefined;
  }
  const record = payload as Record<string, unknown>;
  if (typeof record.id === 'string') return record.id;
  if (typeof record.siteId === 'string') return record.siteId;
  return undefined;
}

async function executeCatalogHelperTool(
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  serverId: string,
  toolName: string,
  rawArgs: Record<string, unknown>,
  targetContext?: IMcpTargetContext
): Promise<unknown> {
  const server = getCatalogEntry(serverId);
  if (!server || !store.mcpEnvironmentId) {
    return undefined;
  }

  const currentStore = useGrimoireStore.getState();
  const execution = await executeCatalogMcpTool({
    serverId,
    serverName: server.name,
    serverUrl: resolveServerUrl(serverId, store.mcpEnvironmentId),
    toolName,
    rawArgs,
    connections: currentStore.mcpConnections,
    getConnections: () => useGrimoireStore.getState().mcpConnections,
    mcpClient,
    sessionHelpers: {
      findExistingSession,
      connectToM365Server
    },
    getToken: currentStore.getToken,
    explicitTargetContext: targetContext,
    taskContext: hybridInteractionEngine.getCurrentTaskContext(),
    artifacts: hybridInteractionEngine.getCurrentArtifacts(),
    currentSiteUrl: currentStore.userContext?.currentSiteUrl
  });

  if (!execution.success || !execution.mcpResult) {
    logService.warning('mcp', `Helper MCP execution failed for ${toolName}: ${execution.error || 'Unknown error'}`);
    return undefined;
  }

  return extractStructuredMcpPayload(execution.mcpResult.content).payload;
}

export function extractDocumentLibraryId(payload: unknown): string | undefined {
  if (!payload || typeof payload !== 'object') {
    return undefined;
  }
  const record = payload as Record<string, unknown>;
  if (typeof record.id === 'string') return record.id;
  if (typeof record.driveId === 'string') return record.driveId;
  if (typeof record.documentLibraryId === 'string') return record.documentLibraryId;
  const value = Array.isArray(record.value) ? record.value[0] : undefined;
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    return extractDocumentLibraryId(value);
  }
  return undefined;
}

function normalizeLookupValue(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function extractDocumentLibraryEntries(payload: unknown): Array<Record<string, unknown>> {
  if (Array.isArray(payload)) {
    return payload.filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object' && !Array.isArray(entry));
  }

  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return [];
  }

  const record = payload as Record<string, unknown>;
  const candidates = [record.value, record.items, record.results];
  for (let i = 0; i < candidates.length; i++) {
    const candidateArray = candidates[i];
    if (!Array.isArray(candidateArray)) {
      continue;
    }

    return candidateArray.filter((entry: unknown): entry is Record<string, unknown> => (
      !!entry && typeof entry === 'object' && !Array.isArray(entry)
    ));
  }

  return [record];
}

export async function resolveNamedDocumentLibraryId(
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  targetContext?: IMcpTargetContext,
  options: IOdspResolutionOptions = {}
): Promise<{ documentLibraryId?: string; checkedByName: boolean }> {
  const requestedLibraryName = pickFirstNonEmptyString(targetContext?.documentLibraryName);
  if (!requestedLibraryName) {
    return { checkedByName: false };
  }

  if (targetContext?.documentLibraryId) {
    return {
      documentLibraryId: targetContext.documentLibraryId,
      checkedByName: true
    };
  }

  const effectiveTargetContext = options.siteUrl
    ? { ...targetContext, siteUrl: options.siteUrl }
    : targetContext;
  const siteId = effectiveTargetContext?.siteId || await resolveSiteIdForTargetContext(
    store,
    mcpClient,
    effectiveTargetContext,
    options
  );
  const payload = await executeCatalogHelperTool(
    store,
    mcpClient,
    'mcp_ODSPRemoteServer',
    'listDocumentLibrariesInSite',
    siteId ? { siteId } : {},
    effectiveTargetContext
  );

  const normalizedTargetName = normalizeLookupValue(requestedLibraryName);
  const match = extractDocumentLibraryEntries(payload).find((entry) => {
    const candidateName = pickFirstNonEmptyString(
      entry.displayName,
      entry.name,
      entry.driveName,
      entry.documentLibraryName,
      entry.title
    );
    return candidateName ? normalizeLookupValue(candidateName) === normalizedTargetName : false;
  });

  return {
    documentLibraryId: extractDocumentLibraryId(match),
    checkedByName: true
  };
}

export async function resolveSiteIdForTargetContext(
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  targetContext?: IMcpTargetContext,
  options: IOdspResolutionOptions = {}
): Promise<string | undefined> {
  const currentSiteUrl = pickFirstNonEmptyString(
    options.siteUrl,
    targetContext?.siteUrl,
    targetContext?.listUrl,
    targetContext?.documentLibraryUrl,
    targetContext?.fileOrFolderUrl,
    options.currentSiteUrl,
    store.userContext?.currentSiteUrl,
    getBrowserCurrentSiteUrl()
  );
  if (!currentSiteUrl) {
    return undefined;
  }

  const siteRef = parseSharePointSiteReference(currentSiteUrl);
  if (!siteRef.hostname || !siteRef.serverRelativePath) {
    return undefined;
  }

  const effectiveTargetContext = options.siteUrl
    ? { ...targetContext, siteUrl: options.siteUrl }
    : targetContext;
  const payload = await executeCatalogHelperTool(
    store,
    mcpClient,
    'mcp_SharePointListsTools',
    'getSiteByPath',
    {
      hostname: siteRef.hostname,
      serverRelativePath: siteRef.serverRelativePath
    },
    effectiveTargetContext
  );
  return extractSiteId(payload);
}

export async function resolveDefaultDocumentLibraryId(
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  targetContext?: IMcpTargetContext,
  options: IOdspResolutionOptions = {}
): Promise<string | undefined> {
  if (targetContext?.documentLibraryId) {
    return targetContext.documentLibraryId;
  }

  const effectiveTargetContext = options.siteUrl
    ? { ...targetContext, siteUrl: options.siteUrl }
    : targetContext;
  const namedResolution = await resolveNamedDocumentLibraryId(
    store,
    mcpClient,
    effectiveTargetContext,
    options
  );
  if (namedResolution.documentLibraryId) {
    return namedResolution.documentLibraryId;
  }

  if (namedResolution.checkedByName) {
    const requestedLibraryName = pickFirstNonEmptyString(effectiveTargetContext?.documentLibraryName) || 'the requested library';
    throw new Error(`Document library "${requestedLibraryName}" was not found in the selected site.`);
  }

  const siteId = effectiveTargetContext?.siteId || await resolveSiteIdForTargetContext(
    store,
    mcpClient,
    effectiveTargetContext,
    options
  );
  const payload = await executeCatalogHelperTool(
    store,
    mcpClient,
    'mcp_ODSPRemoteServer',
    'getDefaultDocumentLibraryInSite',
    siteId ? { siteId } : {},
    effectiveTargetContext
  );
  return extractDocumentLibraryId(payload);
}

function normalizePersonalSiteUrl(siteUrl: string): string | undefined {
  try {
    const url = new URL(siteUrl);
    const match = decodeURIComponent(url.pathname).match(/^\/personal\/[^/]+/i);
    return match?.[0] ? `${url.origin}${match[0]}` : undefined;
  } catch {
    return undefined;
  }
}

export function isPersonalOneDriveSiteUrl(siteUrl: string): boolean {
  return !!normalizePersonalSiteUrl(siteUrl);
}

export function normalizeM365Identity(identity: string | undefined): string | undefined {
  const rawValue = pickFirstNonEmptyString(identity);
  if (!rawValue) {
    return undefined;
  }

  let candidate = rawValue.trim();
  if (candidate.includes('|')) {
    const tail = candidate.split('|').pop();
    if (tail) {
      candidate = tail.trim();
    }
  }

  if (candidate.toLowerCase().startsWith('smtp:')) {
    candidate = candidate.substring(5).trim();
  }

  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(candidate) ? candidate : undefined;
}

export function derivePersonalOneDriveSiteUrl(options: {
  siteUrl?: string;
  currentSiteUrl?: string;
  loginName?: string;
  email?: string;
}): string | undefined {
  const explicitPersonalSiteUrl = pickFirstNonEmptyString(options.siteUrl, options.currentSiteUrl);
  if (explicitPersonalSiteUrl && isPersonalOneDriveSiteUrl(explicitPersonalSiteUrl)) {
    return normalizePersonalSiteUrl(explicitPersonalSiteUrl);
  }

  const hostSource = pickFirstNonEmptyString(options.siteUrl, options.currentSiteUrl);
  if (!hostSource) {
    return undefined;
  }

  let parsedUrl: URL;
  try {
    parsedUrl = new URL(hostSource);
  } catch {
    return undefined;
  }

  const normalizedIdentity = normalizeM365Identity(options.loginName)
    || normalizeM365Identity(options.email);
  if (!normalizedIdentity) {
    return undefined;
  }

  const loweredHost = parsedUrl.hostname.toLowerCase();
  let personalHost: string | undefined;
  if (loweredHost.endsWith('-my.sharepoint.com')) {
    personalHost = loweredHost;
  } else if (loweredHost.endsWith('.sharepoint.com')) {
    personalHost = loweredHost.replace(/\.sharepoint\.com$/i, '-my.sharepoint.com');
  }
  if (!personalHost) {
    return undefined;
  }

  const personalSegment = normalizedIdentity
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (!personalSegment) {
    return undefined;
  }

  return `${parsedUrl.protocol}//${personalHost}/personal/${personalSegment}`;
}

export async function resolvePersonalOneDriveRootLocation(
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  targetContext?: IMcpTargetContext,
  options: IOdspResolutionOptions = {}
): Promise<IResolvedPersonalOneDriveLocation> {
  const siteUrl = derivePersonalOneDriveSiteUrl({
    siteUrl: targetContext?.siteUrl,
    currentSiteUrl: options.currentSiteUrl || store.userContext?.currentSiteUrl,
    loginName: store.userContext?.loginName,
    email: store.userContext?.email
  });
  if (!siteUrl) {
    throw new Error('I could not derive your personal OneDrive site URL for MCP browsing.');
  }

  const personalTargetContext: IMcpTargetContext = {
    ...targetContext,
    siteUrl
  };
  const siteId = await resolveSiteIdForTargetContext(
    store,
    mcpClient,
    personalTargetContext,
    {
      siteUrl,
      currentSiteUrl: options.currentSiteUrl
    }
  );
  if (!siteId) {
    throw new Error('I could not resolve your personal OneDrive site through MCP.');
  }

  const documentLibraryId = await resolveDefaultDocumentLibraryId(
    store,
    mcpClient,
    {
      ...personalTargetContext,
      siteId
    },
    {
      siteUrl,
      currentSiteUrl: options.currentSiteUrl
    }
  );
  if (!documentLibraryId) {
    throw new Error('I could not resolve your personal OneDrive document library through MCP.');
  }

  return {
    siteUrl,
    siteId,
    documentLibraryId
  };
}
