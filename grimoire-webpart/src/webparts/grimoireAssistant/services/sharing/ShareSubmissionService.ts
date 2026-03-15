import type { ISelectionListData, IFormData, ISelectionItem } from '../../models/IBlock';
import type { IMcpContent } from '../../models/IMcpTypes';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IFunctionCallStore } from '../tools/ToolRuntimeContracts';
import type { IFormSubmissionResult } from '../forms/FormSubmissionService';
import { getCatalogEntry, resolveServerUrl } from '../../models/McpServerCatalog';
import { McpClientService } from '../mcp/McpClientService';
import { connectToM365Server, findExistingSession } from '../tools/ToolRuntimeSharedHelpers';
import { extractMcpTextParts } from '../mcp/mcpUtils';
import { logService } from '../logging/LogService';
import { mapMcpResultToBlock } from '../mcp/McpResultMapper';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { createCorrelationId } from '../hie/HAEContracts';
import type { IHieSourceContext } from '../hie/HIETypes';
import { executeCatalogMcpTool, type IMcpExecutionTrace } from '../mcp/McpExecutionAdapter';
import { describeMcpTargetContext, mergeMcpTargetContexts, type IMcpTargetContext } from '../mcp/McpTargetContext';

interface IUserIdentity {
  id: string;
  mail?: string;
  userPrincipalName?: string;
}

interface ITeamsChannelDestinationValue {
  teamId: string;
  channelId: string;
  teamName?: string;
  channelName?: string;
}

export interface INamedSelectionMatch<TItem extends ISelectionItem = ISelectionItem> {
  item?: TItem;
  error?: string;
}

const INTERNAL_SHARE_SERVER_ID = 'internal-share';
const SHARE_CHAT_PRESET = 'share-teams-chat';
const SHARE_CHANNEL_PRESET = 'share-teams-channel';
const URL_PATTERN = /https?:\/\/[^\s<>"')\]]+/gi;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function tryParseJson(text: string): unknown {
  const trimmed = text.trim();
  if (!trimmed) return undefined;

  try {
    return JSON.parse(trimmed);
  } catch {
    const firstBrace = trimmed.indexOf('{');
    const lastBrace = trimmed.lastIndexOf('}');
    if (firstBrace !== -1 && lastBrace > firstBrace) {
      try {
        return JSON.parse(trimmed.slice(firstBrace, lastBrace + 1));
      } catch {
        return undefined;
      }
    }

    const firstBracket = trimmed.indexOf('[');
    const lastBracket = trimmed.lastIndexOf(']');
    if (firstBracket !== -1 && lastBracket > firstBracket) {
      try {
        return JSON.parse(trimmed.slice(firstBracket, lastBracket + 1));
      } catch {
        return undefined;
      }
    }

    return undefined;
  }
}

function flattenRecords(value: unknown): Record<string, unknown>[] {
  if (Array.isArray(value)) {
    return value.filter(isRecord);
  }

  if (!isRecord(value)) {
    return [];
  }

  const nestedCollections = ['value', 'items', 'results', 'data', 'channels', 'teams'];
  for (let i = 0; i < nestedCollections.length; i++) {
    const nested = value[nestedCollections[i]];
    if (Array.isArray(nested)) {
      return nested.filter(isRecord);
    }
  }

  return [value];
}

function pickString(record: Record<string, unknown>, keys: string[]): string | undefined {
  for (let i = 0; i < keys.length; i++) {
    const value = record[keys[i]];
    if (typeof value === 'string' && value.trim()) {
      return value.trim();
    }
  }
  return undefined;
}

function joinLabels(items: ISelectionItem[]): string {
  return items.map((item) => item.label).filter(Boolean).slice(0, 8).join(', ');
}

function normalizeLabel(value: string): string {
  return value.toLowerCase().replace(/\s+/g, ' ').trim();
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function linkifyLine(line: string): string {
  const matches = Array.from(line.matchAll(URL_PATTERN));
  if (matches.length === 0) {
    return escapeHtml(line);
  }

  let cursor = 0;
  let html = '';
  matches.forEach((match) => {
    const rawUrl = match[0];
    const index = match.index ?? 0;
    html += escapeHtml(line.slice(cursor, index));
    const safeUrl = escapeHtml(rawUrl);
    html += `<a href="${safeUrl}" target="_blank" rel="noreferrer noopener">${safeUrl}</a>`;
    cursor = index + rawUrl.length;
  });
  html += escapeHtml(line.slice(cursor));
  return html;
}

export function buildTeamsMessageBody(content: string): { content: string; contentType: 'html' } {
  const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  const html = normalized
    ? normalized.split('\n').map((line) => linkifyLine(line)).join('<br/>')
    : '';

  return {
    content: html,
    contentType: 'html'
  };
}

function formatMatchError(kind: 'team' | 'channel', label: string, items: ISelectionItem[], reason: 'missing' | 'ambiguous'): string {
  const candidates = joinLabels(items);
  if (reason === 'missing') {
    return candidates
      ? `No ${kind} matched "${label}". Available ${kind}s: ${candidates}.`
      : `No ${kind}s were returned for "${label}".`;
  }

  return `More than one ${kind} matched "${label}": ${candidates}.`;
}

export function resolveNamedSelectionItem(
  items: ISelectionItem[],
  requestedLabel: string,
  kind: 'team' | 'channel'
): INamedSelectionMatch {
  const normalizedRequested = normalizeLabel(requestedLabel);
  const exactMatches = items.filter((item) => normalizeLabel(item.label) === normalizedRequested);
  if (exactMatches.length === 1) {
    return { item: exactMatches[0] };
  }
  if (exactMatches.length > 1) {
    return { error: formatMatchError(kind, requestedLabel, exactMatches, 'ambiguous') };
  }

  const partialMatches = items.filter((item) => {
    const normalizedItem = normalizeLabel(item.label);
    return normalizedItem.includes(normalizedRequested) || normalizedRequested.includes(normalizedItem);
  });
  if (partialMatches.length === 1) {
    return { item: partialMatches[0] };
  }
  if (partialMatches.length > 1) {
    return { error: formatMatchError(kind, requestedLabel, partialMatches, 'ambiguous') };
  }

  return { error: formatMatchError(kind, requestedLabel, items, 'missing') };
}

export function isInternalSharePreset(preset: string): boolean {
  return preset === SHARE_CHAT_PRESET || preset === SHARE_CHANNEL_PRESET;
}

function tryParseTeamsChannelDestination(rawValue: string): ITeamsChannelDestinationValue | undefined {
  const parsed = tryParseJson(rawValue);
  if (!isRecord(parsed)) {
    return undefined;
  }

  const teamId = pickString(parsed, ['teamId']);
  const channelId = pickString(parsed, ['channelId']);
  if (!teamId || !channelId) {
    return undefined;
  }

  return {
    teamId,
    channelId,
    teamName: pickString(parsed, ['teamName']),
    channelName: pickString(parsed, ['channelName'])
  };
}

function resolveCatalogServer(
  store: IFunctionCallStore,
  serverId: string
): { serverUrl: string; serverName: string } {
  const server = getCatalogEntry(serverId);
  if (!server) {
    throw new Error(`Server "${serverId}" not found in M365 catalog.`);
  }

  const envId = store.mcpEnvironmentId;
  if (!envId) {
    throw new Error('MCP Environment ID is not configured.');
  }

  return {
    serverUrl: resolveServerUrl(serverId, envId),
    serverName: server.name
  };
}

async function executeCatalogTool(
  mcpClient: McpClientService,
  store: IFunctionCallStore,
  serverId: string,
  toolName: string,
  args: Record<string, unknown>,
  sourceContext?: IHieSourceContext,
  explicitTargetContext?: IMcpTargetContext
): Promise<{ content: IMcpContent[]; trace: IMcpExecutionTrace }> {
  const { serverUrl, serverName } = resolveCatalogServer(store, serverId);
  const execution = await executeCatalogMcpTool({
    serverId,
    serverName,
    serverUrl,
    toolName,
    rawArgs: args,
    connections: store.mcpConnections,
    getConnections: () => useGrimoireStore.getState().mcpConnections,
    mcpClient,
    sessionHelpers: {
      findExistingSession,
      connectToM365Server
    },
    getToken: store.getToken,
    explicitTargetContext,
    sourceContext,
    taskContext: hybridInteractionEngine.getCurrentTaskContext(),
    artifacts: hybridInteractionEngine.getCurrentArtifacts(),
    currentSiteUrl: store.userContext?.currentSiteUrl
  });

  logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));

  if (!execution.success || !execution.mcpResult) {
    throw new Error(execution.error || `${toolName} failed.`);
  }

  return {
    content: execution.mcpResult.content,
    trace: execution.trace
  };
}

export function resolveShareSubmissionTargetContext(
  formData: IFormData,
  sourceContext?: IHieSourceContext
): IMcpTargetContext | undefined {
  return mergeMcpTargetContexts(
    sourceContext?.targetContext,
    formData.submissionTarget.targetContext
  );
}

function buildShareResolvedArgs(
  formData: IFormData,
  fieldValues: Record<string, string>,
  emailTags: Record<string, string[]>
): Record<string, unknown> {
  if (formData.preset === SHARE_CHAT_PRESET) {
    return {
      recipients: emailTags.recipients || [],
      topic: (fieldValues.topic || '').trim() || undefined
    };
  }

  if (formData.preset === SHARE_CHANNEL_PRESET) {
    const destination = tryParseTeamsChannelDestination((fieldValues.destination || '').trim());
    return {
      teamId: (fieldValues.teamId || '').trim() || destination?.teamId,
      teamName: (fieldValues.teamName || '').trim() || destination?.teamName,
      channelId: (fieldValues.channelId || '').trim() || destination?.channelId,
      channelName: (fieldValues.channelName || '').trim() || destination?.channelName
    };
  }

  return {};
}

export function buildShareSubmissionTrace(
  formData: IFormData,
  fieldValues: Record<string, string>,
  emailTags: Record<string, string[]>,
  sourceContext: IHieSourceContext | undefined,
  result?: IFormSubmissionResult
): IMcpExecutionTrace {
  const targetContext = resolveShareSubmissionTargetContext(formData, sourceContext);
  const rawArgs: Record<string, unknown> = {
    fieldValues: Object.fromEntries(
      Object.entries(fieldValues).filter(([, value]) => typeof value === 'string' && value.trim().length > 0)
    )
  };
  if ((emailTags.recipients || []).length > 0) {
    rawArgs.recipients = emailTags.recipients;
  }

  return {
    serverId: INTERNAL_SHARE_SERVER_ID,
    serverName: 'Internal share flow',
    toolName: formData.preset,
    rawArgs,
    normalizedArgs: rawArgs,
    resolvedArgs: buildShareResolvedArgs(formData, fieldValues, emailTags),
    targetContext,
    targetSource: targetContext?.source || 'none',
    targetSummary: describeMcpTargetContext(targetContext),
    recoverySteps: [],
    finalSummary: result?.message
  };
}

function emitShareSubmissionEvent(
  eventName: string,
  correlationId: string,
  formData: IFormData,
  message: string | undefined,
  sourceContext: IHieSourceContext | undefined,
  success?: boolean
): void {
  hybridInteractionEngine.emitEvent({
    eventName,
    source: 'form',
    surface: 'action-panel',
    correlationId,
    payload: {
      preset: formData.preset,
      formDescription: formData.description,
      message,
      success,
      sourceBlockId: sourceContext?.sourceBlockId,
      sourceBlockType: sourceContext?.sourceBlockType,
      sourceBlockTitle: sourceContext?.sourceBlockTitle,
      sourceArtifactId: sourceContext?.sourceArtifactId,
      sourceTaskKind: sourceContext?.sourceTaskKind,
      sourceEventName: sourceContext?.sourceEventName,
      sourceCorrelationId: sourceContext?.correlationId,
      sourceTurnId: sourceContext?.sourceTurnId,
      sourceRootTurnId: sourceContext?.sourceRootTurnId,
      sourceParentTurnId: sourceContext?.sourceParentTurnId,
      selectedItems: sourceContext?.selectedItems,
      targetContext: resolveShareSubmissionTargetContext(formData, sourceContext)
    },
    exposurePolicy: { mode: 'store-only', relevance: 'contextual' },
    turnId: sourceContext?.sourceTurnId,
    rootTurnId: sourceContext?.sourceRootTurnId,
    parentTurnId: sourceContext?.sourceParentTurnId
  });
}

async function getCurrentUserIdentity(
  mcpClient: McpClientService,
  store: IFunctionCallStore
): Promise<IUserIdentity> {
  const { content } = await executeCatalogTool(mcpClient, store, 'mcp_MeServer', 'GetMyDetails', {
    select: 'id,mail,userPrincipalName'
  });

  const texts = extractMcpTextParts(content);
  for (let i = 0; i < texts.length; i++) {
    const parsed = tryParseJson(texts[i]);
    const record = flattenRecords(parsed)[0];
    if (!record) continue;

    const id = pickString(record, ['id', 'userId']);
    const mail = pickString(record, ['mail', 'email']);
    const userPrincipalName = pickString(record, ['userPrincipalName', 'upn']);
    if (id) {
      return { id, mail, userPrincipalName };
    }
  }

  throw new Error('Could not resolve the current Microsoft 365 user identity.');
}

function extractSelectionItems(serverId: string, toolName: string, content: IMcpContent[]): ISelectionItem[] {
  const mappedBlocks: Array<{ type: string; data: unknown }> = [];
  mapMcpResultToBlock(serverId, toolName, content, (block) => {
    mappedBlocks.push({ type: block.type, data: block.data });
  });

  const mappedSelection = mappedBlocks.find((block) => block.type === 'selection-list');
  if (mappedSelection) {
    const data = mappedSelection.data as ISelectionListData;
    return data.items.filter((item) => item.id && item.label);
  }

  const fallbackItems: ISelectionItem[] = [];
  const texts = extractMcpTextParts(content);
  for (let i = 0; i < texts.length; i++) {
    const parsed = tryParseJson(texts[i]);
    const records = flattenRecords(parsed);
    records.forEach((record) => {
      const id = pickString(record, ['id', 'teamId', 'channelId']);
      const label = pickString(record, ['displayName', 'name', 'topic', 'label']);
      if (id && label) {
        fallbackItems.push({ id, label, description: pickString(record, ['description']) });
      }
    });
  }

  return fallbackItems;
}

async function listTeams(
  mcpClient: McpClientService,
  store: IFunctionCallStore,
  userId: string
): Promise<ISelectionItem[]> {
  const { content } = await executeCatalogTool(mcpClient, store, 'mcp_TeamsServer', 'ListTeams', { userId });
  return extractSelectionItems('mcp_TeamsServer', 'ListTeams', content);
}

async function listChannels(
  mcpClient: McpClientService,
  store: IFunctionCallStore,
  teamId: string
): Promise<ISelectionItem[]> {
  const { content } = await executeCatalogTool(mcpClient, store, 'mcp_TeamsServer', 'ListChannels', { teamId });
  return extractSelectionItems('mcp_TeamsServer', 'ListChannels', content);
}

export async function loadTeamsChannelDestinationOptions(
  store: IFunctionCallStore
): Promise<Array<{ key: string; text: string }>> {
  const proxyConfig = store.proxyConfig;
  if (!proxyConfig) {
    return [];
  }

  const mcpClient = new McpClientService(proxyConfig.proxyUrl, proxyConfig.proxyApiKey);
  const me = await getCurrentUserIdentity(mcpClient, store);
  const teams = await listTeams(mcpClient, store, me.id);

  const optionGroups = await Promise.all(teams.map(async (team) => {
    try {
      const channels = await listChannels(mcpClient, store, team.id);
      return channels.map((channel) => ({
        key: JSON.stringify({
          teamId: team.id,
          channelId: channel.id,
          teamName: team.label,
          channelName: channel.label
        }),
        text: `${team.label} / ${channel.label}`
      }));
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      logService.warning('mcp', `Could not load channels for ${team.label}: ${message}`);
      return [];
    }
  }));

  return optionGroups
    .flat()
    .sort((left, right) => left.text.localeCompare(right.text, undefined, { sensitivity: 'base' }));
}

async function createChat(
  mcpClient: McpClientService,
  store: IFunctionCallStore,
  membersUpns: string[],
  topic?: string
): Promise<string> {
  const chatType = membersUpns.length === 2 ? 'oneOnOne' : 'group';
  const { content } = await executeCatalogTool(mcpClient, store, 'mcp_TeamsServer', 'CreateChat', {
    chatType,
    members_upns: membersUpns,
    ...(chatType === 'group' && topic ? { topic } : {})
  });

  const texts = extractMcpTextParts(content);
  for (let i = 0; i < texts.length; i++) {
    const parsed = tryParseJson(texts[i]);
    const record = flattenRecords(parsed)[0];
    if (!record) continue;
    const chatId = pickString(record, ['id', 'chatId']);
    if (chatId) {
      return chatId;
    }
  }

  throw new Error('The Teams chat was created, but its chat ID was not returned.');
}

async function postTeamsMessage(
  mcpClient: McpClientService,
  store: IFunctionCallStore,
  chatId: string,
  content: string
): Promise<void> {
  const payload = buildTeamsMessageBody(content);
  await executeCatalogTool(mcpClient, store, 'mcp_TeamsServer', 'PostMessage', {
    chatId,
    content: payload.content,
    contentType: payload.contentType
  });
}

async function postTeamsChannelMessage(
  mcpClient: McpClientService,
  store: IFunctionCallStore,
  teamId: string,
  channelId: string,
  content: string
): Promise<void> {
  const payload = buildTeamsMessageBody(content);
  await executeCatalogTool(mcpClient, store, 'mcp_TeamsServer', 'PostChannelMessage', {
    teamId,
    channelId,
    content: payload.content,
    contentType: payload.contentType
  });
}

function dedupe(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  values.forEach((value) => {
    const trimmed = value.trim();
    if (!trimmed) return;
    const normalized = trimmed.toLowerCase();
    if (seen.has(normalized)) return;
    seen.add(normalized);
    result.push(trimmed);
  });
  return result;
}

async function submitTeamsChatShare(
  fieldValues: Record<string, string>,
  emailTags: Record<string, string[]>,
  store: IFunctionCallStore
): Promise<IFormSubmissionResult> {
  const proxyConfig = store.proxyConfig;
  if (!proxyConfig) {
    return { success: false, message: 'No proxy config available.' };
  }

  const recipients = emailTags.recipients || [];
  const message = (fieldValues.content || '').trim();
  if (recipients.length === 0 || !message) {
    return { success: false, message: 'Recipients and message are required.' };
  }

  const mcpClient = new McpClientService(proxyConfig.proxyUrl, proxyConfig.proxyApiKey);
  const me = await getCurrentUserIdentity(mcpClient, store);
  const currentUserAddress = me.userPrincipalName || me.mail || store.userContext?.email || store.userContext?.loginName;
  if (!currentUserAddress) {
    return { success: false, message: 'Could not resolve your Teams identity.' };
  }

  const membersUpns = dedupe([currentUserAddress, ...recipients]);
  if (membersUpns.length < 2) {
    return { success: false, message: 'A Teams chat share needs at least one recipient other than you.' };
  }

  const chatId = await createChat(mcpClient, store, membersUpns, fieldValues.topic || undefined);
  await postTeamsMessage(mcpClient, store, chatId, message);

  return {
    success: true,
    message: `Shared to a new Teams chat with ${recipients.length} recipient${recipients.length === 1 ? '' : 's'}.`
  };
}

async function submitTeamsChannelShare(
  fieldValues: Record<string, string>,
  store: IFunctionCallStore
): Promise<IFormSubmissionResult> {
  const proxyConfig = store.proxyConfig;
  if (!proxyConfig) {
    return { success: false, message: 'No proxy config available.' };
  }

  const directTeamId = (fieldValues.teamId || '').trim();
  const directChannelId = (fieldValues.channelId || '').trim();
  const teamName = (fieldValues.teamName || '').trim();
  const channelName = (fieldValues.channelName || '').trim();
  const message = (fieldValues.content || '').trim();
  const destination = tryParseTeamsChannelDestination((fieldValues.destination || '').trim());
  if (!message) {
    return { success: false, message: 'A destination and message are required.' };
  }

  const mcpClient = new McpClientService(proxyConfig.proxyUrl, proxyConfig.proxyApiKey);
  if (directTeamId && directChannelId) {
    await postTeamsChannelMessage(mcpClient, store, directTeamId, directChannelId, message);
    return {
      success: true,
      message: `Shared to ${teamName || 'selected team'} / ${channelName || 'selected channel'}.`
    };
  }

  if (destination) {
    await postTeamsChannelMessage(mcpClient, store, destination.teamId, destination.channelId, message);
    return {
      success: true,
      message: `Shared to ${destination.teamName || 'selected team'} / ${destination.channelName || 'selected channel'}.`
    };
  }

  if (!teamName || !channelName) {
    return { success: false, message: 'Team name, channel name, and message are required.' };
  }

  const me = await getCurrentUserIdentity(mcpClient, store);
  const teams = await listTeams(mcpClient, store, me.id);
  const teamMatch = resolveNamedSelectionItem(teams, teamName, 'team');
  if (!teamMatch.item) {
    return { success: false, message: teamMatch.error || `No team matched "${teamName}".` };
  }

  const channels = await listChannels(mcpClient, store, teamMatch.item.id);
  const channelMatch = resolveNamedSelectionItem(channels, channelName, 'channel');
  if (!channelMatch.item) {
    return { success: false, message: channelMatch.error || `No channel matched "${channelName}".` };
  }

  await postTeamsChannelMessage(mcpClient, store, teamMatch.item.id, channelMatch.item.id, message);
  return {
    success: true,
    message: `Shared to ${teamMatch.item.label} / ${channelMatch.item.label}.`
  };
}

export async function executeShareSubmission(
  formData: IFormData,
  fieldValues: Record<string, string>,
  emailTags: Record<string, string[]>,
  store: IFunctionCallStore,
  sourceContext?: IHieSourceContext
): Promise<IFormSubmissionResult> {
  const submissionCorrelationId = createCorrelationId('shareform');

  if (!isInternalSharePreset(formData.preset) && formData.submissionTarget.serverId !== INTERNAL_SHARE_SERVER_ID) {
    emitShareSubmissionEvent(
      'form.execution.failed',
      submissionCorrelationId,
      formData,
      'This form is not configured for an internal share flow.',
      sourceContext,
      false
    );
    return { success: false, message: 'This form is not configured for an internal share flow.' };
  }

  emitShareSubmissionEvent(
    'form.execution.started',
    submissionCorrelationId,
    formData,
    undefined,
    sourceContext
  );

  try {
    if (formData.preset === SHARE_CHAT_PRESET) {
      const result = await submitTeamsChatShare(fieldValues, emailTags, store);
      logService.debug('mcp', 'Share execution trace', JSON.stringify(
        buildShareSubmissionTrace(formData, fieldValues, emailTags, sourceContext, result)
      ));
      emitShareSubmissionEvent(
        result.success ? 'form.execution.completed' : 'form.execution.failed',
        submissionCorrelationId,
        formData,
        result.message,
        sourceContext,
        result.success
      );
      return result;
    }
    if (formData.preset === SHARE_CHANNEL_PRESET) {
      const result = await submitTeamsChannelShare(fieldValues, store);
      logService.debug('mcp', 'Share execution trace', JSON.stringify(
        buildShareSubmissionTrace(formData, fieldValues, emailTags, sourceContext, result)
      ));
      emitShareSubmissionEvent(
        result.success ? 'form.execution.completed' : 'form.execution.failed',
        submissionCorrelationId,
        formData,
        result.message,
        sourceContext,
        result.success
      );
      return result;
    }

    emitShareSubmissionEvent(
      'form.execution.failed',
      submissionCorrelationId,
      formData,
      'Unknown share form preset.',
      sourceContext,
      false
    );
    return { success: false, message: 'Unknown share form preset.' };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Share failed.';
    logService.error('mcp', `Share submission error: ${message}`);
    logService.debug('mcp', 'Share execution trace', JSON.stringify(
      buildShareSubmissionTrace(
        formData,
        fieldValues,
        emailTags,
        sourceContext,
        { success: false, message }
      )
    ));
    emitShareSubmissionEvent(
      'form.execution.failed',
      submissionCorrelationId,
      formData,
      message,
      sourceContext,
      false
    );
    return { success: false, message };
  }
}
