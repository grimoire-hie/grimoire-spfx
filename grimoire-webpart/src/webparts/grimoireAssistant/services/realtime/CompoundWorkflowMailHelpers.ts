import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IBlock, IMarkdownData } from '../../models/IBlock';
import { getCatalogEntry, resolveServerUrl } from '../../models/McpServerCatalog';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { McpClientService } from '../mcp/McpClientService';
import { executeCatalogMcpTool, extractStructuredMcpPayload } from '../mcp/McpExecutionAdapter';
import { connectToM365Server, findExistingSession } from '../tools/ToolRuntimeSharedHelpers';

export class CompoundWorkflowStopError extends Error {
  public readonly assistantMessage: string;

  constructor(message: string, assistantMessage?: string) {
    super(message);
    this.name = 'CompoundWorkflowStopError';
    this.assistantMessage = assistantMessage || message;
    Object.setPrototypeOf(this, CompoundWorkflowStopError.prototype);
  }
}

export interface IMarkdownEmailCandidate {
  index: number;
  messageId?: string;
  subject: string;
  sender?: string;
  date?: string;
  preview?: string;
}

export type MailThreadMatchKind =
  | 'exact-subject'
  | 'subject-prefix'
  | 'subject-contains'
  | 'broad-fallback';

export interface IMailThreadCandidate {
  messageId: string;
  subject: string;
  sender?: string;
  date?: string;
  preview?: string;
  matchScore: number;
  matchKind: MailThreadMatchKind;
}

export interface IMailDiscussionRecipientSets {
  to: string[];
  cc: string[];
}

const KV_FIELD_PATTERN = /^\*{0,2}([^*]+?)\*{0,2}[:\s]+(.+)$/;
const HEADING_PREFIX_PATTERN = /^#+\s*/;

export function extractMarkdownEmailCandidates(block: IBlock): IMarkdownEmailCandidate[] {
  if (block.type !== 'markdown') return [];

  const data = block.data as IMarkdownData;
  const lines = (data.content || '').split('\n');
  const groups: Array<{ index: number; lines: string[] }> = [];
  let currentGroup: { index: number; lines: string[] } | undefined;

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    const normalizedLine = trimmed.replace(/\*\*/g, '');
    const startMatch = /^(?:#+\s*)?(\d+)[.)]?\s*(.*)$/.exec(normalizedLine);
    if (startMatch) {
      if (currentGroup) groups.push(currentGroup);
      currentGroup = { index: parseInt(startMatch[1], 10), lines: [] };
      if (startMatch[2]) {
        currentGroup.lines.push(startMatch[2]);
      }
      continue;
    }
    if (!currentGroup) continue;
    if (/^[-*_]{3,}$/.test(trimmed)) continue;
    currentGroup.lines.push(lines[i]);
  }
  if (currentGroup) groups.push(currentGroup);

  const candidates: IMarkdownEmailCandidate[] = [];
  const itemIds = data.itemIds || {};
  for (let i = 0; i < groups.length; i++) {
    const fields: Record<string, string> = {};
    for (let j = 0; j < groups[i].lines.length; j++) {
      const line = groups[i].lines[j].trim();
      const kvMatch = KV_FIELD_PATTERN.exec(line.replace(HEADING_PREFIX_PATTERN, ''));
      if (!kvMatch) continue;
      const key = kvMatch[1].replace(/\*/g, '').replace(/:$/, '').trim().toLowerCase();
      const value = kvMatch[2].replace(/\*+/g, '').trim();
      if (key && value) {
        fields[key] = value;
      }
    }

    const subject = fields.subject || '';
    if (!subject) continue;
    candidates.push({
      index: groups[i].index,
      messageId: itemIds[groups[i].index],
      subject,
      sender: fields.from,
      date: fields.date,
      preview: fields.preview
    });
  }
  return candidates;
}

function normalizeEmailAddress(value: string | undefined): string | undefined {
  const trimmed = value?.trim().toLowerCase() || '';
  if (!trimmed || !trimmed.includes('@')) {
    return undefined;
  }
  return trimmed;
}

function extractEmailAddressFromRecipient(value: unknown): string | undefined {
  if (typeof value === 'string') {
    return normalizeEmailAddress(value);
  }
  if (!value || typeof value !== 'object') {
    return undefined;
  }

  const record = value as Record<string, unknown>;
  return normalizeEmailAddress(
    typeof record.address === 'string'
      ? record.address
      : (typeof record.mail === 'string'
        ? record.mail
        : (typeof record.userPrincipalName === 'string'
          ? record.userPrincipalName
          : undefined))
  ) || extractEmailAddressFromRecipient(record.emailAddress);
}

function extractEmailAddresses(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value
      .map((entry) => extractEmailAddressFromRecipient(entry))
      .filter((entry): entry is string => !!entry);
  }

  const single = extractEmailAddressFromRecipient(value);
  return single ? [single] : [];
}

function dedupeEmailAddresses(values: string[], excluded: Set<string>): string[] {
  const seen = new Set<string>();
  const deduped: string[] = [];
  for (let i = 0; i < values.length; i++) {
    const normalized = normalizeEmailAddress(values[i]);
    if (!normalized || excluded.has(normalized) || seen.has(normalized)) {
      continue;
    }
    seen.add(normalized);
    deduped.push(normalized);
  }
  return deduped;
}

function getCurrentUserEmailAliases(): Set<string> {
  const store = useGrimoireStore.getState();
  const aliases = new Set<string>();
  const email = normalizeEmailAddress(store.userContext?.email);
  const loginName = normalizeEmailAddress(store.userContext?.loginName);
  if (email) {
    aliases.add(email);
  }
  if (loginName) {
    aliases.add(loginName);
  }
  return aliases;
}

export async function fetchMailDiscussionRecipients(messageId: string): Promise<IMailDiscussionRecipientSets> {
  const store = useGrimoireStore.getState();
  const proxyConfig = store.proxyConfig;
  const envId = store.mcpEnvironmentId;
  const mailServer = getCatalogEntry('mcp_MailTools');
  if (!proxyConfig || !envId || !mailServer) {
    throw new CompoundWorkflowStopError(
      'Mail recipient lookup is unavailable.',
      'I could not access the selected email details to prepare the recipients.'
    );
  }

  const execution = await executeCatalogMcpTool({
    serverId: mailServer.id,
    serverName: mailServer.name,
    serverUrl: resolveServerUrl(mailServer.id, envId),
    toolName: 'GetMessage',
    rawArgs: {
      id: messageId,
      bodyPreviewOnly: true
    },
    connections: store.mcpConnections,
    getConnections: () => useGrimoireStore.getState().mcpConnections,
    mcpClient: new McpClientService(proxyConfig.proxyUrl, proxyConfig.proxyApiKey),
    sessionHelpers: {
      findExistingSession,
      connectToM365Server
    },
    getToken: store.getToken,
    taskContext: hybridInteractionEngine.getCurrentTaskContext(),
    artifacts: hybridInteractionEngine.getCurrentArtifacts(),
    currentSiteUrl: store.userContext?.currentSiteUrl
  });

  if (!execution.success || !execution.mcpResult) {
    throw new CompoundWorkflowStopError(
      execution.error || 'Mail recipient lookup failed.',
      'I found the email, but I could not read its recipient list.'
    );
  }

  const payload = extractStructuredMcpPayload(execution.mcpResult.content).payload as Record<string, unknown> | undefined;
  if (!payload || typeof payload !== 'object') {
    throw new CompoundWorkflowStopError(
      'Mail recipient lookup returned no payload.',
      'I found the email, but its recipient list was unavailable.'
    );
  }

  const currentUserAliases = getCurrentUserEmailAliases();
  const fromRecipients = extractEmailAddresses(payload.from);
  const toRecipients = extractEmailAddresses(payload.toRecipients);
  const ccRecipients = extractEmailAddresses(payload.ccRecipients);
  const to = dedupeEmailAddresses([...fromRecipients, ...toRecipients], currentUserAliases);
  const toSet = new Set<string>(to);
  const excludedCcRecipients = new Set<string>(Array.from(currentUserAliases).concat(Array.from(toSet)));
  const cc = dedupeEmailAddresses(ccRecipients, excludedCcRecipients);

  if (to.length === 0 && cc.length === 0) {
    throw new CompoundWorkflowStopError(
      'No recipients were available on the selected email.',
      'I found the email, but there were no usable recipients to prefill in the draft.'
    );
  }

  return { to, cc };
}

export function normalizeMailSubjectForMatch(value: string | undefined): string {
  const trimmed = value?.trim().toLowerCase() || '';
  if (!trimmed) {
    return '';
  }

  let normalized = trimmed;
  while (/^(?:fwd|fw|r[ée]f|rif|re|wg|aw|sv|odp|vs|ynt)(?:\s*:\s*|\s+)/i.test(normalized)) {
    normalized = normalized.replace(/^(?:fwd|fw|r[ée]f|rif|re|wg|aw|sv|odp|vs|ynt)(?:\s*:\s*|\s+)/i, '').trim();
  }

  return normalized
    .replace(/[_-]+/g, ' ')
    .replace(/[^0-9a-z\u00c0-\u024f\u0400-\u04ff\u0600-\u06ff\u0590-\u05ff\u3040-\u30ff\u4e00-\u9fff\uac00-\ud7af\u0e00-\u0e7f]+/gi, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

export function tokenizeNormalizedText(value: string): string[] {
  return value
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 1);
}

export function estimateMailDateRecencyBoost(dateHint: string | undefined): number {
  const normalized = dateHint?.trim().toLowerCase() || '';
  if (!normalized) {
    return 0;
  }
  if (/(today|this morning|this afternoon|this evening|just now)/i.test(normalized)) {
    return 4;
  }
  if (/(yesterday|last night)/i.test(normalized)) {
    return 2;
  }
  return 0;
}

export function scoreMailThreadCandidate(
  candidate: IMarkdownEmailCandidate,
  query: string,
  mode: 'subject-first' | 'fallback'
): IMailThreadCandidate | undefined {
  const messageId = candidate.messageId?.trim();
  if (!messageId) {
    return undefined;
  }

  const normalizedQuery = normalizeMailSubjectForMatch(query);
  const normalizedSubject = normalizeMailSubjectForMatch(candidate.subject);
  if (!normalizedQuery || !normalizedSubject) {
    return undefined;
  }

  const queryTokens = tokenizeNormalizedText(normalizedQuery);
  const subjectTokens = new Set(tokenizeNormalizedText(normalizedSubject));
  let matchKind: MailThreadMatchKind | undefined;
  let matchScore = 0;

  if (normalizedSubject === normalizedQuery) {
    matchKind = 'exact-subject';
    matchScore = 100;
  } else if (normalizedSubject.startsWith(normalizedQuery) || normalizedQuery.startsWith(normalizedSubject)) {
    matchKind = 'subject-prefix';
    matchScore = 88;
  } else if (
    normalizedSubject.includes(normalizedQuery)
    || (queryTokens.length > 0 && queryTokens.every((token) => subjectTokens.has(token)))
  ) {
    matchKind = mode === 'fallback' ? 'broad-fallback' : 'subject-contains';
    matchScore = mode === 'fallback' ? 72 : 78;
  } else if (mode === 'fallback' && candidate.preview) {
    const normalizedPreview = normalizeMailSubjectForMatch(candidate.preview);
    const previewTokens = new Set(tokenizeNormalizedText(normalizedPreview));
    if (
      normalizedPreview.includes(normalizedQuery)
      || (queryTokens.length > 0 && queryTokens.every((token) => previewTokens.has(token)))
    ) {
      matchKind = 'broad-fallback';
      matchScore = 68;
    } else {
      return undefined;
    }
  } else {
    return undefined;
  }

  matchScore += estimateMailDateRecencyBoost(candidate.date);

  return {
    messageId,
    subject: candidate.subject,
    sender: candidate.sender,
    date: candidate.date,
    preview: candidate.preview,
    matchScore,
    matchKind
  };
}

export function dedupeMailThreadCandidates(candidates: IMailThreadCandidate[]): IMailThreadCandidate[] {
  const byMessageId = new Map<string, IMailThreadCandidate>();
  for (let i = 0; i < candidates.length; i++) {
    const current = byMessageId.get(candidates[i].messageId);
    if (!current || current.matchScore < candidates[i].matchScore) {
      byMessageId.set(candidates[i].messageId, candidates[i]);
    }
  }

  return Array.from(byMessageId.values()).sort((left, right) => {
    if (right.matchScore !== left.matchScore) {
      return right.matchScore - left.matchScore;
    }
    return left.subject.localeCompare(right.subject);
  });
}

export function getPlausibleMailThreadCandidates(
  block: IBlock | undefined,
  query: string,
  mode: 'subject-first' | 'fallback'
): IMailThreadCandidate[] {
  if (!block) {
    return [];
  }

  const scored = extractMarkdownEmailCandidates(block)
    .map((candidate) => scoreMailThreadCandidate(candidate, query, mode))
    .filter((candidate): candidate is IMailThreadCandidate => !!candidate);
  const deduped = dedupeMailThreadCandidates(scored);
  const minScore = mode === 'fallback' ? 65 : 72;
  return deduped.filter((candidate) => candidate.matchScore >= minScore);
}

export function tryDropLeadingProjectPrefix(query: string): string | undefined {
  const tokens = query.trim().split(/\s+/);
  if (tokens.length < 3) return undefined;
  const shortened = tokens.slice(1).join(' ');
  return shortened.length >= 3 ? shortened : undefined;
}

export function buildMailThreadSubjectFirstSearchQuery(query: string): string {
  return `Find emails whose subject is exactly or very close to "${query}". Prefer subject matches over preview or body mentions. Exclude drafts.`;
}

export function buildMailThreadFallbackSearchQuery(query: string): string {
  return `Find emails about "${query}". Prefer messages whose subject closely matches the phrase. Exclude drafts.`;
}

export function buildMailThreadChooserDescription(candidate: IMailThreadCandidate): string | undefined {
  const parts = [candidate.sender, candidate.date].filter((value): value is string => !!value && value.trim().length > 0);
  return parts.length > 0 ? parts.join(' \u2022 ') : undefined;
}
