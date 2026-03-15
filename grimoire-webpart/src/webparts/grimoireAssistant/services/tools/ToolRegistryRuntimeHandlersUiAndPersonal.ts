import type { Expression } from '../avatar/ExpressionEngine';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import type {
  IBlock,
  IActivityFeedData,
  IChartData,
  IConfirmationDialogData,
  IDocumentLibraryData,
  IErrorData,
  IFormData,
  IFormFieldDefinition,
  IFormSubmissionTarget,
  IInfoCardData,
  IMarkdownData,
  IPermissionsViewData,
  IProgressTrackerData,
  ISearchResult,
  ISearchResultsData,
  ISelectionListData,
  IUserCardData,
  FormPresetId
} from '../../models/IBlock';
import { createBlock } from '../../models/IBlock';
import { getFormPreset } from '../forms/FormPresets';
import { logService } from '../logging/LogService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { getTools } from '../realtime/ToolRegistry';
import { GraphService } from '../graph/GraphService';
import { PersistenceService } from '../context/PersistenceService';
import {
  findCapabilityPermissionInspectionTool,
  getCapabilityFocusLabel,
  normalizeCapabilityFocus,
  resolveServerUrl
} from '../../models/McpServerCatalog';
import { McpClientService } from '../mcp/McpClientService';
import {
  ensureCatalogMcpConnection,
  executeCatalogMcpTool,
  resolveEffectiveMcpTargetContext
} from '../mcp/McpExecutionAdapter';
import { deriveMcpTargetContextFromUnknown, mergeMcpTargetContexts } from '../mcp/McpTargetContext';
import { mapMcpResultToBlock } from '../mcp/McpResultMapper';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type { IToolRuntimeHandlerDeps, ToolRuntimeHandler, ToolRuntimeHandlerResult } from './ToolRuntimeHandlerTypes';
import type { UiPersonalRuntimeToolName } from './ToolRuntimeHandlerPartitions';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import { trackCreatedBlock, trackToolCompletion, trackUpdatedBlock } from './ToolRuntimeHieHelpers';
import { resolveUrlFromBlocks } from './ToolRuntimeUrlHelpers';
import { completeOutcome, errorOutcome } from './ToolRuntimeOutcomeHelpers';
import { connectToM365Server, extractMcpReply, findExistingSession } from './ToolRuntimeSharedHelpers';
import { shouldSuppressSelectionList } from './ToolSelectionListGuard';
import { SessionShareFormatter, hasShareableSessionContent } from '../sharing/SessionShareFormatter';
import { loadTeamsChannelDestinationOptions } from '../sharing/ShareSubmissionService';
import { isInitialized as isPnpInitialized } from '../pnp/pnpContext';
import {
  readComposeShareScope,
  resolveContextualComposeScope
} from '../sharing/ContextualComposeScope';
import {
  dedupeStringList,
  toMcpToolDescriptor,
  findConnectedServerSession,
  buildCapabilityServerViews,
  buildCapabilityOverviewMarkdown,
  buildFocusedCapabilityMarkdown
} from './ToolHandlerCapabilityBrowser';
import type { IMcpToolDescriptor } from './ToolHandlerCapabilityBrowser';
import {
  buildContextualFileName,
  extractFocusedBlockContent,
  looksLikeWordDocumentCreateRequest,
  sanitizeCreatedFileName
} from './FormCreateHelpers';

interface IComposeContentPrefillTarget {
  subjectKey?: string;
  contentKey?: string;
  allowAttachments?: boolean;
}

interface IResolvedComposePreset {
  preset: FormPresetId;
  prefill: Record<string, string>;
}

interface IComposeFieldEnhancementResult {
  fields: IFormFieldDefinition[];
  descriptionNote?: string;
}

const TEAMS_CHANNEL_PICKER_FALLBACK_NOTE = 'The Teams channel picker is unavailable right now, so enter Team name and Channel name manually. Names are resolved when you submit.';

const COMPOSE_CONTENT_PREFILL_TARGETS: Partial<Record<FormPresetId, IComposeContentPrefillTarget>> = {
  'email-compose': { subjectKey: 'subject', contentKey: 'body', allowAttachments: true },
  'email-reply': { contentKey: 'comment' },
  'email-forward': { contentKey: 'introComment', allowAttachments: true },
  'email-reply-all-thread': { contentKey: 'introComment' },
  'teams-message': { contentKey: 'content' },
  'teams-channel-message': { contentKey: 'content' },
  'share-teams-chat': { subjectKey: 'topic', contentKey: 'content' },
  'share-teams-channel': { contentKey: 'content' },
  'file-create': { contentKey: 'contentText' },
  'word-document-create': { contentKey: 'contentText' }
};

const THIN_SHARE_PREFILL_PATTERNS: ReadonlyArray<RegExp> = [
  /available in the .*panel/i,
  /shown in the .*results/i,
  /listed in the action panel/i,
  /please add recipients/i,
  /please review/i,
  /please attach/i,
  /links? to (?:these|those|the) .* (?:available|shown)/i,
  /\bi found these\b/i,
  /\bi found the following\b/i,
  /\bshared from grimoire\b/i
];

function hasNonEmptyStaticArg(staticArgs: Record<string, unknown>, key: string): boolean {
  const value = staticArgs[key];
  return typeof value === 'string' && value.trim().length > 0;
}

function normalizeComposePrefillAliases(
  prefill: Record<string, string>,
  formFields: IFormFieldDefinition[]
): Record<string, string> {
  const nextPrefill = { ...prefill };
  const fieldKeys = new Set(formFields.map((field) => field.key));

  if (fieldKeys.has('filename') && !nextPrefill.filename?.trim()) {
    const filenameLikeValue = nextPrefill.file_name || nextPrefill.fileName || nextPrefill.name;
    if (typeof filenameLikeValue === 'string' && filenameLikeValue.trim()) {
      nextPrefill.filename = filenameLikeValue;
    }
  }

  if (fieldKeys.has('contentText') && !nextPrefill.contentText?.trim()) {
    const contentLikeValue = nextPrefill.file_content || nextPrefill.fileContent || nextPrefill.content || nextPrefill.body;
    if (typeof contentLikeValue === 'string' && contentLikeValue.trim()) {
      nextPrefill.contentText = contentLikeValue;
    }
  }

  if (fieldKeys.has('content') && !nextPrefill.content?.trim()) {
    const messageLikeValue = nextPrefill.message || nextPrefill.body || nextPrefill.comment;
    if (typeof messageLikeValue === 'string' && messageLikeValue.trim()) {
      nextPrefill.content = messageLikeValue;
    }
  }

  if (fieldKeys.has('topic') && !nextPrefill.topic?.trim()) {
    const topicLikeValue = nextPrefill.subject || nextPrefill.title;
    if (typeof topicLikeValue === 'string' && topicLikeValue.trim()) {
      nextPrefill.topic = topicLikeValue;
    }
  }

  if (fieldKeys.has('subject') && !nextPrefill.subject?.trim()) {
    const subjectLikeValue = nextPrefill.title || nextPrefill.topic || nextPrefill.name;
    if (typeof subjectLikeValue === 'string' && subjectLikeValue.trim()) {
      nextPrefill.subject = subjectLikeValue;
    }
  }

  if (fieldKeys.has('body') && !nextPrefill.body?.trim()) {
    const bodyLikeValue = nextPrefill.message || nextPrefill.content;
    if (typeof bodyLikeValue === 'string' && bodyLikeValue.trim()) {
      nextPrefill.body = bodyLikeValue;
    }
  }

  if (fieldKeys.has('startDateTime') && !nextPrefill.startDateTime?.trim()) {
    const startLikeValue = nextPrefill.start || nextPrefill.startDate || nextPrefill.start_date_time;
    if (typeof startLikeValue === 'string' && startLikeValue.trim()) {
      nextPrefill.startDateTime = startLikeValue;
    }
  }

  if (fieldKeys.has('endDateTime') && !nextPrefill.endDateTime?.trim()) {
    const endLikeValue = nextPrefill.end || nextPrefill.endDate || nextPrefill.end_date_time;
    if (typeof endLikeValue === 'string' && endLikeValue.trim()) {
      nextPrefill.endDateTime = endLikeValue;
    }
  }

  if (fieldKeys.has('attendeeEmails') && !nextPrefill.attendeeEmails?.trim()) {
    const attendeeLikeValue = nextPrefill.attendees || nextPrefill.attendee || nextPrefill.attendee_emails;
    if (typeof attendeeLikeValue === 'string' && attendeeLikeValue.trim()) {
      nextPrefill.attendeeEmails = attendeeLikeValue;
    }
  }

  if (fieldKeys.has('bodyContent') && !nextPrefill.bodyContent?.trim()) {
    const bodyContentValue = nextPrefill.body || nextPrefill.description || nextPrefill.content || nextPrefill.message;
    if (typeof bodyContentValue === 'string' && bodyContentValue.trim()) {
      nextPrefill.bodyContent = bodyContentValue;
    }
  }

  return nextPrefill;
}

function normalizePrefillValue(value: string): string {
  if (!value.includes('%')) {
    return value;
  }

  return value
    .replace(/%0D%0A/gi, '\n')
    .replace(/%0D/gi, '\n')
    .replace(/%0A/gi, '\n');
}

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function parseJsonLikeArg(raw: unknown): unknown {
  if (typeof raw === 'string') {
    return JSON.parse(raw);
  }
  return raw;
}

function normalizeDateTimeForForm(value: unknown): string {
  if (typeof value !== 'string' || !value.trim()) {
    return '';
  }
  const raw = value.trim();

  // Already ISO with T separator: "2026-03-16T10:00" or "2026-03-16T10:00:00"
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(raw)) {
    return raw.substring(0, 16);
  }

  // Space separator: "2026-03-16 10:00"
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}/.test(raw)) {
    return raw.substring(0, 16).replace(' ', 'T');
  }

  // Date only: "2026-03-16"
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw + 'T09:00';
  }

  // Relative dates
  const lower = raw.toLowerCase();
  const now = new Date();
  let resolved: Date | undefined;
  if (lower === 'today') {
    resolved = now;
  } else if (lower === 'tomorrow') {
    resolved = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  }
  if (resolved) {
    const y = resolved.getFullYear();
    const m = String(resolved.getMonth() + 1).padStart(2, '0');
    const d = String(resolved.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}T09:00`;
  }

  // Try native Date parsing as last resort
  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    const h = String(parsed.getHours()).padStart(2, '0');
    const min = String(parsed.getMinutes()).padStart(2, '0');
    return `${y}-${m}-${d}T${h}:${min}`;
  }

  return '';
}

function normalizeParsedPrefill(prefill: Record<string, unknown>): Record<string, string> {
  const normalized: Record<string, string> = {};
  Object.keys(prefill).forEach((key) => {
    const value = prefill[key];
    if (typeof value === 'string') {
      normalized[key] = normalizePrefillValue(value);
      return;
    }

    if (typeof value === 'number' || typeof value === 'boolean') {
      normalized[key] = String(value);
      return;
    }

    if (Array.isArray(value)) {
      const serialized = value
        .filter((entry): entry is string | number | boolean => (
          typeof entry === 'string'
          || typeof entry === 'number'
          || typeof entry === 'boolean'
        ))
        .map((entry) => normalizePrefillValue(String(entry)))
        .filter((entry) => entry.trim().length > 0)
        .join('; ');
      if (serialized) {
        normalized[key] = serialized;
      }
    }
  });

  if (!normalized.filename?.trim()) {
    const filename = pickResolvedString(normalized.file_name, normalized.fileName, normalized.name);
    if (filename) {
      normalized.filename = filename;
    }
  }

  if (!normalized.contentText?.trim()) {
    const contentText = pickResolvedString(normalized.file_content, normalized.fileContent);
    if (contentText) {
      normalized.contentText = contentText;
    }
  }

  if (!normalized.siteUrl?.trim()) {
    const siteUrl = pickResolvedString(normalized.site_url, normalized.targetSiteUrl);
    if (siteUrl) {
      normalized.siteUrl = siteUrl;
    }
  }

  if (!normalized.documentLibraryName?.trim()) {
    const documentLibraryName = pickResolvedString(
      normalized.documentLibraryName,
      normalized.library_name,
      normalized.libraryName,
      normalized.targetLibrary
    );
    if (documentLibraryName) {
      normalized.documentLibraryName = documentLibraryName;
    }
  }

  return normalized;
}

function readStaticArgStringList(staticArgs: Record<string, unknown>, key: string): string[] {
  const raw = staticArgs[key];
  if (Array.isArray(raw)) {
    return raw.filter((value): value is string => typeof value === 'string' && value.trim().length > 0);
  }
  if (typeof raw === 'string' && raw.trim().length > 0) {
    return [raw];
  }
  return [];
}

function normalizeComposeStaticArgs(staticArgs: Record<string, unknown>): Record<string, unknown> {
  const nextStaticArgs = { ...staticArgs };
  const attachmentUris = dedupeStringList([
    ...readStaticArgStringList(nextStaticArgs, 'attachmentUris'),
    ...readStaticArgStringList(nextStaticArgs, 'attachmentUrls')
  ]);

  if (attachmentUris.length > 0) {
    nextStaticArgs.attachmentUris = attachmentUris;
  }

  const siteUrl = pickResolvedString(
    nextStaticArgs.siteUrl,
    nextStaticArgs.site_url,
    nextStaticArgs.targetSiteUrl
  );
  if (siteUrl) {
    nextStaticArgs.siteUrl = siteUrl;
  }

  const documentLibraryName = pickResolvedString(
    nextStaticArgs.documentLibraryName,
    nextStaticArgs.libraryName,
    nextStaticArgs.library_name,
    nextStaticArgs.targetLibrary
  );
  if (documentLibraryName) {
    nextStaticArgs.documentLibraryName = documentLibraryName;
  }

  delete nextStaticArgs.attachmentUrls;
  return nextStaticArgs;
}

function appendComposeDescriptionNote(description: string | undefined, note: string | undefined): string | undefined {
  const trimmedDescription = description?.trim();
  const trimmedNote = note?.trim();
  if (!trimmedNote) {
    return trimmedDescription;
  }
  if (!trimmedDescription) {
    return trimmedNote;
  }
  if (trimmedDescription.includes(trimmedNote)) {
    return trimmedDescription;
  }
  return `${trimmedDescription} ${trimmedNote}`;
}

function shouldSuppressComposeDescription(preset: FormPresetId): boolean {
  return preset === 'email-compose';
}

function resolveComposePreset(
  requestedPreset: FormPresetId,
  title: string,
  description: string | undefined,
  staticArgs: Record<string, unknown>,
  prefill: Record<string, string>
): IResolvedComposePreset {
  let preset = requestedPreset;

  if (requestedPreset === 'teams-message' && !hasNonEmptyStaticArg(staticArgs, 'chatId')) {
    preset = 'share-teams-chat';
  } else if (
    requestedPreset === 'teams-channel-message'
    && (!hasNonEmptyStaticArg(staticArgs, 'teamId') || !hasNonEmptyStaticArg(staticArgs, 'channelId'))
  ) {
    preset = 'share-teams-channel';
  } else if (looksLikeWordDocumentCreateRequest(
    requestedPreset,
    title,
    description,
    staticArgs,
    prefill,
    [getLatestUserTranscriptText(), getLatestUserUtterance()].filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
  )) {
    preset = 'word-document-create';
  }

  const presetConfig = getFormPreset(preset);
  return {
    preset,
    prefill: normalizeComposePrefillAliases(prefill, presetConfig.fields)
  };
}

function resolveComposePrefillTarget(
  preset: FormPresetId,
  formFields: IFormFieldDefinition[]
): IComposeContentPrefillTarget | undefined {
  const mappedTarget = COMPOSE_CONTENT_PREFILL_TARGETS[preset];
  if (!mappedTarget) {
    return undefined;
  }

  const fieldKeys = new Set(formFields.map((field) => field.key));
  return {
    subjectKey: mappedTarget.subjectKey && fieldKeys.has(mappedTarget.subjectKey) ? mappedTarget.subjectKey : undefined,
    contentKey: mappedTarget.contentKey && fieldKeys.has(mappedTarget.contentKey) ? mappedTarget.contentKey : undefined,
    allowAttachments: mappedTarget.allowAttachments
  };
}

function isThinSharePrefill(value: string | undefined): boolean {
  const trimmed = value?.trim();
  if (!trimmed) {
    return true;
  }
  if (trimmed.length <= 80) {
    return true;
  }
  return THIN_SHARE_PREFILL_PATTERNS.some((pattern) => pattern.test(trimmed));
}

function isVisibleContextReferencePrefill(value: string): boolean {
  const normalized = value.toLowerCase();
  const hasPanelReference = normalized.includes('action panel')
    || normalized.includes('search results')
    || normalized.includes('results panel')
    || normalized.includes('visible above')
    || normalized.includes('listed below');
  const hasEnumeratedList = /^\s*\d+[.)]\s+/m.test(value);
  const hasCollectionWord = /\b(results?|documents?|files?|folders?|items?|entries|events?|emails?|messages?|records?)\b/i.test(value);
  const hasShareInstruction = /\b(attach|include|review|send|share)\b/i.test(value);

  return (hasPanelReference && hasCollectionWord) || (hasEnumeratedList && hasCollectionWord && hasShareInstruction);
}

function mergeVisibleContextText(existingValue: string | undefined, detailedPlainText: string): string | undefined {
  const visibleContent = detailedPlainText.trim();
  if (!visibleContent) {
    return existingValue;
  }

  const trimmedExisting = existingValue?.trim();
  if (!trimmedExisting) {
    return visibleContent;
  }
  if (trimmedExisting.includes(visibleContent)) {
    return trimmedExisting;
  }
  if (isThinSharePrefill(trimmedExisting) || isVisibleContextReferencePrefill(trimmedExisting)) {
    return `${trimmedExisting}\n\n${visibleContent}`.trim();
  }
  return trimmedExisting;
}

function mergeComposeVisibleContent(
  preset: FormPresetId,
  existingValue: string | undefined,
  visibleContent: string | undefined
): string | undefined {
  const trimmedVisibleContent = visibleContent?.trim();
  if (!trimmedVisibleContent) {
    return existingValue;
  }

  if (
    (preset === 'file-create' || preset === 'word-document-create')
    && existingValue
    && (isThinSharePrefill(existingValue) || isVisibleContextReferencePrefill(existingValue))
  ) {
    return trimmedVisibleContent;
  }

  return mergeVisibleContextText(existingValue, trimmedVisibleContent);
}

function sanitizeComposeFilenamePrefill(
  preset: FormPresetId,
  prefill: Record<string, string>
): Record<string, string> {
  if (preset !== 'file-create' && preset !== 'word-document-create') {
    return prefill;
  }

  const extension = preset === 'word-document-create' ? 'docx' : 'txt';
  const fileName = prefill.filename?.trim();
  if (!fileName) {
    return prefill;
  }

  return {
    ...prefill,
    filename: sanitizeCreatedFileName(fileName, extension)
  };
}

function mergeAttachmentUris(
  staticArgs: Record<string, unknown>,
  attachmentUris: string[]
): Record<string, unknown> {
  if (attachmentUris.length === 0) {
    return staticArgs;
  }

  const existingValues = Array.isArray(staticArgs.attachmentUris)
    ? staticArgs.attachmentUris.filter((value): value is string => typeof value === 'string')
    : [];

  return {
    ...staticArgs,
    attachmentUris: dedupeStringList([...existingValues, ...attachmentUris])
  };
}

function parsePrefillAttachmentTitles(prefill: Record<string, string>): string[] {
  const raw = prefill.attachments?.trim();
  if (!raw) {
    return [];
  }

  const primaryParts = raw
    .split(/\s*;\s*|\n+/)
    .map((value) => value.trim())
    .filter((value) => value.length > 0);
  if (primaryParts.length > 1) {
    return dedupeStringList(primaryParts);
  }

  const commaParts = raw
    .split(/\s*,\s*/)
    .map((value) => value.trim())
    .filter((value) => value.length > 0);
  if (commaParts.length > 1) {
    return dedupeStringList(commaParts);
  }

  return primaryParts.length > 0 ? primaryParts : [raw];
}

function resolveAttachmentUrisFromPrefillTitles(prefill: Record<string, string>): string[] {
  const titles = parsePrefillAttachmentTitles(prefill);
  if (titles.length === 0) {
    return [];
  }

  const resolved = titles
    .map((title) => {
      const placeholder = `https://placeholder.invalid/${encodeURIComponent(title)}`;
      return resolveUrlFromBlocks(placeholder, title);
    })
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0);

  return dedupeStringList(resolved);
}

function normalizeSelectionLookupValue(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/\.[a-z0-9]{2,8}$/i, '')
    .replace(/[^a-z0-9]+/gi, ' ');
}

function deriveShareSelectionIndicesFromAttachments(
  blocks: IBlock[],
  activeBlockId: string | undefined,
  staticArgs: Record<string, unknown>
): number[] | undefined {
  if (!activeBlockId) {
    return undefined;
  }

  const attachmentUris = readStaticArgStringList(staticArgs, 'attachmentUris');
  if (attachmentUris.length === 0) {
    return undefined;
  }

  const activeBlock = blocks.find((block) => block.id === activeBlockId);
  if (!activeBlock) {
    return undefined;
  }

  const attachmentUrlSet = new Set(attachmentUris.map((uri) => uri.trim()));
  const attachmentTitleKeys = new Set(
    attachmentUris
      .map((uri) => getAttachmentDisplayName(uri))
      .filter((value) => value.length > 0)
      .map(normalizeSelectionLookupValue)
  );

  const matchedIndices: number[] = [];
  const pushIndexIfMatch = (index: number, url: string | undefined, title: string | undefined): void => {
    const normalizedUrl = typeof url === 'string' ? url.trim() : '';
    const normalizedTitle = typeof title === 'string' ? normalizeSelectionLookupValue(title) : '';
    if (
      (normalizedUrl && attachmentUrlSet.has(normalizedUrl))
      || (normalizedTitle && attachmentTitleKeys.has(normalizedTitle))
    ) {
      matchedIndices.push(index);
    }
  };

  if (activeBlock.type === 'search-results') {
    const data = activeBlock.data as ISearchResultsData;
    data.results.forEach((result, index) => {
      pushIndexIfMatch(index + 1, result.url, result.title);
    });
  } else if (activeBlock.type === 'document-library') {
    const data = activeBlock.data as IDocumentLibraryData;
    data.items.forEach((item, index) => {
      pushIndexIfMatch(index + 1, item.url, item.name);
    });
  }

  return matchedIndices.length > 0 ? matchedIndices : undefined;
}

function readComposeShareSelectionIndices(staticArgs: Record<string, unknown>): number[] | undefined {
  const raw = staticArgs.shareSelectionIndices;
  if (!Array.isArray(raw)) {
    return undefined;
  }

  const parsed = raw
    .map((value) => typeof value === 'number' ? value : parseInt(String(value), 10))
    .filter((value) => Number.isFinite(value) && value > 0);
  return parsed.length > 0 ? parsed : undefined;
}

function hasExplicitComposeShareHints(staticArgs: Record<string, unknown>): boolean {
  return typeof staticArgs.shareBlockId === 'string'
    || typeof staticArgs.shareScopeMode === 'string'
    || typeof staticArgs.shareItemTitle === 'string'
    || (Array.isArray(staticArgs.shareSelectionIndices) && staticArgs.shareSelectionIndices.length > 0);
}

function resolveComposeShareBlockId(
  blocks: IBlock[],
  rawShareBlockId: unknown,
  preferAttachableFallback: boolean
): string | undefined {
  const trimmedShareBlockId = typeof rawShareBlockId === 'string' ? rawShareBlockId.trim() : '';
  if (trimmedShareBlockId) {
    const exactMatch = blocks.find((block) => block.id === trimmedShareBlockId);
    if (exactMatch) {
      return exactMatch.id;
    }

    if (trimmedShareBlockId.startsWith('search_')) {
      const latestSearchResultsBlock = blocks
        .slice()
        .reverse()
        .find((block) => block.type === 'search-results');
      if (latestSearchResultsBlock) {
        return latestSearchResultsBlock.id;
      }
    }
  }

  if (!preferAttachableFallback) {
    return undefined;
  }

  return blocks
    .slice()
    .reverse()
    .find((block) => block.type === 'search-results' || block.type === 'document-library' || block.type === 'file-preview')
    ?.id;
}

function sanitizeComposeStaticArgs(staticArgs: Record<string, unknown>): Record<string, unknown> {
  const nextStaticArgs = { ...staticArgs };
  delete nextStaticArgs.shareSelectionIndices;
  delete nextStaticArgs.shareBlockId;
  delete nextStaticArgs.shareItemTitle;
  delete nextStaticArgs.shareScopeMode;
  delete nextStaticArgs.shareScopeExplicit;
  delete nextStaticArgs.shareScopeResolved;
  delete nextStaticArgs.fileOrFolderUrl;
  delete nextStaticArgs.fileOrFolderName;
  delete nextStaticArgs.skipSessionHydration;
  return nextStaticArgs;
}

function getAttachmentDisplayName(uri: string): string {
  const trimmed = uri.trim();
  if (!trimmed) {
    return '';
  }

  try {
    const parsed = new URL(trimmed);
    const pathSegments = parsed.pathname.split('/').filter(Boolean);
    return decodeURIComponent(pathSegments[pathSegments.length - 1] || trimmed);
  } catch {
    return trimmed;
  }
}

function buildComposeAttachmentDescriptionNote(
  staticArgs: Record<string, unknown>,
  attachmentUris: string[]
): string | undefined {
  if (attachmentUris.length === 0) {
    return undefined;
  }

  const explicitTitle = typeof staticArgs.shareItemTitle === 'string'
    ? staticArgs.shareItemTitle.trim()
    : '';
  if (attachmentUris.length === 1) {
    const displayName = getAttachmentDisplayName(attachmentUris[0]) || explicitTitle || 'Selected file';
    return `Attachment prepared: ${displayName}`;
  }

  const names = attachmentUris
    .map(getAttachmentDisplayName)
    .filter((value) => value.length > 0);
  if (names.length === 0) {
    return `Attachments prepared: ${attachmentUris.length} files`;
  }

  const preview = names.slice(0, 3).join(', ');
  return `Attachments prepared: ${preview}${names.length > 3 ? `, +${names.length - 3} more` : ''}`;
}

function resolveComposeSubmissionTargetContext(
  store: IFunctionCallStore,
  sourceContext: IToolRuntimeHandlerDeps['sourceContext'] | undefined,
  existingTargetContext: IFormSubmissionTarget['targetContext'],
  staticArgs: Record<string, unknown>,
  prefill: Record<string, string>
): IFormSubmissionTarget['targetContext'] {
  const explicitTargetContext = mergeMcpTargetContexts(
    existingTargetContext,
    deriveMcpTargetContextFromUnknown(staticArgs, 'explicit-user'),
    deriveMcpTargetContextFromUnknown(prefill, 'explicit-user')
  );

  return resolveEffectiveMcpTargetContext({
    explicitTargetContext,
    sourceContext,
    taskContext: hybridInteractionEngine.getCurrentTaskContext(),
    artifacts: hybridInteractionEngine.getCurrentArtifacts(),
    currentSiteUrl: store.userContext?.currentSiteUrl
  }).targetContext;
}

function getPositiveSelectedIndices(
  selectedItems?: Array<{ index?: number }>
): number[] | undefined {
  if (!selectedItems || selectedItems.length === 0) {
    return undefined;
  }

  const indices = Array.from(new Set(
    selectedItems
      .map((item) => item.index)
      .filter((index): index is number => typeof index === 'number' && index > 0)
  ));

  return indices.length > 0 ? indices : undefined;
}

function getLatestUserTranscriptText(): string | undefined {
  const transcript = useGrimoireStore.getState().transcript;
  for (let i = transcript.length - 1; i >= 0; i--) {
    const entry = transcript[i];
    if (entry.role !== 'user') {
      continue;
    }

    const trimmed = entry.text.trim();
    if (trimmed) {
      return trimmed;
    }
  }

  return undefined;
}

function enrichComposePrefillFromSession(
  preset: FormPresetId,
  title: string,
  formFields: IFormFieldDefinition[],
  prefill: Record<string, string>,
  staticArgs: Record<string, unknown>
): {
  prefill: Record<string, string>;
  staticArgs: Record<string, unknown>;
  descriptionNote?: string;
  targetContext?: IFormSubmissionTarget['targetContext'];
} {
  const target = resolveComposePrefillTarget(preset, formFields);
  if (!target) {
    return { prefill, staticArgs };
  }

  const skipSessionHydration = staticArgs.skipSessionHydration === true || staticArgs.skipSessionHydration === 'true';
  if (skipSessionHydration) {
    return {
      prefill: { ...prefill },
      staticArgs: sanitizeComposeStaticArgs(staticArgs),
      targetContext: deriveMcpTargetContextFromUnknown(staticArgs, 'explicit-user')
    };
  }

  const state = useGrimoireStore.getState();
  if (!hasShareableSessionContent(state.blocks, state.transcript)) {
    return { prefill, staticArgs };
  }

  const hieSourceContext = hybridInteractionEngine.captureCurrentSourceContext();
  const hieSelectedIndices = getPositiveSelectedIndices(hieSourceContext?.selectedItems);
  const fallbackSelectedIndices = hieSelectedIndices || state.selectedActionIndices;
  const activeBlockIdFromArgs = resolveComposeShareBlockId(
    state.blocks,
    staticArgs.shareBlockId,
    target.allowAttachments === true && hasExplicitComposeShareHints(staticArgs)
  );
  const normalizedStaticArgs = activeBlockIdFromArgs
    ? { ...staticArgs, shareBlockId: activeBlockIdFromArgs }
    : staticArgs;
  const baseActiveBlockId = activeBlockIdFromArgs || hieSourceContext?.sourceBlockId || state.activeActionBlockId;
  const explicitShareScope = readComposeShareScope(normalizedStaticArgs);
  const inferredShareScope = explicitShareScope
    ? undefined
    : resolveContextualComposeScope({
      text: getLatestUserTranscriptText(),
      blocks: state.blocks,
      activeBlockId: baseActiveBlockId,
      selectedActionIndices: fallbackSelectedIndices
    });
  const effectiveShareScope = explicitShareScope || inferredShareScope;
  const effectiveActiveBlockId = effectiveShareScope?.blockId || baseActiveBlockId;
  const explicitSelectedIndices = effectiveShareScope?.explicit
    ? effectiveShareScope.selectedIndices
      : (readComposeShareSelectionIndices(staticArgs)
        || deriveShareSelectionIndicesFromAttachments(state.blocks, effectiveActiveBlockId, staticArgs));

  const shareContent = new SessionShareFormatter().format({
    blocks: state.blocks,
    transcript: state.transcript,
    activeBlockId: effectiveActiveBlockId,
    selectedActionIndices: effectiveShareScope?.explicit
      ? explicitSelectedIndices
      : (explicitSelectedIndices || fallbackSelectedIndices),
    selectionBehavior: effectiveShareScope?.explicit ? 'strict' : 'default'
  });
  const preferredFocusedBlockId = (
    typeof hieSourceContext?.sourceArtifactId === 'string'
    && state.blocks.some((block) => block.id === hieSourceContext.sourceArtifactId)
  )
    ? hieSourceContext.sourceArtifactId
    : effectiveActiveBlockId;
  const focusedBlockContent = extractFocusedBlockContent(state.blocks, preferredFocusedBlockId, pickResolvedString)
    || extractFocusedBlockContent(state.blocks, effectiveActiveBlockId, pickResolvedString);
  const focusedBlockTitle = state.blocks.find((block) => block.id === preferredFocusedBlockId)?.title
    || state.blocks.find((block) => block.id === effectiveActiveBlockId)?.title;

  const nextPrefill = { ...prefill };
  if (target.subjectKey && !nextPrefill[target.subjectKey]?.trim()) {
    nextPrefill[target.subjectKey] = shareContent.subject;
  }
  if (!nextPrefill.filename?.trim() && (preset === 'file-create' || preset === 'word-document-create')) {
    const contextualFileName = buildContextualFileName(preset, title, focusedBlockTitle);
    if (contextualFileName) {
      nextPrefill.filename = contextualFileName;
    }
  }
  if (target.contentKey) {
    const mergedContent = mergeComposeVisibleContent(
      preset,
      nextPrefill[target.contentKey],
      focusedBlockContent || shareContent.detailedPlainText
    );
    if (mergedContent) {
      nextPrefill[target.contentKey] = mergedContent;
    }
  }

  const nextStaticArgs = sanitizeComposeStaticArgs(
    target.allowAttachments
      ? mergeAttachmentUris(normalizedStaticArgs, shareContent.attachmentUris)
      : normalizedStaticArgs
  );
  const voiceResolvedAttachmentUris = target.allowAttachments
    ? resolveAttachmentUrisFromPrefillTitles(nextPrefill)
    : [];
  const nextStaticArgsWithResolvedPrefillAttachments = target.allowAttachments
    ? mergeAttachmentUris(nextStaticArgs, voiceResolvedAttachmentUris)
    : nextStaticArgs;
  const attachmentUris = Array.isArray(nextStaticArgsWithResolvedPrefillAttachments.attachmentUris)
    ? nextStaticArgsWithResolvedPrefillAttachments.attachmentUris.filter((value): value is string => typeof value === 'string')
    : [];

  return {
    prefill: nextPrefill,
    staticArgs: nextStaticArgsWithResolvedPrefillAttachments,
    descriptionNote: target.allowAttachments
      ? buildComposeAttachmentDescriptionNote(normalizedStaticArgs, attachmentUris)
      : undefined,
    targetContext: deriveMcpTargetContextFromUnknown(normalizedStaticArgs, 'explicit-user')
  };
}

function buildTeamsChannelContentField(formFields: IFormFieldDefinition[]): IFormFieldDefinition {
  return formFields.find((field) => field.key === 'content') || {
    key: 'content',
    label: 'Message',
    type: 'textarea' as const,
    required: true,
    placeholder: 'Review the share message...',
    rows: 8,
    group: 'Message'
  };
}

function buildTeamsChannelManualFields(formFields: IFormFieldDefinition[]): IFormFieldDefinition[] {
  return [
    { key: 'teamName', label: 'Team name', type: 'text', required: true, placeholder: 'Marketing', group: 'Destination' },
    { key: 'teamId', label: 'Team ID', type: 'hidden', required: false, group: 'Destination' },
    { key: 'channelName', label: 'Channel name', type: 'text', required: true, placeholder: 'General', group: 'Destination' },
    { key: 'channelId', label: 'Channel ID', type: 'hidden', required: false, group: 'Destination' },
    buildTeamsChannelContentField(formFields)
  ];
}

function buildTeamsChannelDropdownFields(
  formFields: IFormFieldDefinition[],
  options: Array<{ key: string; text: string }>
): IFormFieldDefinition[] {
  return [
    {
      key: 'destination',
      label: 'Channel',
      type: 'dropdown',
      required: true,
      placeholder: 'Select a team and channel',
      group: 'Destination',
      options
    },
    { key: 'teamId', label: 'Team ID', type: 'hidden', required: false, group: 'Destination' },
    { key: 'teamName', label: 'Team name', type: 'hidden', required: false, group: 'Destination' },
    { key: 'channelId', label: 'Channel ID', type: 'hidden', required: false, group: 'Destination' },
    { key: 'channelName', label: 'Channel name', type: 'hidden', required: false, group: 'Destination' },
    buildTeamsChannelContentField(formFields)
  ];
}

async function enhanceComposeFieldsForSharePreset(
  preset: FormPresetId,
  formFields: IFormFieldDefinition[],
  store: IFunctionCallStore
): Promise<IComposeFieldEnhancementResult> {
  if (preset !== 'share-teams-channel') {
    return { fields: formFields };
  }

  if (isPnpInitialized()) {
    return { fields: formFields };
  }

  try {
    const options = await loadTeamsChannelDestinationOptions(store);
    if (options.length === 0) {
      logService.warning('mcp', 'Teams channel picker returned no options; falling back to manual entry fields.');
      return {
        fields: buildTeamsChannelManualFields(formFields),
        descriptionNote: TEAMS_CHANNEL_PICKER_FALLBACK_NOTE
      };
    }

    return {
      fields: buildTeamsChannelDropdownFields(formFields, options)
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    logService.warning('mcp', `Could not build Teams channel picker: ${message}`);
    return {
      fields: buildTeamsChannelManualFields(formFields),
      descriptionNote: TEAMS_CHANNEL_PICKER_FALLBACK_NOTE
    };
  }
}

function getLatestUserUtterance(): string {
  const transcript = useGrimoireStore.getState().transcript;
  for (let i = transcript.length - 1; i >= 0; i--) {
    if (transcript[i].role === 'user') return transcript[i].text;
  }
  return '';
}

function extractFileExtensionFromInsightsUrl(url: string | undefined): string | undefined {
  if (!url) {
    return undefined;
  }

  try {
    const parsed = new URL(url);
    const fileParam = parsed.searchParams.get('file');
    if (fileParam) {
      const fileMatch = fileParam.match(/\.([a-z0-9]+)$/i);
      if (fileMatch) {
        return fileMatch[1].toLowerCase();
      }
    }

    const pathname = parsed.pathname || '';
    const match = pathname.match(/\.([a-z0-9]+)$/i);
    if (match) {
      return match[1].toLowerCase();
    }
  } catch {
    return undefined;
  }

  return undefined;
}

const BLOCKED_INSIGHTS_KINDS: ReadonlySet<string> = new Set([
  'web',
  'webpage',
  'site',
  'sitepage',
  'spsite',
  'sharepointsite',
  'aspx',
  'page',
  'link'
]);

const ALLOWED_INSIGHTS_KINDS: ReadonlySet<string> = new Set([
  'doc',
  'docx',
  'document',
  'word',
  'xls',
  'xlsx',
  'excel',
  'ppt',
  'pptx',
  'powerpoint',
  'pdf',
  'txt',
  'text',
  'rtf',
  'csv',
  'zip',
  'archive',
  'onenote',
  'one',
  'onepkg',
  'image',
  'jpg',
  'jpeg',
  'png',
  'gif',
  'webp',
  'bmp',
  'svg',
  'audio',
  'mp3',
  'wav',
  'video',
  'mp4',
  'mov',
  'avi',
  'mkv'
]);

function normalizeInsightsKind(value: string | undefined): string {
  return (value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function looksLikeInsightsDocument(item: Record<string, unknown>, result: ISearchResult): boolean {
  const viz = item.resourceVisualization as Record<string, unknown> | undefined;
  const ref = item.resourceReference as Record<string, unknown> | undefined;
  const normalizedKinds = [
    normalizeInsightsKind(result.fileType),
    normalizeInsightsKind(typeof viz?.type === 'string' ? viz.type : undefined),
    normalizeInsightsKind(typeof ref?.type === 'string' ? ref.type : undefined)
  ].filter((kind) => kind.length > 0);

  if (normalizedKinds.some((kind) => BLOCKED_INSIGHTS_KINDS.has(kind))) {
    return false;
  }

  const extension = extractFileExtensionFromInsightsUrl(result.url)
    || (() => {
      const titleMatch = (result.title || '').match(/\.([a-z0-9]+)$/i);
      return titleMatch ? titleMatch[1].toLowerCase() : undefined;
    })();

  if (extension) {
    return !BLOCKED_INSIGHTS_KINDS.has(extension);
  }

  return normalizedKinds.some((kind) => ALLOWED_INSIGHTS_KINDS.has(kind));
}

function fetchInsightsDocuments(
  toolName: string,
  label: string,
  endpoint: string,
  mapItem: (item: Record<string, unknown>) => ISearchResult,
  deps: IToolRuntimeHandlerDeps
): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> {
  const { aadClient, store, awaitAsync } = deps;
  if (!aadClient) {
    return errorOutcome(JSON.stringify({ success: false, error: 'Graph API not available' }));
  }

  const graphSvc = new GraphService(aadClient);
  const block = createBlock('search-results', label, {
    kind: 'search-results', query: label.toLowerCase(), results: [], totalCount: 0, source: 'pending'
  } as ISearchResultsData);
  trackCreatedBlock(store, block, deps);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const promise = graphSvc.get<{ value?: any[] }>(endpoint)
    .then((result) => {
      if (!result.success || !result.data?.value) {
        store.updateBlock(block.id, {
          data: { kind: 'search-results', query: label.toLowerCase(), results: [], totalCount: 0, source: 'insights' } as ISearchResultsData
        });
        trackToolCompletion(toolName, block.id, false, 0, deps);
        return errorOutcome(JSON.stringify({ success: false, error: result.error || 'Insights API returned no data. This may be disabled for this tenant.' }));
      }
      const mappedEntries = result.data.value.map((item) => ({
        raw: item as Record<string, unknown>,
        mapped: mapItem(item as Record<string, unknown>)
      }));
      const results: ISearchResult[] = mappedEntries
        .filter((entry) => looksLikeInsightsDocument(entry.raw, entry.mapped))
        .map((entry) => entry.mapped);
      const droppedKinds = mappedEntries
        .filter((entry) => !looksLikeInsightsDocument(entry.raw, entry.mapped))
        .map((entry) => entry.mapped.fileType || 'unknown');
      logService.info(
        'graph',
        `${toolName}: filtered Insights items`,
        JSON.stringify({
          rawCount: mappedEntries.length,
          keptCount: results.length,
          droppedCount: mappedEntries.length - results.length,
          droppedKinds: Array.from(new Set(droppedKinds)).slice(0, 8)
        })
      );
      const data = { kind: 'search-results', query: label, results, totalCount: results.length, source: 'insights' } as ISearchResultsData;
      trackUpdatedBlock(store, block.id, { data }, { ...block, data }, deps);
      trackToolCompletion(toolName, block.id, true, results.length, deps);
      return completeOutcome(JSON.stringify({ success: true, count: results.length, message: `Found ${results.length} ${label.toLowerCase()}.` }));
    })
    .catch((err: Error) => {
      store.updateBlock(block.id, {
        data: { kind: 'search-results', query: label.toLowerCase(), results: [], totalCount: 0, source: 'insights' } as ISearchResultsData
      });
      trackToolCompletion(toolName, block.id, false, 0, deps);
      logService.error('graph', `${toolName} failed: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message || 'Insights request failed' }));
    });
  if (awaitAsync) return promise;
  return completeOutcome(JSON.stringify({ success: true, message: `Loading ${label.toLowerCase()}...` }));
}

function getSchemaPropertyNames(inputSchema: Record<string, unknown>): string[] {
  const properties = (inputSchema as { properties?: Record<string, unknown> }).properties || {};
  return Object.keys(properties);
}

function getSchemaRequiredNames(inputSchema: Record<string, unknown>): string[] {
  const required = (inputSchema as { required?: string[] }).required;
  return Array.isArray(required) ? required : [];
}

function getMissingRequiredSchemaArgs(
  inputSchema: Record<string, unknown>,
  args: Record<string, unknown>
): string[] {
  return getSchemaRequiredNames(inputSchema).filter((name) => {
    const value = args[name];
    return value === undefined || value === null || value === '';
  });
}

function buildPermissionLookupArgs(
  inputSchema: Record<string, unknown>,
  targetUrl: string,
  targetName: string
): Record<string, unknown> {
  const args: Record<string, unknown> = {};
  const propertyNames = getSchemaPropertyNames(inputSchema);
  propertyNames.forEach((propertyName) => {
    const lowered = propertyName.toLowerCase();
    if ((lowered.includes('url') || lowered.includes('link')) && !args[propertyName]) {
      args[propertyName] = targetUrl;
      return;
    }
    if ((lowered.includes('name') || lowered.includes('title')) && targetName && !args[propertyName]) {
      args[propertyName] = targetName;
    }
  });
  return args;
}

function tryParseMcpJsonText(text: string): unknown | undefined {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function extractMcpJsonObject(content: Array<{ type: string; text?: string }>): Record<string, unknown> | undefined {
  const candidates: string[] = [];
  const { reply, raw } = extractMcpReply(content);
  if (reply) {
    candidates.push(reply);
  }
  if (raw) {
    candidates.push(raw);
  }
  content.forEach((item) => {
    if (item.type === 'text' && item.text) {
      candidates.push(item.text);
    }
  });

  for (let i = 0; i < candidates.length; i++) {
    const parsed = tryParseMcpJsonText(candidates[i]);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      continue;
    }
    const obj = parsed as Record<string, unknown>;
    const data = obj.data;
    if (data && typeof data === 'object' && !Array.isArray(data)) {
      return data as Record<string, unknown>;
    }
    return obj;
  }

  return undefined;
}

function pickResolvedString(...values: unknown[]): string | undefined {
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

function fillPermissionArgsFromMetadata(
  args: Record<string, unknown>,
  inputSchema: Record<string, unknown>,
  metadata: Record<string, unknown>,
  targetUrl: string
): Record<string, unknown> {
  const nextArgs = { ...args };
  const parentReference = metadata.parentReference as Record<string, unknown> | undefined;
  const driveId = pickResolvedString(
    parentReference?.driveId,
    metadata.driveId,
    metadata.documentLibraryId
  );
  const itemId = pickResolvedString(
    metadata.id,
    metadata.itemId,
    metadata.fileId,
    metadata.fileOrFolderId
  );
  const siteId = pickResolvedString(
    parentReference?.siteId,
    metadata.siteId
  );

  getSchemaPropertyNames(inputSchema).forEach((propertyName) => {
    if (nextArgs[propertyName] !== undefined && nextArgs[propertyName] !== null && nextArgs[propertyName] !== '') {
      return;
    }

    const lowered = propertyName.toLowerCase();
    if ((lowered.includes('url') || lowered.includes('link')) && targetUrl) {
      nextArgs[propertyName] = targetUrl;
      return;
    }
    if (lowered.includes('siteid') && siteId) {
      nextArgs[propertyName] = siteId;
      return;
    }
    if (
      (lowered.includes('documentlibraryid') || lowered === 'driveid' || lowered.includes('driveid'))
      && driveId
    ) {
      nextArgs[propertyName] = driveId;
      return;
    }
    if (
      lowered.includes('fileorfolderid')
      || lowered === 'fileid'
      || lowered === 'itemid'
      || lowered === 'folderid'
      || lowered.includes('driveitemid')
    ) {
      if (itemId) {
        nextArgs[propertyName] = itemId;
      }
    }
  });

  return nextArgs;
}

async function resolvePermissionToolArgs(
  mcpClient: McpClientService,
  permissionTool: IMcpToolDescriptor,
  availableTools: IMcpToolDescriptor[],
  targetUrl: string,
  targetName: string,
  serverUrl: string,
  serverName: string,
  store: IFunctionCallStore,
  deps: IToolRuntimeHandlerDeps
): Promise<Record<string, unknown>> {
  const baseArgs = buildPermissionLookupArgs(permissionTool.inputSchema, targetUrl, targetName);
  const missingRequired = getMissingRequiredSchemaArgs(permissionTool.inputSchema, baseArgs);
  if (missingRequired.length === 0) {
    return baseArgs;
  }

  const metadataTool = availableTools.find((tool) => tool.name.toLowerCase() === 'getfileorfoldermetadatabyurl');
  if (!metadataTool) {
    return baseArgs;
  }

  const metadataArgs = buildPermissionLookupArgs(metadataTool.inputSchema, targetUrl, targetName);
  const metadataExecution = await executeCatalogMcpTool({
    serverId: 'mcp_ODSPRemoteServer',
    serverName,
    serverUrl,
    toolName: metadataTool.name,
    rawArgs: metadataArgs,
    connections: useGrimoireStore.getState().mcpConnections,
    getConnections: () => useGrimoireStore.getState().mcpConnections,
    mcpClient,
    sessionHelpers: {
      findExistingSession,
      connectToM365Server
    },
    getToken: store.getToken,
    explicitTargetContext: deriveMcpTargetContextFromUnknown({
      url: targetUrl,
      title: targetName
    }, 'explicit-user'),
    sourceContext: deps.sourceContext,
    taskContext: hybridInteractionEngine.getCurrentTaskContext(),
    artifacts: hybridInteractionEngine.getCurrentArtifacts(),
    currentSiteUrl: store.userContext?.currentSiteUrl
  });
  if (!metadataExecution.success || !metadataExecution.mcpResult) {
    logService.warning('mcp', `Metadata resolution for permissions failed: ${metadataExecution.error}`);
    return baseArgs;
  }

  const metadata = extractMcpJsonObject(metadataExecution.mcpResult.content);
  if (!metadata) {
    return baseArgs;
  }

  return fillPermissionArgsFromMetadata(baseArgs, permissionTool.inputSchema, metadata, targetUrl);
}

function pushPermissionsErrorBlock(
  deps: IToolRuntimeHandlerDeps,
  store: IFunctionCallStore,
  targetName: string,
  message: string,
  detail?: string,
  unsupported?: boolean
): ToolRuntimeHandlerResult {
  const errorBlock = createBlock('error', `Permissions: ${targetName}`, {
    kind: 'error',
    message,
    detail
  } as IErrorData);
  trackCreatedBlock(store, errorBlock, deps);
  trackToolCompletion('show_permissions', errorBlock.id, false, 0, deps);
  return errorOutcome(JSON.stringify({
    success: false,
    unsupported: unsupported === true,
    targetName,
    error: message,
    detail
  }));
}

function buildPermissionUnsupportedMessage(targetName: string): string {
  const label = targetName.trim() ? `"${targetName}"` : 'the selected item';
  return `I couldn't inspect permissions for ${label} because the connected SharePoint & OneDrive MCP server does not expose a permission-inspection tool.`;
}

export function buildUiAndPersonalRuntimeHandlers(
): Pick<Record<RuntimeHandledToolName, ToolRuntimeHandler>, UiPersonalRuntimeToolName> {
  return {
    show_info_card: (args, deps): ToolRuntimeHandlerResult => {
      const rawIcon = typeof args.icon === 'string' ? args.icon.trim() : '';
      const icon = rawIcon.toLowerCase() === 'book' ? undefined : (rawIcon || undefined);
      const data: IInfoCardData = {
        kind: 'info-card',
        heading: args.heading as string,
        body: args.body as string,
        icon
      };
      const infoBlock = createBlock('info-card', args.heading as string, data);
      trackCreatedBlock(deps.store, infoBlock, deps);
      return completeOutcome(JSON.stringify({ success: true }));
    },

    show_markdown: (args, deps): ToolRuntimeHandlerResult => {
      const mdBlock = createBlock('markdown', args.title as string, {
        kind: 'markdown',
        content: args.content as string
      } as IMarkdownData);
      trackCreatedBlock(deps.store, mdBlock, deps);
      return completeOutcome(JSON.stringify({ success: true }));
    },

    ask_confirmation: (args, deps): ToolRuntimeHandlerResult => {
      const data: IConfirmationDialogData = {
        kind: 'confirmation-dialog',
        message: args.message as string,
        confirmLabel: (args.confirm_label as string) || 'Confirm',
        cancelLabel: (args.cancel_label as string) || 'Cancel',
        onConfirmAction: ''
      };
      const confirmBlock = createBlock('confirmation-dialog', 'Confirm', data, false);
      trackCreatedBlock(deps.store, confirmBlock, deps);
      return completeOutcome(JSON.stringify({ success: true, message: 'Confirmation shown to user.' }));
    },

    clear_action_panel: (_args, deps): ToolRuntimeHandlerResult => {
      const currentBlocks = useGrimoireStore.getState().blocks;
      currentBlocks.forEach((b) => hybridInteractionEngine.onBlockRemoved(b.id));
      deps.store.clearBlocks();
      return completeOutcome(JSON.stringify({ success: true }));
    },

    set_expression: (args, deps): ToolRuntimeHandlerResult => {
      if (deps.store.avatarEnabled === false) {
        return completeOutcome(JSON.stringify({
          success: true,
          suppressed: true,
          message: 'Avatar is disabled for this user.'
        }));
      }
      deps.store.setExpression(args.expression as Expression);
      hybridInteractionEngine.onLlmExpression();
      return completeOutcome(JSON.stringify({ success: true }));
    },

    show_progress: (args, deps): ToolRuntimeHandlerResult => {
      const progress = parseInt((args.progress as string) || '0', 10);
      const data: IProgressTrackerData = {
        kind: 'progress-tracker',
        label: args.label as string,
        progress: Math.min(100, Math.max(0, progress)),
        status: progress >= 100 ? 'complete' : 'running'
      };
      const progressBlock = createBlock('progress-tracker', args.label as string, data);
      trackCreatedBlock(deps.store, progressBlock, deps);
      return completeOutcome(JSON.stringify({ success: true }));
    },

    show_error: (args, deps): ToolRuntimeHandlerResult => {
      const data: IErrorData = {
        kind: 'error',
        message: args.message as string,
        detail: args.detail as string
      };
      const errorBlock = createBlock('error', 'Error', data);
      trackCreatedBlock(deps.store, errorBlock, deps);
      trackToolCompletion('show_error', errorBlock.id, true, 0, deps);
      return completeOutcome(JSON.stringify({ success: true }));
    },

    show_selection_list: (args, deps): ToolRuntimeHandlerResult => {
      const prompt = typeof args.prompt === 'string' ? args.prompt : '';
      const latestState = useGrimoireStore.getState();
      const latestBlockType = latestState.blocks.length > 0
        ? latestState.blocks[latestState.blocks.length - 1].type
        : undefined;
      const latestUserUtterance = getLatestUserUtterance();
      if (shouldSuppressSelectionList(prompt, latestBlockType, latestUserUtterance)) {
        logService.info('llm', 'show_selection_list suppressed: actionable results already visible');
        return completeOutcome(JSON.stringify({
          success: true,
          suppressed: true,
          message: 'Current results are already selectable. Skipped extra options list.'
        }));
      }

      let items: Array<{ id: string; label: string; description?: string }> = [];
      try {
        items = JSON.parse(args.items_json as string);
      } catch (parseErr) {
        logService.warning('llm', `Selection list JSON parse error: ${(parseErr as Error).message}`);
        try {
          const raw = (args.items_json as string).trim();
          const wrapped = raw.startsWith('[') ? raw : `[${raw}]`;
          const cleaned = wrapped
            .replace(/,\s*([}\]])/g, '$1')
            .replace(/[\u201C\u201D]/g, '"')
            .replace(/[\u2018\u2019]/g, "'");
          items = JSON.parse(cleaned);
          logService.info('llm', `Selection list JSON repaired (${items.length} items)`);
        } catch {
          // no-op
        }
      }
      const data: ISelectionListData = {
        kind: 'selection-list',
        prompt,
        items: items.map((item) => ({ ...item, selected: false })),
        multiSelect: (args.multi_select as string) === 'true'
      };
      const selBlock = createBlock('selection-list', prompt, data, false);
      trackCreatedBlock(deps.store, selBlock, deps);
      return completeOutcome(JSON.stringify({ success: true, message: 'Selection list shown to user.' }));
    },

    show_chart: (args, deps): ToolRuntimeHandlerResult => {
      let labels: string[] = [];
      let values: number[] = [];
      try { labels = JSON.parse(args.labels_json as string); } catch { /* no-op */ }
      try { values = JSON.parse(args.values_json as string); } catch { /* no-op */ }
      const data: IChartData = {
        kind: 'chart',
        chartType: (args.chart_type as 'bar' | 'pie' | 'line') || 'bar',
        title: args.title as string,
        labels,
        values
      };
      const chartBlock = createBlock('chart', args.title as string, data);
      trackCreatedBlock(deps.store, chartBlock, deps);
      return completeOutcome(JSON.stringify({ success: true }));
    },

    show_permissions: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const { store, awaitAsync } = deps;
      const targetName = args.target_name as string;
      const rawTargetUrl = args.target_url as string;
      const targetUrl = resolveUrlFromBlocks(rawTargetUrl, targetName) || rawTargetUrl;
      logService.info('mcp', `Permissions: ${targetName} (${targetUrl})`);
      store.setExpression('thinking');

      if (!targetUrl) {
        return pushPermissionsErrorBlock(
          deps,
          store,
          targetName || 'Item',
          'I could not resolve a URL for this item.',
          'Select a file or site first, then try again.'
        );
      }

      const existingConnection = findConnectedServerSession(store.mcpConnections, 'mcp_ODSPRemoteServer', store.mcpEnvironmentId);
      const serverUrl = existingConnection?.serverUrl
        || (store.mcpEnvironmentId ? resolveServerUrl('mcp_ODSPRemoteServer', store.mcpEnvironmentId) : undefined);
      const serverName = 'SharePoint & OneDrive';
      const liveTools = existingConnection?.tools.map((tool) => toMcpToolDescriptor(tool));
      const permissionTool = liveTools ? findCapabilityPermissionInspectionTool(liveTools) : undefined;
      if (existingConnection && liveTools && !permissionTool) {
        return pushPermissionsErrorBlock(
          deps,
          store,
          targetName,
          buildPermissionUnsupportedMessage(targetName),
          'It currently exposes sharing actions, but not a readable permission lookup.',
          true
        );
      }

      if (!serverUrl || !store.proxyConfig || !store.mcpEnvironmentId) {
        return pushPermissionsErrorBlock(
          deps,
          store,
          targetName || 'Item',
          'Microsoft 365 MCP is not configured for permission inspection.',
          'Set the MCP Environment ID and proxy configuration before trying again.'
        );
      }

      const asyncResult = (async (): Promise<ToolRuntimeHandlerResult> => {
        const mcpClient = new McpClientService(store.proxyConfig!.proxyUrl, store.proxyConfig!.proxyApiKey);
        const connectionResult = await ensureCatalogMcpConnection({
          serverId: 'mcp_ODSPRemoteServer',
          serverName,
          serverUrl,
          connections: useGrimoireStore.getState().mcpConnections,
          getConnections: () => useGrimoireStore.getState().mcpConnections,
          mcpClient,
          sessionHelpers: {
            findExistingSession,
            connectToM365Server
          },
          getToken: store.getToken
        });
        if (!connectionResult.success || !connectionResult.sessionId || !connectionResult.connection) {
          return pushPermissionsErrorBlock(
            deps,
            store,
            targetName,
            'I could not connect to the SharePoint & OneDrive MCP server.',
            connectionResult.error || 'Failed to connect to the SharePoint & OneDrive MCP server.'
          );
        }

        const latestConnections = useGrimoireStore.getState().mcpConnections;
        const availableTools = connectionResult.connection.tools.map((tool) => toMcpToolDescriptor(tool)) || [];
        const permissionLookupTool = findCapabilityPermissionInspectionTool(availableTools);
        if (!permissionLookupTool) {
          return pushPermissionsErrorBlock(
            deps,
            store,
            targetName,
            buildPermissionUnsupportedMessage(targetName),
            'It currently exposes sharing actions, but not a readable permission lookup.',
            true
          );
        }

        const permissionArgs = await resolvePermissionToolArgs(
          mcpClient,
          permissionLookupTool,
          availableTools,
          targetUrl,
          targetName,
          serverUrl,
          serverName,
          store,
          deps
        );
        const missingArgs = getMissingRequiredSchemaArgs(permissionLookupTool.inputSchema, permissionArgs);
        if (missingArgs.length > 0) {
          return pushPermissionsErrorBlock(
            deps,
            store,
            targetName,
            'The MCP permission tool needs more identifiers than Grimoire can derive from this item.',
            `Missing required inputs: ${missingArgs.join(', ')}.`,
            true
          );
        }

        const execution = await executeCatalogMcpTool({
          serverId: 'mcp_ODSPRemoteServer',
          serverName,
          serverUrl,
          toolName: permissionLookupTool.name,
          rawArgs: permissionArgs,
          connections: latestConnections,
          getConnections: () => useGrimoireStore.getState().mcpConnections,
          mcpClient,
          sessionHelpers: {
            findExistingSession,
            connectToM365Server
          },
          getToken: store.getToken,
          explicitTargetContext: deriveMcpTargetContextFromUnknown({
            url: targetUrl,
            title: targetName
          }, 'explicit-user'),
          sourceContext: deps.sourceContext,
          taskContext: hybridInteractionEngine.getCurrentTaskContext(),
          artifacts: hybridInteractionEngine.getCurrentArtifacts(),
          currentSiteUrl: store.userContext?.currentSiteUrl
        });
        if (!execution.success || !execution.mcpResult) {
          logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));
          return pushPermissionsErrorBlock(
            deps,
            store,
            targetName,
            'The MCP permission lookup failed.',
            execution.error || 'The server returned an unknown error.'
          );
        }

        const capturedBlocks: IBlock[] = [];
        mapMcpResultToBlock(
          'mcp_ODSPRemoteServer',
          permissionLookupTool.name,
          execution.mcpResult.content,
          (block) => {
            capturedBlocks.push(block);
          }
        );

        const mappedPermissionsBlock = capturedBlocks.find((block) => block.type === 'permissions-view');
        if (!mappedPermissionsBlock) {
          return pushPermissionsErrorBlock(
            deps,
            store,
            targetName,
            'The MCP server responded, but it did not return permission data in a readable format.',
            `Tool used: ${permissionLookupTool.name}.`,
            true
          );
        }

        const mappedData = mappedPermissionsBlock.data as IPermissionsViewData;
        if (!Array.isArray(mappedData.permissions) || mappedData.permissions.length === 0) {
          return pushPermissionsErrorBlock(
            deps,
            store,
            targetName,
            'The MCP permission tool did not return any readable permission entries.',
            `Tool used: ${permissionLookupTool.name}.`,
            true
          );
        }

        logService.debug('mcp', 'MCP execution trace', JSON.stringify({
          ...execution.trace,
          finalSummary: `Resolved ${mappedData.permissions.length} permission entries for ${targetName}.`
        }));

        const permissionsBlock = createBlock('permissions-view', `Permissions: ${targetName}`, {
          kind: 'permissions-view',
          targetName,
          targetUrl,
          permissions: mappedData.permissions
        } as IPermissionsViewData);
        trackCreatedBlock(store, permissionsBlock, deps);
        trackToolCompletion('show_permissions', permissionsBlock.id, true, mappedData.permissions.length, deps);

        return completeOutcome(JSON.stringify({
          success: true,
          targetName,
          permissionCount: mappedData.permissions.length,
          toolName: permissionLookupTool.name
        }));
      })();

      if (awaitAsync) return asyncResult;
      return completeOutcome(JSON.stringify({ success: true, message: `Checking MCP permissions for ${targetName}...` }));
    },

    show_activity_feed: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const { store, sitesService, awaitAsync } = deps;
      const siteUrl = args.site_url as string;
      const maxItems = parseInt((args.max_items as string) || '20', 10);
      logService.info('graph', `Activity feed: ${siteUrl}`);
      store.setExpression('thinking');

      const feedData: IActivityFeedData = {
        kind: 'activity-feed',
        activities: []
      };
      const feedBlock = createBlock('activity-feed', 'Recent Activity', feedData);
      trackCreatedBlock(store, feedBlock, deps);

      if (!sitesService) {
        logService.warning('graph', 'AadHttpClient not available — cannot fetch activity');
        return errorOutcome(JSON.stringify({ success: false, error: 'SharePoint connection not available. Please ensure you are signed in.' }));
      }

      const asyncResult = sitesService.getActivities(siteUrl, maxItems).then((resp) => {
        const currentStore = useGrimoireStore.getState();
        if (resp.success && resp.data) {
          const updatedData: IActivityFeedData = {
            kind: 'activity-feed',
            activities: resp.data
          };
          const updatedFeedBlock = { ...feedBlock, data: updatedData };
          trackUpdatedBlock(currentStore, feedBlock.id, { data: updatedData }, updatedFeedBlock, deps);
          trackToolCompletion('show_activity_feed', feedBlock.id, true, updatedData.activities.length, deps);
          return completeOutcome(JSON.stringify({ success: true, activityCount: updatedData.activities.length }));
        }
        trackToolCompletion('show_activity_feed', feedBlock.id, false, 0, deps);
        logService.error('graph', `Activity feed failed: ${resp.error}`);
        return errorOutcome(JSON.stringify({ success: false, error: resp.error }));
      }).catch((err: Error) => {
        trackToolCompletion('show_activity_feed', feedBlock.id, false, 0, deps);
        logService.error('graph', `Activity feed error: ${err.message}`);
        return errorOutcome(JSON.stringify({ success: false, error: err.message }));
      });

      if (awaitAsync) return asyncResult;
      return completeOutcome(JSON.stringify({ success: true, message: `Loading activity feed for ${siteUrl}...` }));
    },

    list_m365_servers: (args, deps): ToolRuntimeHandlerResult => {
      const { store } = deps;
      logService.info('mcp', 'List M365 servers');
      const tools = getTools({ avatarEnabled: store.avatarEnabled });
      const hasMcpEnv = !!store.mcpEnvironmentId;
      const focus = normalizeCapabilityFocus(typeof args.focus === 'string' ? args.focus : undefined);
      const serverViews = buildCapabilityServerViews(store, focus);
      const isOverview = !focus;
      const overviewCapabilities = isOverview
        ? buildCapabilityOverviewMarkdown(tools, hasMcpEnv, serverViews)
        : undefined;
      const focusedCapabilities = focus
        ? buildFocusedCapabilityMarkdown(focus, serverViews)
        : undefined;
      const blockTitle = overviewCapabilities ? 'Grimoire Capabilities' : (focusedCapabilities?.title || 'Grimoire Capabilities');
      const blockContent = overviewCapabilities?.content || focusedCapabilities?.content || '';
      const serversBlock = createBlock('markdown', blockTitle, {
        kind: 'markdown', content: blockContent
      } as IMarkdownData);
      trackCreatedBlock(store, serversBlock, deps);
      return completeOutcome(JSON.stringify({
        success: true,
        focus: focus || undefined,
        totalBuiltInToolCount: tools.length,
        userFacingToolCount: overviewCapabilities?.userFacingCount,
        internalToolCount: overviewCapabilities?.internalToolCount,
        focusedToolCount: focusedCapabilities?.focusedToolCount,
        m365Available: hasMcpEnv,
        message: isOverview
          ? 'The action panel shows a plain-language capability overview. Briefly tell the user what you can help with — do NOT read the whole list aloud unless they ask.'
          : `The action panel shows focused ${getCapabilityFocusLabel(focus!)} capabilities with tool-level detail. Briefly summarize the most relevant things Grimoire can do there instead of reading the full list aloud.`
      }));
    },

    show_compose_form: async (args, deps): Promise<ToolRuntimeHandlerResult> => {
      const { store } = deps;
      const requestedPreset = (args.preset as string || 'generic') as FormPresetId;
      const title = args.title as string || 'Compose';
      let formDescription = args.description as string | undefined;

      let prefill: Record<string, string> = {};
      if (args.prefill_json) {
        try {
          const parsedPrefill = parseJsonLikeArg(args.prefill_json);
          if (!isObjectRecord(parsedPrefill)) {
            throw new Error('prefill_json must be an object');
          }
          prefill = normalizeParsedPrefill(parsedPrefill);
        }
        catch { logService.warning('llm', 'Invalid prefill_json for show_compose_form'); }
      }
      let staticArgs: Record<string, unknown> = {};
      if (args.static_args_json) {
        try {
          const parsedStaticArgs = parseJsonLikeArg(args.static_args_json);
          if (!isObjectRecord(parsedStaticArgs)) {
            throw new Error('static_args_json must be an object');
          }
          staticArgs = normalizeComposeStaticArgs(parsedStaticArgs);
        }
        catch { logService.warning('llm', 'Invalid static_args_json for show_compose_form'); }
      }

      const resolvedPreset = resolveComposePreset(requestedPreset, title, formDescription, staticArgs, prefill);
      const preset = resolvedPreset.preset;
      prefill = sanitizeComposeFilenamePrefill(preset, resolvedPreset.prefill);
      const presetConfig = getFormPreset(preset);
      let formFields: IFormFieldDefinition[] = presetConfig.fields;
      let submissionTarget: IFormSubmissionTarget = presetConfig.submissionTarget;

      if (preset === 'generic') {
        if (args.custom_fields_json) {
          try {
            const parsedFields = parseJsonLikeArg(args.custom_fields_json);
            if (!Array.isArray(parsedFields)) {
              throw new Error('custom_fields_json must be an array');
            }
            formFields = parsedFields as IFormFieldDefinition[];
          }
          catch { logService.warning('llm', 'Invalid custom_fields_json for show_compose_form'); }
        }
        if (args.custom_target_json) {
          try {
            const parsedTarget = parseJsonLikeArg(args.custom_target_json);
            if (!isObjectRecord(parsedTarget)) {
              throw new Error('custom_target_json must be an object');
            }
            submissionTarget = parsedTarget as unknown as IFormSubmissionTarget;
          }
          catch { logService.warning('llm', 'Invalid custom_target_json for show_compose_form'); }
        }
      }

      const enrichedComposePrefill = enrichComposePrefillFromSession(preset, title, formFields, prefill, staticArgs);
      prefill = enrichedComposePrefill.prefill;
      staticArgs = enrichedComposePrefill.staticArgs;
      submissionTarget.targetContext = mergeMcpTargetContexts(
        submissionTarget.targetContext,
        enrichedComposePrefill.targetContext
      );
      formDescription = appendComposeDescriptionNote(formDescription, enrichedComposePrefill.descriptionNote);
      const enhancedFields = await enhanceComposeFieldsForSharePreset(preset, formFields, store);
      formFields = enhancedFields.fields;
      formDescription = appendComposeDescriptionNote(formDescription, enhancedFields.descriptionNote);
      if (shouldSuppressComposeDescription(preset)) {
        formDescription = undefined;
      }

      const prefillKeys = Object.keys(prefill);
      for (let i = 0; i < prefillKeys.length; i++) {
        const key = prefillKeys[i];
        const field = formFields.find((f) => f.key === key);
        if (field) {
          field.defaultValue = field.type === 'datetime'
            ? normalizeDateTimeForForm(prefill[key])
            : prefill[key];
        }
      }

      submissionTarget.staticArgs = { ...submissionTarget.staticArgs, ...staticArgs };
      submissionTarget.targetContext = resolveComposeSubmissionTargetContext(
        store,
        deps.sourceContext,
        submissionTarget.targetContext,
        submissionTarget.staticArgs,
        prefill
      );
      const formData: IFormData = {
        kind: 'form',
        preset,
        description: formDescription,
        fields: formFields,
        submissionTarget,
        status: 'editing'
      };

      const formBlock = createBlock('form', title, formData);
      trackCreatedBlock(store, formBlock, deps);
      trackToolCompletion('show_compose_form', formBlock.id, true, 0, deps);

      if (preset !== requestedPreset) {
        logService.info('llm', `Form preset normalized: ${requestedPreset} → ${preset}`);
      }
      logService.info('llm', `Form displayed: ${preset} — ${title}`);
      return completeOutcome(JSON.stringify({
        success: true,
        message: 'Form displayed in the action panel. Wait for the user to fill it in and submit. Do NOT ask for the same information via chat — the form handles it.'
      }));
    },

    get_my_profile: (_args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const { store, aadClient, awaitAsync } = deps;
      const uc = store.userContext;
      if (uc) {
        const profileBlock = createBlock('user-card', uc.displayName || 'My Profile', {
          kind: 'user-card',
          displayName: uc.displayName,
          email: uc.email,
          jobTitle: uc.jobTitle,
          department: uc.department
        } as IUserCardData);
        trackCreatedBlock(store, profileBlock, deps);
        const manager = uc.manager ? `, manager: ${uc.manager}` : '';
        return completeOutcome(JSON.stringify({
          success: true,
          displayName: uc.displayName, jobTitle: uc.jobTitle, department: uc.department,
          manager: uc.manager || undefined,
          message: `Profile card displayed. The user is ${uc.displayName}, ${uc.jobTitle || 'no title'}${manager}.`
        }));
      }
      if (!aadClient) {
        return errorOutcome(JSON.stringify({ success: false, error: 'Graph API not available — cannot fetch profile' }));
      }

      const graphSvc = new GraphService(aadClient);
      const profilePromise = graphSvc.get<{
        displayName?: string; mail?: string; jobTitle?: string;
        department?: string; officeLocation?: string; mobilePhone?: string;
        manager?: { displayName?: string };
      }>('/me?$select=displayName,mail,jobTitle,department,officeLocation,mobilePhone&$expand=manager($select=displayName)')
        .then((result) => {
          if (!result.success || !result.data) {
            return errorOutcome(JSON.stringify({ success: false, error: result.error || 'Failed to fetch profile' }));
          }
          const d = result.data;
          const profileBlock = createBlock('user-card', d.displayName || 'My Profile', {
            kind: 'user-card',
            displayName: d.displayName || '',
            email: d.mail || '',
            jobTitle: d.jobTitle,
            department: d.department,
            officeLocation: d.officeLocation,
            phone: d.mobilePhone
          } as IUserCardData);
          trackCreatedBlock(store, profileBlock, deps);
          const manager = d.manager?.displayName ? `, manager: ${d.manager.displayName}` : '';
          return completeOutcome(JSON.stringify({
            success: true,
            displayName: d.displayName, jobTitle: d.jobTitle, department: d.department,
            manager: d.manager?.displayName || undefined,
            message: `Profile card displayed. The user is ${d.displayName}, ${d.jobTitle || 'no title'}${manager}.`
          }));
        });
      if (awaitAsync) return profilePromise;
      return completeOutcome(JSON.stringify({ success: true, message: 'Loading your profile...' }));
    },

    get_recent_documents: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const maxRecent = parseInt((args.max_results as string) || '10', 10);
      return fetchInsightsDocuments(
        'get_recent_documents', 'Recent Documents',
        `/me/insights/used?$top=${maxRecent}&$orderby=lastUsed/lastAccessedDateTime desc`,
        (item: Record<string, unknown>) => {
          const viz = item.resourceVisualization as Record<string, string> | undefined;
          const ref = item.resourceReference as Record<string, string> | undefined;
          const lastUsed = item.lastUsed as Record<string, string> | undefined;
          return {
            title: viz?.title || 'Untitled',
            url: ref?.webUrl || '',
            summary: `Last accessed: ${lastUsed?.lastAccessedDateTime ? new Date(lastUsed.lastAccessedDateTime).toLocaleDateString() : 'unknown'}`,
            fileType: viz?.type || '',
            sources: ['insights']
          };
        },
        deps
      );
    },

    get_trending_documents: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const maxTrending = parseInt((args.max_results as string) || '10', 10);
      return fetchInsightsDocuments(
        'get_trending_documents', 'Trending Documents',
        `/me/insights/trending?$top=${maxTrending}`,
        (item: Record<string, unknown>) => {
          const viz = item.resourceVisualization as Record<string, string> | undefined;
          const ref = item.resourceReference as Record<string, string> | undefined;
          const weight = item.weight as number | undefined;
          return {
            title: viz?.title || 'Untitled',
            url: ref?.webUrl || '',
            summary: `Trending ${weight ? `(relevance: ${weight})` : ''}`,
            fileType: viz?.type || '',
            sources: ['insights']
          };
        },
        deps
      );
    },

    save_note: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const noteText = args.text as string;
      if (!noteText) {
        return errorOutcome(JSON.stringify({ success: false, error: 'Missing required field: text' }));
      }
      if (!deps.store.proxyConfig || !deps.store.userApiClient) {
        return errorOutcome(JSON.stringify({ success: false, error: 'Secured user API is not available for saving notes' }));
      }
      const tagStr = args.tags as string || '';
      const savePromise = PersistenceService.saveNote(
        deps.store.proxyConfig,
        deps.store.userApiClient,
        noteText,
        tagStr ? tagStr.split(',').map((t: string) => t.trim()) : []
      ).then((result: { id: string }) => {
        return completeOutcome(JSON.stringify({ success: true, noteId: result.id, message: `Note saved. I'll remember: "${noteText.substring(0, 80)}..."` }));
      }).catch((err: Error) => {
        return errorOutcome(JSON.stringify({ success: false, error: err.message || 'Failed to save note' }));
      });
      if (deps.awaitAsync) return savePromise;
      return completeOutcome(JSON.stringify({ success: true, message: 'Saving note...' }));
    },

    recall_notes: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const keyword = args.keyword as string || '';
      if (!deps.store.proxyConfig || !deps.store.userApiClient) {
        return errorOutcome(JSON.stringify({ success: false, error: 'Secured user API is not available for recalling notes' }));
      }
      const recallPromise = PersistenceService.listNotes(deps.store.proxyConfig, deps.store.userApiClient, keyword)
        .then((notes: Array<{ id: string; text: string; tags: string[]; createdAt: string }>) => {
          if (notes.length === 0) {
            return completeOutcome(JSON.stringify({ success: true, count: 0, message: keyword ? `No notes found matching "${keyword}".` : 'No saved notes yet.' }));
          }
          const noteLines = notes.map((n: { id: string; text: string; tags: string[]; createdAt: string }, i: number) => {
            const date = new Date(n.createdAt).toLocaleDateString();
            const tags = n.tags.length > 0 ? ` [${n.tags.join(', ')}]` : '';
            return `${i + 1}. ${n.text}${tags} *(${date})* — ID: \`${n.id}\``;
          });
          const notesContent = keyword
            ? `## Notes matching "${keyword}"\n\n${noteLines.join('\n')}`
            : `## Saved Notes\n\n${noteLines.join('\n')}`;
          const notesBlock = createBlock('markdown', keyword ? `Notes: ${keyword}` : 'Saved Notes', {
            kind: 'markdown', content: notesContent
          } as IMarkdownData);
          trackCreatedBlock(deps.store, notesBlock, deps);
          const notesSummary = notes.map((n: { id: string; text: string; tags: string[] }) => ({
            id: n.id, text: n.text.substring(0, 100), tags: n.tags
          }));
          return completeOutcome(JSON.stringify({ success: true, count: notes.length, notes: notesSummary, message: `Found ${notes.length} note(s). Displayed in the action panel.` }));
        })
        .catch((err: Error) => {
          return errorOutcome(JSON.stringify({ success: false, error: err.message || 'Failed to recall notes' }));
        });
      if (deps.awaitAsync) return recallPromise;
      return completeOutcome(JSON.stringify({ success: true, message: 'Recalling notes...' }));
    },

    delete_note: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const noteId = args.note_id as string;
      if (!noteId) {
        return errorOutcome(JSON.stringify({ success: false, error: 'Missing required field: note_id' }));
      }
      if (!deps.store.proxyConfig || !deps.store.userApiClient) {
        return errorOutcome(JSON.stringify({ success: false, error: 'Secured user API is not available for deleting notes' }));
      }

      const deletePromise = (async (): Promise<ToolRuntimeHandlerResult> => {
        try {
          if (noteId === 'all') {
            const allNotes = await PersistenceService.listNotes(deps.store.proxyConfig!, deps.store.userApiClient);
            if (allNotes.length === 0) {
              return completeOutcome(JSON.stringify({ success: true, deleted: 0, message: 'No notes to delete.' }));
            }
            const deletePromises = allNotes.map((n: { id: string }) =>
              PersistenceService.deleteNote(deps.store.proxyConfig!, deps.store.userApiClient!, n.id)
            );
            await Promise.all(deletePromises);
            return completeOutcome(JSON.stringify({ success: true, deleted: allNotes.length, message: `Deleted all ${allNotes.length} note(s).` }));
          }
          await PersistenceService.deleteNote(deps.store.proxyConfig!, deps.store.userApiClient, noteId);
          return completeOutcome(JSON.stringify({ success: true, deleted: 1, message: `Note ${noteId} deleted.` }));
        } catch (err) {
          const msg = err instanceof Error ? err.message : 'Failed to delete note';
          return errorOutcome(JSON.stringify({ success: false, error: msg }));
        }
      })();
      if (deps.awaitAsync) return deletePromise;
      return completeOutcome(JSON.stringify({ success: true, message: 'Deleting note...' }));
    }
  };
}
