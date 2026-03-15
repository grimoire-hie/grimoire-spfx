import type { IProxyConfig } from '../../store/useGrimoireStore';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import {
  createBlock,
  type IBlock,
  type IInfoCardData,
  type IMarkdownData,
  type IProgressTrackerData,
  type IProgressTrackerStep,
  type ProgressTrackerStepStatus,
  type ISearchResultsData,
  type IUserCardData
} from '../../models/IBlock';
import { COMPOUND_WORKFLOW_PLANNER_SYSTEM_PROMPT } from '../../config/promptCatalog';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { getNanoService } from '../nano/NanoService';
import { logService } from '../logging/LogService';
import { BlockRecapService, getRecapOriginTool } from '../recap/BlockRecapService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { resolveCurrentArtifactContext, resolveLatestArtifactContext } from '../hie/HieArtifactLinkage';
import {
  clearPendingMailDiscussionReplyAllPlan,
  getCurrentMailDiscussionMessageId,
  MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
  setPendingMailDiscussionReplyAllPlan
} from './MailDiscussionReplyAllWorkflow';
import {
  CompoundWorkflowStopError,
  extractMarkdownEmailCandidates,
  fetchMailDiscussionRecipients,
  getPlausibleMailThreadCandidates,
  tryDropLeadingProjectPrefix,
  buildMailThreadSubjectFirstSearchQuery,
  buildMailThreadFallbackSearchQuery,
  buildMailThreadChooserDescription,
  type IMarkdownEmailCandidate,
  type IMailThreadCandidate
} from './CompoundWorkflowMailHelpers';

type WorkflowSearchToolName = 'search_sharepoint' | 'search_emails';
type WorkflowStepKind =
  | 'search'
  | 'find-person'
  | 'recap-results'
  | 'summarize-document'
  | 'summarize-email'
  | 'compose-email'
  | 'resolve-mail-thread'
  | 'share-teams-chat'
  | 'share-teams-channel';

export type CompoundWorkflowFamilyId =
  | 'search_sharepoint_recap'
  | 'search_sharepoint_recap_email'
  | 'search_sharepoint_recap_teams_chat'
  | 'search_sharepoint_recap_teams_channel'
  | 'search_sharepoint_summarize_document'
  | 'search_sharepoint_summarize_document_email'
  | 'search_people_email'
  | 'search_emails_summarize_email'
  | typeof MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID;

type CompoundWorkflowPlannerCode =
  | 'cw1'
  | 'cw2'
  | 'cw3'
  | 'cw4'
  | 'cw5'
  | 'cw6'
  | 'cw7'
  | 'cw8';

type CompoundWorkflowPlannerDomainCode = 'sp' | 'pe' | 'em';
type CompoundWorkflowPlannerSummaryTargetCode = 'r' | 'i' | 'u';
type CompoundWorkflowPlannerFollowUpActionCode = 'n' | 'e' | 'tc' | 'tn';
type CompoundWorkflowPlanningDomain = 'search_sharepoint' | 'search_people' | 'search_emails';
type CompoundWorkflowSummaryTarget = 'results' | 'item' | 'unknown';
type CompoundWorkflowFollowUpAction = 'none' | 'email' | 'teams-chat' | 'teams-channel';

export interface ICompoundWorkflowSlots {
  query: string;
  selectionHint?: 'none' | 'first' | 'top' | 'best' | 'latest';
}

export interface ICompoundWorkflowStep {
  id: string;
  kind: WorkflowStepKind;
  label: string;
}

export interface ICompoundWorkflowFamily {
  id: CompoundWorkflowFamilyId;
  label: string;
  steps: ICompoundWorkflowStep[];
  searchTool?: WorkflowSearchToolName;
}

export interface ICompoundWorkflowPlan {
  shouldPlan: true;
  familyId: CompoundWorkflowFamilyId;
  confidence: number;
  slots: ICompoundWorkflowSlots;
  steps: ICompoundWorkflowStep[];
  label: string;
  userText: string;
}

export interface ICompoundWorkflowExecutionState {
  plan: ICompoundWorkflowPlan;
  trackerBlockId: string;
  steps: IProgressTrackerStep[];
  searchBlockId?: string;
  recapBlockId?: string;
  summaryBlockId?: string;
  composeBlockId?: string;
  selectedDocumentUrl?: string;
  selectedDocumentTitle?: string;
  selectedPersonEmail?: string;
  selectedPersonName?: string;
  selectedEmailSubject?: string;
  selectedEmailSender?: string;
  selectedEmailDateHint?: string;
  selectedMailMessageId?: string;
  selectedMailThreadSubject?: string;
}

interface ILegacyCompoundWorkflowPlannerResponse {
  shouldPlan?: boolean;
  familyId?: string;
  confidence?: number;
  slots?: {
    query?: string;
    selectionHint?: string;
  };
}

interface ICompactCompoundWorkflowPlannerResponse {
  p?: boolean | number;
  d?: string;
  q?: string;
  t?: string;
  a?: string;
  s?: string;
  c?: number;
  f?: string;
}

interface ICompoundWorkflowPlannerSlotsDecision {
  domain?: CompoundWorkflowPlanningDomain;
  query?: string;
  summaryTarget: CompoundWorkflowSummaryTarget;
  followUpAction: CompoundWorkflowFollowUpAction;
  selectionHint: CompoundWorkflowSelectionHint;
}

interface ICompoundWorkflowPlannerNormalization {
  shouldPlan: boolean;
  confidence: number;
  slots: ICompoundWorkflowPlannerSlotsDecision;
}

interface ICompoundWorkflowShareArtifact {
  blockId: string;
  title: string;
  body: string;
}

type CompoundWorkflowPlannerRejectionReason =
  | 'not_compound'
  | 'malformed_json'
  | 'missing_query'
  | 'unknown_domain'
  | 'unknown_summary_target'
  | 'unsupported_combo'
  | 'low_confidence';

interface ICompoundWorkflowPlannerEvaluation {
  normalized?: ICompoundWorkflowPlannerNormalization;
  plan?: ICompoundWorkflowPlan;
  rejectionReason?: CompoundWorkflowPlannerRejectionReason;
}

type CompoundWorkflowSelectionHint = NonNullable<ICompoundWorkflowSlots['selectionHint']>;

interface IToolInvocationCallbacks {
  onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string | Promise<string>;
}

interface INanoClassifier {
  classify: (
    systemPrompt: string,
    userMessage: string,
    timeoutMs?: number,
    maxTokens?: number
  ) => Promise<string | undefined>;
}

interface IToolInvocationResult {
  parsed?: Record<string, unknown>;
  raw: string;
  createdBlocks: IBlock[];
}

const EXTERNAL_WEB_HINTS: ReadonlyArray<RegExp> = [
  /\bweb\b/i,
  /\binternet\b/i,
  /\bwebsite\b/i,
  /\bgithub\b/i,
  /\bwikipedia\b/i,
  /https?:\/\//i
];

const SUMMARY_REQUEST_HINTS: ReadonlyArray<RegExp> = [
  /\bsummar/i,
  /\brecap\b/i,
  /riassum/i,
  /resum/i,
  /résum/i,
  /zusammenfass/i,
  /\bfass(?:e|t|en)?\b/i
];

const SHARE_DESTINATION_HINTS: ReadonlyArray<RegExp> = [
  /\be-?mail\b/i,
  /\bmail\b/i,
  /\bposta\b/i,
  /\bcorreo\b/i,
  /\bcourriel\b/i,
  /\bteams?\b/i
];

const FAMILY_CATALOG: Record<CompoundWorkflowFamilyId, ICompoundWorkflowFamily> = {
  search_sharepoint_recap: {
    id: 'search_sharepoint_recap',
    label: 'Search and recap results',
    searchTool: 'search_sharepoint',
    steps: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' }
    ]
  },
  search_sharepoint_recap_email: {
    id: 'search_sharepoint_recap_email',
    label: 'Search, recap, and prepare email',
    searchTool: 'search_sharepoint',
    steps: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ]
  },
  search_sharepoint_recap_teams_chat: {
    id: 'search_sharepoint_recap_teams_chat',
    label: 'Search, recap, and prepare Teams chat share',
    searchTool: 'search_sharepoint',
    steps: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' },
      { id: 'share-teams-chat', kind: 'share-teams-chat', label: 'Open Teams chat share' }
    ]
  },
  search_sharepoint_recap_teams_channel: {
    id: 'search_sharepoint_recap_teams_channel',
    label: 'Search, recap, and prepare Teams channel share',
    searchTool: 'search_sharepoint',
    steps: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' },
      { id: 'share-teams-channel', kind: 'share-teams-channel', label: 'Open Teams channel share' }
    ]
  },
  search_sharepoint_summarize_document: {
    id: 'search_sharepoint_summarize_document',
    label: 'Search and summarize the top document',
    searchTool: 'search_sharepoint',
    steps: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'summarize-document', kind: 'summarize-document', label: 'Summarize the selected document' }
    ]
  },
  search_sharepoint_summarize_document_email: {
    id: 'search_sharepoint_summarize_document_email',
    label: 'Search, summarize the top document, and prepare email',
    searchTool: 'search_sharepoint',
    steps: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'summarize-document', kind: 'summarize-document', label: 'Summarize the selected document' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ]
  },
  search_people_email: {
    id: 'search_people_email',
    label: 'Find a person and prepare email',
    steps: [
      { id: 'find-person', kind: 'find-person', label: 'Find the person' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ]
  },
  search_emails_summarize_email: {
    id: 'search_emails_summarize_email',
    label: 'Search emails and summarize the top email',
    searchTool: 'search_emails',
    steps: [
      { id: 'search', kind: 'search', label: 'Search emails' },
      { id: 'summarize-email', kind: 'summarize-email', label: 'Summarize the selected email' }
    ]
  },
  visible_recap_reply_all_mail_discussion: {
    id: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
    label: 'Email visible recap to mail participants',
    steps: [
      { id: 'resolve-mail-thread', kind: 'resolve-mail-thread', label: 'Resolve the mail discussion thread' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ]
  }
};

const HACKATHON_SUPPORTED_COMPOUND_FAMILIES: ReadonlySet<CompoundWorkflowFamilyId> = new Set<CompoundWorkflowFamilyId>([
  'search_sharepoint_recap_email',
  'search_sharepoint_recap_teams_chat',
  'search_sharepoint_recap_teams_channel',
  'search_sharepoint_summarize_document_email',
  MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
]);

const VALID_FAMILY_IDS: ReadonlySet<string> = new Set(Object.keys(FAMILY_CATALOG));
const COMPOUND_WORKFLOW_CODE_TO_FAMILY_ID: Record<CompoundWorkflowPlannerCode, CompoundWorkflowFamilyId> = {
  cw1: 'search_sharepoint_recap',
  cw2: 'search_sharepoint_recap_email',
  cw3: 'search_sharepoint_recap_teams_chat',
  cw4: 'search_sharepoint_recap_teams_channel',
  cw5: 'search_sharepoint_summarize_document',
  cw6: 'search_sharepoint_summarize_document_email',
  cw7: 'search_people_email',
  cw8: 'search_emails_summarize_email'
};

const LEGACY_FAMILY_ID_TO_SLOTS: Record<
  CompoundWorkflowFamilyId,
  Pick<ICompoundWorkflowPlannerSlotsDecision, 'domain' | 'summaryTarget' | 'followUpAction'>
> = {
  search_sharepoint_recap: {
    domain: 'search_sharepoint',
    summaryTarget: 'results',
    followUpAction: 'none'
  },
  search_sharepoint_recap_email: {
    domain: 'search_sharepoint',
    summaryTarget: 'results',
    followUpAction: 'email'
  },
  search_sharepoint_recap_teams_chat: {
    domain: 'search_sharepoint',
    summaryTarget: 'results',
    followUpAction: 'teams-chat'
  },
  search_sharepoint_recap_teams_channel: {
    domain: 'search_sharepoint',
    summaryTarget: 'results',
    followUpAction: 'teams-channel'
  },
  search_sharepoint_summarize_document: {
    domain: 'search_sharepoint',
    summaryTarget: 'item',
    followUpAction: 'none'
  },
  search_sharepoint_summarize_document_email: {
    domain: 'search_sharepoint',
    summaryTarget: 'item',
    followUpAction: 'email'
  },
  search_people_email: {
    domain: 'search_people',
    summaryTarget: 'unknown',
    followUpAction: 'email'
  },
  search_emails_summarize_email: {
    domain: 'search_emails',
    summaryTarget: 'item',
    followUpAction: 'none'
  },
  visible_recap_reply_all_mail_discussion: {
    domain: 'search_emails',
    summaryTarget: 'unknown',
    followUpAction: 'email'
  }
};

const COMPOUND_WORKFLOW_DOMAIN_BY_CODE: Record<CompoundWorkflowPlannerDomainCode, CompoundWorkflowPlanningDomain> = {
  sp: 'search_sharepoint',
  pe: 'search_people',
  em: 'search_emails'
};

const COMPOUND_WORKFLOW_SUMMARY_TARGET_BY_CODE: Record<
  CompoundWorkflowPlannerSummaryTargetCode,
  CompoundWorkflowSummaryTarget
> = {
  r: 'results',
  i: 'item',
  u: 'unknown'
};

const COMPOUND_WORKFLOW_FOLLOW_UP_ACTION_BY_CODE: Record<
  CompoundWorkflowPlannerFollowUpActionCode,
  CompoundWorkflowFollowUpAction
> = {
  n: 'none',
  e: 'email',
  tc: 'teams-chat',
  tn: 'teams-channel'
};

const COMPOUND_WORKFLOW_SELECTION_HINT_BY_CODE: Record<string, CompoundWorkflowSelectionHint> = {
  n: 'none',
  f: 'first',
  t: 'top',
  b: 'best',
  l: 'latest'
};

function matchesAny(text: string, patterns: ReadonlyArray<RegExp>): boolean {
  return patterns.some((pattern) => pattern.test(text));
}

function detectSelectionHint(text: string): CompoundWorkflowSelectionHint {
  if (/\blatest\b/i.test(text)) return 'latest';
  if (/\btop\b/i.test(text)) return 'top';
  if (/\bbest\b/i.test(text)) return 'best';
  if (/\bfirst\b/i.test(text)) return 'first';
  return 'none';
}

export function shouldConsiderCompoundWorkflow(text: string): boolean {
  const normalized = text.trim();
  if (!normalized) return false;
  if (matchesAny(normalized, EXTERNAL_WEB_HINTS)) return false;
  if (!matchesAny(normalized, SUMMARY_REQUEST_HINTS)) return false;
  if (!matchesAny(normalized, SHARE_DESTINATION_HINTS)) return false;
  return true;
}

function parsePlannerJson(
  raw: string | undefined
): ILegacyCompoundWorkflowPlannerResponse | ICompactCompoundWorkflowPlannerResponse | undefined {
  if (!raw) return undefined;
  const trimmed = raw.trim();
  const codeFenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = codeFenceMatch?.[1]?.trim() || trimmed;
  const firstBrace = candidate.indexOf('{');
  const lastBrace = candidate.lastIndexOf('}');
  if (firstBrace === -1 || lastBrace === -1 || lastBrace < firstBrace) {
    return undefined;
  }

  try {
    return JSON.parse(candidate.slice(firstBrace, lastBrace + 1)) as (
      ILegacyCompoundWorkflowPlannerResponse | ICompactCompoundWorkflowPlannerResponse
    );
  } catch {
    return undefined;
  }
}

function normalizePlannerSelectionHint(
  legacySelectionHint: string | undefined,
  compactSelectionHintCode: string | undefined,
  userText: string,
  allowUserTextFallback: boolean
): CompoundWorkflowSelectionHint {
  const compactSelectionHint = compactSelectionHintCode
    ? COMPOUND_WORKFLOW_SELECTION_HINT_BY_CODE[compactSelectionHintCode.trim().toLowerCase()]
    : undefined;
  if (compactSelectionHint) {
    return compactSelectionHint;
  }

  const plannerSelectionHint = typeof legacySelectionHint === 'string'
    ? legacySelectionHint.trim().toLowerCase()
    : '';
  if (
    plannerSelectionHint === 'none'
    || plannerSelectionHint === 'first'
    || plannerSelectionHint === 'top'
    || plannerSelectionHint === 'best'
    || plannerSelectionHint === 'latest'
  ) {
    return plannerSelectionHint as CompoundWorkflowSelectionHint;
  }

  return allowUserTextFallback ? detectSelectionHint(userText) : 'none';
}

function normalizePlannerResponse(
  parsed: ILegacyCompoundWorkflowPlannerResponse | ICompactCompoundWorkflowPlannerResponse,
  userText: string
): ICompoundWorkflowPlannerNormalization {
  const compactParsed = parsed as ICompactCompoundWorkflowPlannerResponse;
  const legacyParsed = parsed as ILegacyCompoundWorkflowPlannerResponse;
  const compactPlanFlag = typeof compactParsed.p === 'number'
    ? compactParsed.p === 1
    : compactParsed.p === true;
  const legacyPlanFlag = legacyParsed.shouldPlan === true;
  const compactConfidence = typeof compactParsed.c === 'number' ? compactParsed.c : undefined;
  const legacyConfidence = typeof legacyParsed.confidence === 'number' ? legacyParsed.confidence : undefined;
  const shouldPlan = compactPlanFlag || legacyPlanFlag;

  const compactDomainCode = typeof compactParsed.d === 'string'
    ? compactParsed.d.trim().toLowerCase() as CompoundWorkflowPlannerDomainCode
    : undefined;
  const compactSummaryTargetCode = typeof compactParsed.t === 'string'
    ? compactParsed.t.trim().toLowerCase() as CompoundWorkflowPlannerSummaryTargetCode
    : undefined;
  const compactFollowUpActionCode = typeof compactParsed.a === 'string'
    ? compactParsed.a.trim().toLowerCase() as CompoundWorkflowPlannerFollowUpActionCode
    : undefined;
  const compactQuery = typeof compactParsed.q === 'string'
    ? compactParsed.q.trim()
    : '';

  const compactDomain = compactDomainCode
    ? COMPOUND_WORKFLOW_DOMAIN_BY_CODE[compactDomainCode]
    : undefined;
  const compactSummaryTarget = compactSummaryTargetCode
    ? COMPOUND_WORKFLOW_SUMMARY_TARGET_BY_CODE[compactSummaryTargetCode]
    : undefined;
  const compactFollowUpAction = compactFollowUpActionCode
    ? COMPOUND_WORKFLOW_FOLLOW_UP_ACTION_BY_CODE[compactFollowUpActionCode]
    : undefined;

  const compactLegacyFamilyCode = typeof compactParsed.f === 'string'
    ? compactParsed.f.trim().toLowerCase()
    : '';
  const compactLegacyFamilyId = compactLegacyFamilyCode
    ? COMPOUND_WORKFLOW_CODE_TO_FAMILY_ID[compactLegacyFamilyCode as CompoundWorkflowPlannerCode]
    : undefined;
  const legacyFamilyId = typeof legacyParsed.familyId === 'string'
    ? legacyParsed.familyId.trim()
    : '';
  const normalizedLegacyFamilyId = compactLegacyFamilyId
    || (legacyFamilyId && VALID_FAMILY_IDS.has(legacyFamilyId)
      ? legacyFamilyId as CompoundWorkflowFamilyId
      : undefined);
  const legacySlots = normalizedLegacyFamilyId
    ? LEGACY_FAMILY_ID_TO_SLOTS[normalizedLegacyFamilyId]
    : undefined;
  const legacyQuery = typeof legacyParsed.slots?.query === 'string'
    ? legacyParsed.slots.query.trim()
    : '';

  const normalizedSelectionHint = normalizePlannerSelectionHint(
    legacyParsed.slots?.selectionHint,
    typeof compactParsed.s === 'string' ? compactParsed.s : undefined,
    userText,
    !!normalizedLegacyFamilyId || typeof legacyParsed.slots?.selectionHint === 'string'
  );

  const normalizedSlots: ICompoundWorkflowPlannerSlotsDecision = {
    domain: compactDomain || legacySlots?.domain,
    query: compactQuery || legacyQuery || undefined,
    summaryTarget: compactSummaryTarget || legacySlots?.summaryTarget || 'unknown',
    followUpAction: compactFollowUpAction || legacySlots?.followUpAction || 'none',
    selectionHint: normalizedSelectionHint
  };
  if (normalizedSlots.domain === 'search_sharepoint' && normalizedSlots.summaryTarget === 'unknown') {
    normalizedSlots.summaryTarget = 'results';
  }

  if (!shouldPlan) {
    return {
      shouldPlan: false,
      confidence: compactConfidence ?? legacyConfidence ?? 0,
      slots: normalizedSlots
    };
  }

  return {
    shouldPlan: true,
    confidence: compactConfidence ?? legacyConfidence ?? 0,
    slots: normalizedSlots
  };
}

function deriveCompoundWorkflowFamilyId(
  slots: ICompoundWorkflowPlannerSlotsDecision
): CompoundWorkflowFamilyId | undefined {
  switch (slots.domain) {
    case 'search_sharepoint':
      if (slots.summaryTarget === 'results') {
        switch (slots.followUpAction) {
          case 'none':
            return 'search_sharepoint_recap';
          case 'email':
            return 'search_sharepoint_recap_email';
          case 'teams-chat':
            return 'search_sharepoint_recap_teams_chat';
          case 'teams-channel':
            return 'search_sharepoint_recap_teams_channel';
          default:
            return undefined;
        }
      }
      if (slots.summaryTarget === 'item') {
        switch (slots.followUpAction) {
          case 'none':
            return 'search_sharepoint_summarize_document';
          case 'email':
            return 'search_sharepoint_summarize_document_email';
          default:
            return undefined;
        }
      }
      return undefined;
    case 'search_people':
      return slots.followUpAction === 'email' ? 'search_people_email' : undefined;
    case 'search_emails':
      return slots.summaryTarget === 'item' && slots.followUpAction === 'none'
        ? 'search_emails_summarize_email'
        : undefined;
    default:
      return undefined;
  }
}

function evaluateCompoundWorkflowPlannerResponse(
  raw: string | undefined,
  userText: string,
  confidenceThreshold: number = getRuntimeTuningConfig().nano.compoundWorkflowPlannerConfidenceThreshold
): ICompoundWorkflowPlannerEvaluation {
  const parsed = parsePlannerJson(raw);
  if (!parsed) {
    return { rejectionReason: 'malformed_json' };
  }

  const normalized = normalizePlannerResponse(parsed, userText);
  if (!normalized.shouldPlan) {
    return { normalized, rejectionReason: 'not_compound' };
  }

  if (normalized.confidence < confidenceThreshold) {
    return { normalized, rejectionReason: 'low_confidence' };
  }

  if (!normalized.slots.query) {
    return { normalized, rejectionReason: 'missing_query' };
  }

  if (!normalized.slots.domain) {
    return { normalized, rejectionReason: 'unknown_domain' };
  }

  const familyId = deriveCompoundWorkflowFamilyId(normalized.slots);
  if (!familyId) {
    return {
      normalized,
      rejectionReason: normalized.slots.summaryTarget === 'unknown'
        ? 'unknown_summary_target'
        : 'unsupported_combo'
    };
  }
  if (!HACKATHON_SUPPORTED_COMPOUND_FAMILIES.has(familyId)) {
    return { normalized, rejectionReason: 'unsupported_combo' };
  }

  const family = FAMILY_CATALOG[familyId];

  const plan: ICompoundWorkflowPlan = {
    shouldPlan: true,
    familyId,
    confidence: normalized.confidence,
    slots: {
      query: normalized.slots.query,
      selectionHint: normalized.slots.selectionHint
    },
    steps: family.steps.map((step) => ({ ...step })),
    label: family.label,
    userText
  };

  return { normalized, plan };
}

export function parseCompoundWorkflowPlannerResponse(
  raw: string | undefined,
  userText: string,
  confidenceThreshold: number = getRuntimeTuningConfig().nano.compoundWorkflowPlannerConfidenceThreshold
): ICompoundWorkflowPlan | undefined {
  return evaluateCompoundWorkflowPlannerResponse(raw, userText, confidenceThreshold).plan;
}

export async function planCompoundWorkflow(
  text: string,
  proxyConfig: IProxyConfig | undefined,
  nanoService?: Pick<INanoClassifier, 'classify'>
): Promise<ICompoundWorkflowPlan | undefined> {
  if (!shouldConsiderCompoundWorkflow(text)) {
    return undefined;
  }

  const nano = nanoService || getNanoService(proxyConfig);
  if (!nano) {
    logService.debug('llm', 'Compound workflow planner skipped: fast model unavailable');
    return undefined;
  }

  try {
    const tuning = getRuntimeTuningConfig().nano;
    const raw = await nano.classify(
      COMPOUND_WORKFLOW_PLANNER_SYSTEM_PROMPT,
      text,
      tuning.compoundWorkflowPlannerTimeoutMs,
      tuning.compoundWorkflowPlannerMaxTokens
    );
    const evaluation = evaluateCompoundWorkflowPlannerResponse(
      raw,
      text,
      tuning.compoundWorkflowPlannerConfidenceThreshold
    );
    const normalized = evaluation.normalized;
    if (normalized) {
      logService.info(
        'llm',
        `Compound workflow planner slots: ${JSON.stringify({
          shouldPlan: normalized.shouldPlan,
          confidence: Number(normalized.confidence.toFixed(2)),
          domain: normalized.slots.domain || 'unknown',
          query: normalized.slots.query || '',
          summaryTarget: normalized.slots.summaryTarget,
          followUpAction: normalized.slots.followUpAction,
          selectionHint: normalized.slots.selectionHint
        })}`
      );
    }
    if (evaluation.plan) {
      logService.info(
        'llm',
        `Compound workflow planned: ${evaluation.plan.familyId} (${evaluation.plan.confidence.toFixed(2)})`
      );
    } else {
      logService.info(
        'llm',
        `Compound workflow planner rejected: ${evaluation.rejectionReason || 'unknown'}`
      );
    }
    return evaluation.plan;
  } catch (error) {
    logService.debug(
      'llm',
      `Compound workflow planner failed: ${error instanceof Error ? error.message : 'Unknown error'}`
    );
    return undefined;
  }
}

function parseInvocationPayload(raw: string): Record<string, unknown> | undefined {
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
      return parsed as Record<string, unknown>;
    }
    return undefined;
  } catch {
    return undefined;
  }
}

async function invokeTool(
  callbacks: IToolInvocationCallbacks,
  callId: string,
  funcName: string,
  args: Record<string, unknown>
): Promise<IToolInvocationResult> {
  const before = new Set(useGrimoireStore.getState().blocks.map((block) => block.id));
  const raw = await Promise.resolve(callbacks.onFunctionCall(callId, funcName, args));
  const parsed = parseInvocationPayload(raw);
  if (parsed?.success === false) {
    const errorMessage = typeof parsed.error === 'string'
      ? parsed.error
      : (typeof parsed.message === 'string' ? parsed.message : `${funcName} failed`);
    throw new Error(errorMessage);
  }

  const createdBlocks = useGrimoireStore.getState().blocks.filter((block) => !before.has(block.id));
  return { parsed, raw, createdBlocks };
}

function pushWorkflowBlock(block: IBlock): void {
  const store = useGrimoireStore.getState();
  store.pushBlock(block);
  hybridInteractionEngine.onBlockCreated(block);
}

function insertWorkflowBlockAfter(referenceBlockId: string, block: IBlock): void {
  const store = useGrimoireStore.getState();
  store.insertBlockAfter(referenceBlockId, block);
  hybridInteractionEngine.onBlockCreated(block);
}

function updateWorkflowBlock(blockId: string, nextBlock: IBlock): void {
  const store = useGrimoireStore.getState();
  store.updateBlock(blockId, {
    title: nextBlock.title,
    data: nextBlock.data,
    timestamp: nextBlock.timestamp,
    originTool: nextBlock.originTool
  });
  hybridInteractionEngine.onBlockUpdated(blockId, nextBlock);
}

function buildWorkflowStepPreview(plan: ICompoundWorkflowPlan): string {
  return plan.steps.map((step) => step.label).join(' -> ');
}

function buildTrackerData(
  plan: ICompoundWorkflowPlan,
  steps: IProgressTrackerStep[],
  status: IProgressTrackerData['status'],
  detail: string,
  currentStep?: number
): IProgressTrackerData {
  const completedCount = steps.filter((step) => step.status === 'complete').length;
  const progress = steps.length > 0 ? Math.round((completedCount / steps.length) * 100) : 0;

  return {
    kind: 'progress-tracker',
    label: plan.label,
    progress,
    status,
    detail,
    steps,
    currentStep
  };
}

function persistTrackerState(
  execution: ICompoundWorkflowExecutionState,
  status: IProgressTrackerData['status'],
  detail: string,
  currentStep?: number
): void {
  const nextData = buildTrackerData(execution.plan, execution.steps, status, detail, currentStep);
  const currentBlock = useGrimoireStore.getState().blocks.find((block) => block.id === execution.trackerBlockId);
  if (!currentBlock) return;
  const nextBlock: IBlock = {
    ...currentBlock,
    title: execution.plan.label,
    data: nextData,
    timestamp: new Date()
  };
  updateWorkflowBlock(execution.trackerBlockId, nextBlock);
}

function updateStep(
  execution: ICompoundWorkflowExecutionState,
  index: number,
  status: ProgressTrackerStepStatus,
  detail?: string,
  blockRefs?: Partial<Pick<IProgressTrackerStep, 'sourceBlockId' | 'derivedBlockId'>>
): void {
  execution.steps = execution.steps.map((step, stepIndex) => {
    if (stepIndex !== index) return step;
    return {
      ...step,
      status,
      detail: detail || step.detail,
      sourceBlockId: blockRefs?.sourceBlockId || step.sourceBlockId,
      derivedBlockId: blockRefs?.derivedBlockId || step.derivedBlockId
    };
  });
}

function patchExecutionState(
  execution: ICompoundWorkflowExecutionState,
  updates: Partial<ICompoundWorkflowExecutionState>
): void {
  Object.assign(execution, updates);
}

function createWorkflowTracker(plan: ICompoundWorkflowPlan): ICompoundWorkflowExecutionState {
  const baseSteps = plan.steps.map((step) => ({
    id: step.id,
    label: step.label,
    status: 'pending' as const
  }));
  const initialDetail = `Planned steps: ${buildWorkflowStepPreview(plan)}`;

  const trackerBlock = createBlock(
    'progress-tracker',
    plan.label,
    buildTrackerData(plan, baseSteps, 'running', initialDetail, 0),
    true,
    undefined,
    { originTool: 'compound-workflow' }
  );
  const previousActiveActionBlockId = useGrimoireStore.getState().activeActionBlockId;
  pushWorkflowBlock(trackerBlock);
  if (previousActiveActionBlockId) {
    useGrimoireStore.getState().setActiveActionBlock(previousActiveActionBlockId);
  }

  return {
    plan,
    trackerBlockId: trackerBlock.id,
    steps: baseSteps
  };
}

function findFirstCreatedBlock(
  createdBlocks: IBlock[],
  predicate: (block: IBlock) => boolean
): IBlock | undefined {
  for (let i = 0; i < createdBlocks.length; i++) {
    if (predicate(createdBlocks[i])) {
      return createdBlocks[i];
    }
  }
  return undefined;
}

function resolveSingleDocumentSelection(
  execution: ICompoundWorkflowExecutionState
): { title: string; url: string } | undefined {
  const searchBlock = execution.searchBlockId
    ? useGrimoireStore.getState().blocks.find((block) => block.id === execution.searchBlockId)
    : undefined;
  if (!searchBlock || searchBlock.type !== 'search-results') {
    return undefined;
  }

  const data = searchBlock.data as ISearchResultsData;
  const results = data.results.filter((result) => !!result.url && result.fileType !== 'folder');
  if (results.length === 0) return undefined;

  if (results.length === 1 || execution.plan.slots.selectionHint !== 'none') {
    const selected = results[0];
    return {
      title: selected.title || 'Document',
      url: selected.url
    };
  }

  return undefined;
}

function resolveSinglePersonSelection(
  execution: ICompoundWorkflowExecutionState,
  createdBlocks: IBlock[],
  count: number
): { name: string; email: string } | undefined {
  const cards = createdBlocks.filter((block) => block.type === 'user-card');
  if (cards.length === 0) return undefined;

  if (count === 1 || execution.plan.slots.selectionHint !== 'none') {
    const selected = cards[0].data as IUserCardData;
    if (!selected.email) return undefined;
    return {
      name: selected.displayName || selected.email,
      email: selected.email
    };
  }

  return undefined;
}

function resolveSingleEmailSelection(
  execution: ICompoundWorkflowExecutionState
): IMarkdownEmailCandidate | undefined {
  const searchBlock = execution.searchBlockId
    ? useGrimoireStore.getState().blocks.find((block) => block.id === execution.searchBlockId)
    : undefined;
  if (!searchBlock) return undefined;

  const candidates = extractMarkdownEmailCandidates(searchBlock);
  if (candidates.length === 0) return undefined;

  if (candidates.length === 1 || execution.plan.slots.selectionHint !== 'none') {
    return candidates[0];
  }

  return undefined;
}

function buildComposeStaticArgs(execution: ICompoundWorkflowExecutionState): string | undefined {
  if (!execution.selectedDocumentUrl) return undefined;

  return JSON.stringify({
    attachmentUris: [execution.selectedDocumentUrl],
    shareItemTitle: execution.selectedDocumentTitle || 'Selected document',
    fileOrFolderUrl: execution.selectedDocumentUrl,
    fileOrFolderName: execution.selectedDocumentTitle || 'Selected document'
  });
}

function buildMailDiscussionComposeSubject(title: string | undefined): string {
  const normalized = (title || '')
    .replace(/^summary(?:\s*:|\s+of\b)\s*/i, '')
    .replace(/^recap(?:\s*:|\s+of\b)?\s*/i, '')
    .replace(/\.[a-z0-9]+$/i, '')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!normalized) {
    return 'Recap';
  }
  if (/\b(?:recap|summary)\b/i.test(normalized)) {
    return normalized;
  }
  return `${normalized} Recap`;
}

function workflowRequiresSummaryArtifact(plan: ICompoundWorkflowPlan): boolean {
  if (plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID) {
    return true;
  }

  return plan.steps.some(
    (step) => step.kind === 'recap-results' || step.kind === 'summarize-document' || step.kind === 'summarize-email'
  );
}

function looksLikeSummaryArtifactTitle(title: string | undefined): boolean {
  const trimmed = title?.trim() || '';
  if (!trimmed) return false;
  if (/^summary\s*:/i.test(trimmed) || /^recap\s*:/i.test(trimmed)) return true;
  return /\bsummary\b/i.test(trimmed) || /\brecap\b/i.test(trimmed);
}

function getInfoCardShareArtifact(blockId: string | undefined): ICompoundWorkflowShareArtifact | undefined {
  if (!blockId) return undefined;

  const block = useGrimoireStore.getState().blocks.find((candidate) => candidate.id === blockId);
  if (!block || block.type !== 'info-card') {
    return undefined;
  }

  const data = block.data as IInfoCardData;
  const body = typeof data.body === 'string' ? data.body.trim() : '';
  if (!body) {
    return undefined;
  }

  const title = typeof data.heading === 'string' && data.heading.trim()
    ? data.heading.trim()
    : block.title;
  return {
    blockId: block.id,
    title: title || 'Summary',
    body
  };
}

function getMarkdownShareArtifact(blockId: string | undefined): ICompoundWorkflowShareArtifact | undefined {
  if (!blockId) return undefined;

  const block = useGrimoireStore.getState().blocks.find((candidate) => candidate.id === blockId);
  if (!block || block.type !== 'markdown') {
    return undefined;
  }

  const data = block.data as IMarkdownData;
  const body = typeof data.content === 'string' ? data.content.trim() : '';
  if (!body) {
    return undefined;
  }

  return {
    blockId: block.id,
    title: block.title || 'Summary',
    body
  };
}

function getBlockShareArtifact(blockId: string | undefined): ICompoundWorkflowShareArtifact | undefined {
  return getInfoCardShareArtifact(blockId) || getMarkdownShareArtifact(blockId);
}

function resolveVisibleShareArtifactFromSession(): ICompoundWorkflowShareArtifact | undefined {
  const state = useGrimoireStore.getState();
  const currentArtifactContext = resolveCurrentArtifactContext(
    hybridInteractionEngine.getCurrentTaskContext(),
    hybridInteractionEngine.getCurrentArtifacts()
  );
  const latestArtifactContext = resolveLatestArtifactContext(hybridInteractionEngine.getCurrentArtifacts());

  const candidateBlockIds: string[] = [];
  const pushArtifactCandidates = (
    context: ReturnType<typeof resolveCurrentArtifactContext> | ReturnType<typeof resolveLatestArtifactContext>
  ): void => {
    const artifacts = [
      context.currentArtifact,
      context.primaryArtifact,
      ...context.artifactChain.slice().reverse()
    ];
    for (let i = 0; i < artifacts.length; i++) {
      const artifact = artifacts[i];
      if (!artifact) {
        continue;
      }
      if (artifact.artifactKind !== 'summary' && artifact.artifactKind !== 'recap') {
        continue;
      }
      const blockId = artifact.blockId || artifact.artifactId;
      if (blockId) {
        candidateBlockIds.push(blockId);
      }
    }
  };

  pushArtifactCandidates(currentArtifactContext);
  pushArtifactCandidates(latestArtifactContext);

  if (state.activeActionBlockId) {
    const activeBlock = state.blocks.find((block) => block.id === state.activeActionBlockId);
    if (
      activeBlock
      && (
        activeBlock.type === 'info-card'
        || (activeBlock.type === 'markdown' && looksLikeSummaryArtifactTitle(activeBlock.title))
      )
    ) {
      candidateBlockIds.push(state.activeActionBlockId);
    }
  }

  state.blocks
    .slice()
    .reverse()
    .forEach((block) => {
      if ((block.type === 'info-card' || block.type === 'markdown') && looksLikeSummaryArtifactTitle(block.title)) {
        candidateBlockIds.push(block.id);
      }
    });

  const seen = new Set<string>();
  for (let i = 0; i < candidateBlockIds.length; i++) {
    const blockId = candidateBlockIds[i];
    if (!blockId || seen.has(blockId)) {
      continue;
    }
    seen.add(blockId);
    const artifact = getBlockShareArtifact(blockId);
    if (artifact) {
      return artifact;
    }
  }

  return undefined;
}

function getShareArtifact(execution: ICompoundWorkflowExecutionState): ICompoundWorkflowShareArtifact | undefined {
  return getBlockShareArtifact(execution.summaryBlockId)
    || getBlockShareArtifact(execution.recapBlockId)
    || resolveVisibleShareArtifactFromSession();
}

function getRequiredShareArtifact(execution: ICompoundWorkflowExecutionState): ICompoundWorkflowShareArtifact | undefined {
  const artifact = getShareArtifact(execution);
  if (!workflowRequiresSummaryArtifact(execution.plan)) {
    return artifact;
  }
  if (!artifact) {
    throw new CompoundWorkflowStopError(
      'No visible recap or summary is available to share.',
      'I need a visible recap or summary in the panel before I can prepare the email draft. Create the recap first.'
    );
  }
  return artifact;
}

function upsertRecapBlock(
  sourceBlock: IBlock,
  body: string,
  mode: 'loading' | 'ready' | 'error'
): IBlock {
  const recapTitle = `Recap: ${sourceBlock.title}`;
  const recapOriginTool = getRecapOriginTool(sourceBlock.id);
  const icon = mode === 'error' ? 'StatusErrorFull' : 'AlignLeft';
  const recapBody = mode === 'loading' ? 'Generating recap...' : body;
  const data: IInfoCardData = {
    kind: 'info-card',
    heading: recapTitle,
    body: recapBody,
    icon
  };

  const existing = useGrimoireStore.getState().blocks.find(
    (block) => block.type === 'info-card' && block.originTool === recapOriginTool
  );

  if (existing) {
    const nextBlock: IBlock = {
      ...existing,
      title: recapTitle,
      data,
      timestamp: new Date(),
      originTool: recapOriginTool
    };
    updateWorkflowBlock(existing.id, nextBlock);
    return nextBlock;
  }

  const nextBlock = createBlock(
    'info-card',
    recapTitle,
    data,
    true,
    undefined,
    { originTool: recapOriginTool }
  );
  insertWorkflowBlockAfter(sourceBlock.id, nextBlock);
  return nextBlock;
}

async function executeSearchStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number
): Promise<void> {
  const family = FAMILY_CATALOG[execution.plan.familyId];
  const searchTool = family.searchTool;
  if (!searchTool) {
    throw new Error(`Workflow family ${family.id} does not define a search tool.`);
  }

  const label = searchTool === 'search_emails' ? 'Searching emails...' : 'Searching SharePoint...';
  updateStep(execution, stepIndex, 'running', label);
  persistTrackerState(execution, 'running', label, stepIndex);

  const result = await invokeTool(callbacks, `compound-${family.id}-search`, searchTool, {
    query: execution.plan.slots.query
  });
  const searchBlock = searchTool === 'search_emails'
    ? findFirstCreatedBlock(result.createdBlocks, (block) => block.type === 'markdown')
    : findFirstCreatedBlock(result.createdBlocks, (block) => block.type === 'search-results');

  if (!searchBlock) {
    throw new Error('The search step completed without producing a visible result block.');
  }

  patchExecutionState(execution, { searchBlockId: searchBlock.id });
  const count = typeof result.parsed?.count === 'number'
    ? result.parsed.count
    : (typeof result.parsed?.displayedResults === 'number' ? result.parsed.displayedResults : 0);
  if (count <= 0) {
    throw new CompoundWorkflowStopError(
      'No results found for the requested search.',
      `I couldn't find results for "${execution.plan.slots.query}".`
    );
  }
  const detail = count > 0 ? `${count} result${count === 1 ? '' : 's'} displayed.` : 'No results found.';

  updateStep(execution, stepIndex, 'complete', detail, { derivedBlockId: searchBlock.id });
  persistTrackerState(execution, 'running', detail, stepIndex + 1);
}

async function executeFindPersonStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number
): Promise<void> {
  const detail = 'Finding matching people...';
  updateStep(execution, stepIndex, 'running', detail);
  persistTrackerState(execution, 'running', detail, stepIndex);

  const result = await invokeTool(callbacks, `compound-${execution.plan.familyId}-find-person`, 'search_people', {
    query: execution.plan.slots.query
  });
  const count = typeof result.parsed?.count === 'number' ? result.parsed.count : result.createdBlocks.length;
  if (count <= 0) {
    throw new CompoundWorkflowStopError(
      'No people matched the request.',
      `I couldn't find anyone matching "${execution.plan.slots.query}".`
    );
  }
  const selected = resolveSinglePersonSelection(execution, result.createdBlocks, count);
  if (!selected) {
    throw new CompoundWorkflowStopError(
      count > 1
        ? 'Multiple people matched. I stopped before drafting the email.'
        : 'I could not resolve a single person to draft the email.',
      count > 1
        ? 'I found multiple people. Pick one and then ask me to draft the email.'
        : 'I could not resolve a single person to draft the email.'
    );
  }

  patchExecutionState(execution, {
    selectedPersonEmail: selected.email,
    selectedPersonName: selected.name
  });
  updateStep(execution, stepIndex, 'complete', `Matched ${selected.name}.`);
  persistTrackerState(execution, 'running', `Matched ${selected.name}.`, stepIndex + 1);
}

async function executeRecapResultsStep(
  execution: ICompoundWorkflowExecutionState,
  stepIndex: number
): Promise<void> {
  const sourceBlock = execution.searchBlockId
    ? useGrimoireStore.getState().blocks.find((block) => block.id === execution.searchBlockId)
    : undefined;
  if (!sourceBlock) {
    throw new Error('No search-results block is available to recap.');
  }

  const detail = 'Generating recap...';
  updateStep(execution, stepIndex, 'running', detail, { sourceBlockId: sourceBlock.id });
  persistTrackerState(execution, 'running', detail, stepIndex);

  upsertRecapBlock(sourceBlock, detail, 'loading');

  try {
    const text = await new BlockRecapService().generate(sourceBlock, useGrimoireStore.getState().proxyConfig);
    const recapBlock = upsertRecapBlock(sourceBlock, text, 'ready');
    patchExecutionState(execution, { recapBlockId: recapBlock.id });
    updateStep(execution, stepIndex, 'complete', 'Recap ready.', {
      sourceBlockId: sourceBlock.id,
      derivedBlockId: recapBlock.id
    });
    persistTrackerState(execution, 'running', 'Recap ready.', stepIndex + 1);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'The recap could not be generated.';
    const recapBlock = upsertRecapBlock(sourceBlock, message, 'error');
    patchExecutionState(execution, { recapBlockId: recapBlock.id });
    throw new Error(message);
  }
}

async function executeSummarizeDocumentStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number
): Promise<void> {
  const selected = resolveSingleDocumentSelection(execution);
  if (!selected) {
    throw new CompoundWorkflowStopError(
      'Multiple documents matched. I left the results visible so you can pick one.',
      'I found multiple documents. Pick one and I can summarize it next.'
    );
  }

  patchExecutionState(execution, {
    selectedDocumentTitle: selected.title,
    selectedDocumentUrl: selected.url
  });
  const detail = `Summarizing ${selected.title}...`;
  updateStep(execution, stepIndex, 'running', detail, { sourceBlockId: execution.searchBlockId });
  persistTrackerState(execution, 'running', detail, stepIndex);

  const readResult = await invokeTool(callbacks, `compound-${execution.plan.familyId}-read-file`, 'read_file_content', {
    file_url: selected.url,
    file_name: selected.title,
    mode: 'summarize'
  });
  const content = typeof readResult.parsed?.content === 'string' ? readResult.parsed.content.trim() : '';
  if (!content) {
    throw new Error('The document summary step did not return summary text.');
  }

  const infoResult = await invokeTool(callbacks, `compound-${execution.plan.familyId}-show-file-summary`, 'show_info_card', {
    heading: `Summary: ${selected.title}`,
    body: content,
    icon: 'AlignLeft'
  });
  const summaryBlock = findFirstCreatedBlock(infoResult.createdBlocks, (block) => block.type === 'info-card');
  patchExecutionState(execution, { summaryBlockId: summaryBlock?.id });
  updateStep(execution, stepIndex, 'complete', `Summary ready for ${selected.title}.`, {
    sourceBlockId: execution.searchBlockId,
    derivedBlockId: summaryBlock?.id
  });
  persistTrackerState(execution, 'running', `Summary ready for ${selected.title}.`, stepIndex + 1);
}

async function executeSummarizeEmailStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number
): Promise<void> {
  const selected = resolveSingleEmailSelection(execution);
  if (!selected) {
    throw new CompoundWorkflowStopError(
      'Multiple emails matched. I left the results visible so you can pick one.',
      'I found multiple emails. Pick one and I can summarize it next.'
    );
  }

  patchExecutionState(execution, {
    selectedEmailSubject: selected.subject,
    selectedEmailSender: selected.sender,
    selectedEmailDateHint: selected.date
  });
  const detail = `Summarizing "${selected.subject}"...`;
  updateStep(execution, stepIndex, 'running', detail, { sourceBlockId: execution.searchBlockId });
  persistTrackerState(execution, 'running', detail, stepIndex);

  const readResult = await invokeTool(callbacks, `compound-${execution.plan.familyId}-read-email`, 'read_email_content', {
    subject: selected.subject,
    ...(selected.sender ? { sender: selected.sender } : {}),
    ...(selected.date ? { date_hint: selected.date } : {}),
    mode: 'summarize'
  });
  const content = typeof readResult.parsed?.content === 'string' ? readResult.parsed.content.trim() : '';
  if (!content) {
    throw new Error('The email summary step did not return summary text.');
  }

  const infoResult = await invokeTool(callbacks, `compound-${execution.plan.familyId}-show-email-summary`, 'show_info_card', {
    heading: `Summary: ${selected.subject}`,
    body: content,
    icon: 'AlignLeft'
  });
  const summaryBlock = findFirstCreatedBlock(infoResult.createdBlocks, (block) => block.type === 'info-card');
  patchExecutionState(execution, { summaryBlockId: summaryBlock?.id });
  updateStep(execution, stepIndex, 'complete', `Summary ready for "${selected.subject}".`, {
    sourceBlockId: execution.searchBlockId,
    derivedBlockId: summaryBlock?.id
  });
  persistTrackerState(execution, 'running', `Summary ready for "${selected.subject}".`, stepIndex + 1);
}

async function executeComposeEmailStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number
): Promise<void> {
  const detail = execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
    ? 'Preparing email draft with recipients...'
    : 'Opening email draft...';
  updateStep(execution, stepIndex, 'running', detail);
  persistTrackerState(execution, 'running', detail, stepIndex);

  const artifact = getRequiredShareArtifact(execution);
  const toolArgs: Record<string, unknown> = {
    preset: 'email-compose',
    title: execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
      ? 'Email Mail Participants'
      : (execution.selectedPersonName ? `Draft Email to ${execution.selectedPersonName}` : 'Share by Email')
  };

  const prefill: Record<string, string> = execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
    ? (() => {
        const messageId = execution.selectedMailMessageId || getCurrentMailDiscussionMessageId();
        if (!messageId) {
          throw new CompoundWorkflowStopError(
            'No mail discussion thread is selected.',
            'I need a selected email before I can prepare the recipient draft.'
          );
        }

        return {
          subject: buildMailDiscussionComposeSubject(artifact?.title),
          body: artifact?.body || ''
        };
      })()
    : {};

  if (execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID) {
    const messageId = execution.selectedMailMessageId || getCurrentMailDiscussionMessageId();
    if (!messageId) {
      throw new CompoundWorkflowStopError(
        'No mail discussion thread is selected.',
        'I need a selected email before I can prepare the recipient draft.'
      );
    }
    const recipients = await fetchMailDiscussionRecipients(messageId);
    if (recipients.to.length > 0) {
      prefill.to = recipients.to.join(', ');
    }
    if (recipients.cc.length > 0) {
      prefill.cc = recipients.cc.join(', ');
    }
  } else {
    if (execution.selectedPersonEmail) {
      prefill.to = execution.selectedPersonEmail;
    }
    if (artifact?.title) {
      prefill.subject = artifact.title;
    }
    if (artifact?.body) {
      prefill.body = artifact.body;
    }
  }

  if (Object.keys(prefill).length > 0) {
    toolArgs.prefill_json = JSON.stringify(prefill);
  }

  const staticArgsJson = execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
    ? JSON.stringify({ skipSessionHydration: true })
    : buildComposeStaticArgs(execution);
  if (staticArgsJson) {
    toolArgs.static_args_json = staticArgsJson;
  }

  const result = await invokeTool(callbacks, `compound-${execution.plan.familyId}-compose-email`, 'show_compose_form', toolArgs);
  if (execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID) {
    clearPendingMailDiscussionReplyAllPlan();
  }
  const formBlock = findFirstCreatedBlock(result.createdBlocks, (block) => block.type === 'form');
  patchExecutionState(execution, { composeBlockId: formBlock?.id });
  const completionDetail = execution.plan.familyId === MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
    ? 'Email draft ready with recipients.'
    : 'Email draft ready.';
  updateStep(execution, stepIndex, 'complete', completionDetail, { derivedBlockId: formBlock?.id });
  persistTrackerState(execution, 'running', completionDetail, stepIndex + 1);
}

async function executeResolveMailThreadStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number
): Promise<void> {
  const currentMessageId = getCurrentMailDiscussionMessageId();
  const artifact = getRequiredShareArtifact(execution);
  if (!artifact) {
    throw new CompoundWorkflowStopError(
      'No visible recap or summary is available to share.',
      'I need a visible recap or summary in the panel before I can prepare the email draft. Create the recap first.'
    );
  }

  const detail = currentMessageId
    ? 'Using the currently selected mail discussion...'
    : 'Resolving the mail discussion thread...';
  updateStep(execution, stepIndex, 'running', detail);
  persistTrackerState(execution, 'running', detail, stepIndex);

  if (currentMessageId) {
    clearPendingMailDiscussionReplyAllPlan();
    patchExecutionState(execution, { selectedMailMessageId: currentMessageId });
    updateStep(execution, stepIndex, 'complete', 'Using the currently selected mail discussion.');
    persistTrackerState(execution, 'running', 'Using the currently selected mail discussion.', stepIndex + 1);
    return;
  }

  const threadQuery = execution.plan.slots.query.trim();
  if (!threadQuery || threadQuery === 'current mail discussion') {
    throw new CompoundWorkflowStopError(
      'Missing mail discussion subject.',
      'I need the email subject or thread name before I can prepare the draft.'
    );
  }

  const shortenedQuery = tryDropLeadingProjectPrefix(threadQuery);

  const attempts: Array<{ mode: 'subject-first' | 'fallback'; query: string; scoringQuery: string; callId: string }> = [
    {
      mode: 'subject-first',
      query: buildMailThreadSubjectFirstSearchQuery(threadQuery),
      scoringQuery: threadQuery,
      callId: `compound-${execution.plan.familyId}-resolve-mail-thread-subject`
    },
    {
      mode: 'fallback',
      query: buildMailThreadFallbackSearchQuery(threadQuery),
      scoringQuery: threadQuery,
      callId: `compound-${execution.plan.familyId}-resolve-mail-thread-fallback`
    }
  ];
  if (shortenedQuery) {
    attempts.push({
      mode: 'fallback',
      query: buildMailThreadFallbackSearchQuery(shortenedQuery),
      scoringQuery: shortenedQuery,
      callId: `compound-${execution.plan.familyId}-resolve-mail-thread-shortened`
    });
  }

  let chooserCandidates: IMailThreadCandidate[] = [];
  for (let i = 0; i < attempts.length; i++) {
    const attempt = attempts[i];
    const result = await invokeTool(callbacks, attempt.callId, 'search_emails', {
      query: attempt.query,
      max_results: '10'
    });
    const searchBlock = findFirstCreatedBlock(result.createdBlocks, (block) => (
      block.type === 'markdown' && Object.keys(((block.data as IMarkdownData).itemIds || {})).length > 0
    ));
    if (searchBlock) {
      patchExecutionState(execution, { searchBlockId: searchBlock.id });
    }

    const candidates = getPlausibleMailThreadCandidates(searchBlock, attempt.scoringQuery, attempt.mode);
    if (candidates.length === 1 && candidates[0].matchKind === 'exact-subject') {
      clearPendingMailDiscussionReplyAllPlan();
      patchExecutionState(execution, {
        selectedMailMessageId: candidates[0].messageId,
        selectedMailThreadSubject: candidates[0].subject
      });
      const successDetail = `Resolved "${candidates[0].subject}".`;
      updateStep(execution, stepIndex, 'complete', successDetail, { derivedBlockId: searchBlock?.id });
      persistTrackerState(execution, 'running', successDetail, stepIndex + 1);
      return;
    }

    if (candidates.length > 0) {
      chooserCandidates = candidates;
      break;
    }
  }

  if (chooserCandidates.length === 0) {
    clearPendingMailDiscussionReplyAllPlan();
    throw new CompoundWorkflowStopError(
      `No matching email thread found for "${threadQuery}".`,
      `I couldn't find a mail thread whose subject matches "${threadQuery}".`
    );
  }

  await invokeTool(
    callbacks,
    `compound-${execution.plan.familyId}-choose-mail-thread`,
    'show_selection_list',
    {
      prompt: 'Choose the email to use for recipients.',
      items_json: JSON.stringify(
        chooserCandidates.map((candidate) => ({
          id: candidate.messageId,
          label: candidate.subject,
          ...(buildMailThreadChooserDescription(candidate)
            ? { description: buildMailThreadChooserDescription(candidate) }
            : {}),
          itemType: 'email',
          targetContext: {
            mailItemId: candidate.messageId,
            source: 'explicit-user'
          }
        }))
      ),
      multi_select: 'false'
    }
  );
  setPendingMailDiscussionReplyAllPlan(execution.plan);
  throw new CompoundWorkflowStopError(
    'Mail discussion selection required.',
    `I found ${chooserCandidates.length} matching email${chooserCandidates.length === 1 ? '' : 's'}. Choose one in the panel and I will prepare the draft with the recap.`
  );
}

async function executeTeamsShareStep(
  execution: ICompoundWorkflowExecutionState,
  callbacks: IToolInvocationCallbacks,
  stepIndex: number,
  target: 'chat' | 'channel'
): Promise<void> {
  const detail = target === 'chat' ? 'Opening Teams chat share...' : 'Opening Teams channel share...';
  updateStep(execution, stepIndex, 'running', detail);
  persistTrackerState(execution, 'running', detail, stepIndex);

  const toolArgs: Record<string, unknown> = target === 'chat'
    ? {
        preset: 'share-teams-chat',
        title: 'Share to Teams Chat',
        description: 'Choose recipients, review the message, then send.'
      }
    : {
        preset: 'share-teams-channel',
        title: 'Post to a Teams Channel',
        description: 'Choose the destination channel, review the message, then send.'
      };

  const artifact = getRequiredShareArtifact(execution);
  if (artifact) {
    const prefill = target === 'chat'
      ? { topic: artifact.title, content: artifact.body }
      : { content: artifact.body };
    toolArgs.prefill_json = JSON.stringify(prefill);
  }

  const result = await invokeTool(
    callbacks,
    `compound-${execution.plan.familyId}-share-teams-${target}`,
    'show_compose_form',
    toolArgs
  );
  const formBlock = findFirstCreatedBlock(result.createdBlocks, (block) => block.type === 'form');
  patchExecutionState(execution, { composeBlockId: formBlock?.id });
  updateStep(execution, stepIndex, 'complete', 'Share form ready.', { derivedBlockId: formBlock?.id });
  persistTrackerState(execution, 'running', 'Share form ready.', stepIndex + 1);
}

function buildSuccessMessage(execution: ICompoundWorkflowExecutionState): string {
  switch (execution.plan.familyId) {
    case 'search_sharepoint_recap':
      return 'I searched SharePoint and added a recap in the action panel.';
    case 'search_sharepoint_recap_email':
      return 'I searched SharePoint, added a recap, and opened an email draft in the action panel.';
    case 'search_sharepoint_recap_teams_chat':
      return 'I searched SharePoint, added a recap, and opened a Teams chat share form in the action panel.';
    case 'search_sharepoint_recap_teams_channel':
      return 'I searched SharePoint, added a recap, and opened a Teams channel share form in the action panel.';
    case 'search_sharepoint_summarize_document':
      return 'I searched SharePoint and summarized the selected document in the action panel.';
    case 'search_sharepoint_summarize_document_email':
      return 'I searched SharePoint, summarized the selected document, and opened an email draft in the action panel.';
    case 'search_people_email':
      return execution.selectedPersonName
        ? `I found ${execution.selectedPersonName} and opened an email draft in the action panel.`
        : 'I found the person and opened an email draft in the action panel.';
    case 'search_emails_summarize_email':
      return 'I searched your emails and summarized the selected message in the action panel.';
    case 'visible_recap_reply_all_mail_discussion':
      return 'I opened an email draft with the visible recap and the selected email recipients in the action panel.';
    default:
      return 'I finished the workflow in the action panel.';
  }
}

export async function executeCompoundWorkflowPlan(
  plan: ICompoundWorkflowPlan,
  callbacks: IToolInvocationCallbacks
): Promise<string> {
  const execution = createWorkflowTracker(plan);

  try {
    for (let i = 0; i < plan.steps.length; i++) {
      const step = plan.steps[i];
      switch (step.kind) {
        case 'search':
          await executeSearchStep(execution, callbacks, i);
          break;
        case 'find-person':
          await executeFindPersonStep(execution, callbacks, i);
          break;
        case 'recap-results':
          await executeRecapResultsStep(execution, i);
          break;
        case 'summarize-document':
          await executeSummarizeDocumentStep(execution, callbacks, i);
          break;
        case 'summarize-email':
          await executeSummarizeEmailStep(execution, callbacks, i);
          break;
        case 'compose-email':
          await executeComposeEmailStep(execution, callbacks, i);
          break;
        case 'resolve-mail-thread':
          await executeResolveMailThreadStep(execution, callbacks, i);
          break;
        case 'share-teams-chat':
          await executeTeamsShareStep(execution, callbacks, i, 'chat');
          break;
        case 'share-teams-channel':
          await executeTeamsShareStep(execution, callbacks, i, 'channel');
          break;
        default:
          throw new Error(`Unsupported workflow step: ${step.kind}`);
      }
    }

    persistTrackerState(execution, 'complete', 'Workflow complete.');
    return buildSuccessMessage(execution);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'The workflow could not be completed.';
    const failedStepIndex = execution.steps.findIndex((step) => step.status === 'running');
    if (failedStepIndex !== -1) {
      updateStep(execution, failedStepIndex, 'error', message);
    }
    persistTrackerState(execution, 'error', message, failedStepIndex === -1 ? undefined : failedStepIndex);
    if (error instanceof CompoundWorkflowStopError) {
      return error.assistantMessage;
    }
    throw error;
  }
}
