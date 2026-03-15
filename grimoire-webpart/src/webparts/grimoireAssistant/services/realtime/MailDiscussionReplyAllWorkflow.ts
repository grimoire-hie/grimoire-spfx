import type { ICompoundWorkflowPlan } from './CompoundWorkflowExecutor';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { hasShareableSessionContent } from '../sharing/SessionShareFormatter';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { isHiePromptMessage } from '../hie/HiePromptProtocol';
import { deriveMcpTargetContextFromHie } from '../mcp/McpTargetContext';

const MAIL_DISCUSSION_COMPOSE_HINTS: ReadonlyArray<RegExp> = [
  /\bsend\b/i,
  /\bshare\b/i,
  /\bemail\b/i,
  /\bmail\b/i,
  /\breply all\b/i
];

const MAIL_DISCUSSION_THREAD_HINTS: ReadonlyArray<RegExp> = [
  /\breply all\b/i,
  /\b(?:mail|email)\s+(?:discussion|thread|conversation)\b/i,
  /\bthread participants?\b/i,
  /\bdiscussion participants?\b/i,
  /\bauthors?\b/i,
  /\bpeople involved\b/i,
  /\bperson involved\b/i,
  /\beveryone involved\b/i,
  /\ball involved\b/i,
  /\beveryone who replied\b/i,
  /\beveryone who was included\b/i,
  /\bwho replied\b/i,
  /\bwas included\b/i
];

export const MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID = 'visible_recap_reply_all_mail_discussion' as const;

interface IPendingMailDiscussionReplyAllWorkflow {
  plan: ICompoundWorkflowPlan;
  createdAt: number;
}

let pendingMailDiscussionReplyAllWorkflow: IPendingMailDiscussionReplyAllWorkflow | undefined;

function matchesAny(text: string, patterns: ReadonlyArray<RegExp>): boolean {
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(text)) {
      return true;
    }
  }
  return false;
}

function looksLikeSummaryArtifactTitle(title: string | undefined): boolean {
  const trimmed = title?.trim() || '';
  if (!trimmed) {
    return false;
  }

  if (/^summary(?:\s*:|\s+of\b)/i.test(trimmed) || /^recap(?:\s*:|\s+of\b)?/i.test(trimmed)) {
    return true;
  }

  return /\bsummary\b/i.test(trimmed) || /\brecap\b/i.test(trimmed);
}

function hasVisibleReplyAllShareArtifact(): boolean {
  const state = useGrimoireStore.getState();
  const activeBlock = state.activeActionBlockId
    ? state.blocks.find((block) => block.id === state.activeActionBlockId)
    : undefined;

  if (activeBlock && (activeBlock.type === 'info-card' || activeBlock.type === 'markdown') && looksLikeSummaryArtifactTitle(activeBlock.title)) {
    return true;
  }

  for (let i = state.blocks.length - 1; i >= 0; i--) {
    const block = state.blocks[i];
    if ((block.type === 'info-card' || block.type === 'markdown') && looksLikeSummaryArtifactTitle(block.title)) {
      return true;
    }
  }

  return false;
}

function canPlanMailDiscussionReplyAll(): boolean {
  const state = useGrimoireStore.getState();
  return hasShareableSessionContent(state.blocks, state.transcript) && hasVisibleReplyAllShareArtifact();
}

function clonePlan(plan: ICompoundWorkflowPlan): ICompoundWorkflowPlan {
  return {
    ...plan,
    slots: {
      ...plan.slots
    },
    steps: plan.steps.map((step) => ({ ...step }))
  };
}

export function looksLikeMailDiscussionReplyAllRequest(text: string): boolean {
  if (/\breply all\b/i.test(text)) {
    return true;
  }

  return matchesAny(text, MAIL_DISCUSSION_COMPOSE_HINTS) && matchesAny(text, MAIL_DISCUSSION_THREAD_HINTS);
}

function normalizeMailSearchQuery(value: string | undefined): string | undefined {
  let normalized = value?.trim() || '';
  if (!normalized) {
    return undefined;
  }

  while (/^(?:re|fw|fwd|wg)\s*:/i.test(normalized)) {
    normalized = normalized.replace(/^(?:re|fw|fwd|wg)\s*:\s*/i, '').trim();
  }

  normalized = normalized
    .replace(/['"`]/g, '')
    .replace(/[_-]+/g, ' ')
    .replace(/[^a-z0-9\s]/gi, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (normalized.length < 3) {
    return undefined;
  }

  // Deictic pronouns are never meaningful search queries
  if (/^(?:this|that|these|those|current|it)$/i.test(normalized)) {
    return undefined;
  }

  return normalized;
}

export function extractMailDiscussionSearchQuery(text: string): string | undefined {
  const patterns: ReadonlyArray<RegExp> = [
    /\b(?:in|from|on)\s+(?:the\s+)?(.+?)\s+(?:email|mail)\b/i,
    /\b(?:reply all to|reply to|use|find)\s+(?:the\s+)?(.+?)\s+(?:email|mail)\b/i,
    /\b(?:email|mail)\s+(?:about|regarding|on)\s+(.+?)$/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    const normalized = normalizeMailSearchQuery(match?.[1]);
    if (normalized) {
      return normalized;
    }
  }

  return undefined;
}

export function getCurrentMailDiscussionMessageId(): string | undefined {
  const mailTargetContext = deriveMcpTargetContextFromHie(
    hybridInteractionEngine.captureCurrentSourceContext(),
    hybridInteractionEngine.getCurrentTaskContext(),
    hybridInteractionEngine.getCurrentArtifacts()
  );
  return mailTargetContext?.mailItemId;
}

export function buildMailDiscussionReplyAllPlan(text: string): ICompoundWorkflowPlan | undefined {
  const normalizedText = text.trim().toLowerCase();
  if (!normalizedText || isHiePromptMessage(normalizedText)) {
    return undefined;
  }

  if (!canPlanMailDiscussionReplyAll()) {
    return undefined;
  }

  if (!looksLikeMailDiscussionReplyAllRequest(normalizedText)) {
    return undefined;
  }

  const currentMessageId = getCurrentMailDiscussionMessageId();
  const query = extractMailDiscussionSearchQuery(normalizedText);
  if (!currentMessageId && !query) {
    return undefined;
  }

  return {
    shouldPlan: true,
    familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
    confidence: 0.99,
    slots: {
      query: query || 'current mail discussion',
      selectionHint: 'none'
    },
    steps: [
      {
        id: 'resolve-mail-thread',
        kind: 'resolve-mail-thread',
        label: 'Resolve the mail discussion thread'
      },
      {
        id: 'compose-email',
        kind: 'compose-email',
        label: 'Open email draft'
      }
    ],
    label: 'Email visible recap to mail participants',
    userText: text
  };
}

function getMailDiscussionSearchQueryFromToolArgs(args: Record<string, unknown>): string | undefined {
  const rawQuery = typeof args.query === 'string'
    ? args.query
    : (typeof args.subject === 'string' ? args.subject : undefined);

  return normalizeMailSearchQuery(rawQuery);
}

function parseComposePrefill(args: Record<string, unknown>): Record<string, unknown> | undefined {
  const rawPrefill = args.prefill_json;
  if (!rawPrefill) {
    return undefined;
  }

  if (typeof rawPrefill === 'object') {
    return rawPrefill as Record<string, unknown>;
  }

  if (typeof rawPrefill === 'string') {
    try {
      const parsed = JSON.parse(rawPrefill) as Record<string, unknown>;
      return typeof parsed === 'object' && parsed ? parsed : undefined;
    } catch {
      return undefined;
    }
  }

  return undefined;
}

function looksLikeRecipientResolutionComposeRequest(text: string): boolean {
  return /\b(?:participants?|stakeholders?|authors?|everyone involved|all involved|people involved|person involved|reply all)\b/i.test(text);
}

function extractQuotedStrings(text: string): string[] {
  const results: string[] = [];
  const quotePattern = /["'`]([^"'`]{3,}?)["'`]/g;
  let match = quotePattern.exec(text);
  while (match) {
    results.push(match[1]);
    match = quotePattern.exec(text);
  }
  return results;
}

const COMPOSE_TOPIC_PATTERNS: ReadonlyArray<RegExp> = [
  /\b(?:participants?|stakeholders?|authors?)\s+(?:from|of)\s+(?:the\s+)?(.+?)(?:\s+(?:email|mail|document|thread|discussion|conversation)\b|[.,;]|$)/i,
  /\bbehind\s+(?:the\s+)?(.+?)(?:\s+(?:email|mail|document|thread|discussion|conversation)\b|[.,;]|$)/i,
  /\binvolved\s+(?:in|with)\s+(?:the\s+)?(.+?)(?:\s+(?:email|mail|document|thread|discussion|conversation)\b|[.,;]|$)/i
];

function extractMailDiscussionSearchQueryFromComposeToolArgs(args: Record<string, unknown>): string | undefined {
  const description = typeof args.description === 'string' ? args.description.trim() : '';
  const title = typeof args.title === 'string' ? args.title.trim() : '';
  const prefill = parseComposePrefill(args);
  const prefillSubject = typeof prefill?.subject === 'string' ? prefill.subject.trim() : '';

  const composeText = [description, title, prefillSubject].filter(Boolean).join(' ');
  if (!composeText || !looksLikeRecipientResolutionComposeRequest(composeText)) {
    return undefined;
  }

  // Tier 1: Quoted strings — LLM explicitly quoted the source topic
  const quotedStrings = extractQuotedStrings(description);
  for (let i = 0; i < quotedStrings.length; i++) {
    const normalized = normalizeMailSearchQuery(quotedStrings[i]);
    if (normalized) {
      return normalized;
    }
  }

  // Tier 2: Trigger-anchored patterns (participants from X, behind the X, involved in X)
  for (let i = 0; i < COMPOSE_TOPIC_PATTERNS.length; i++) {
    const match = composeText.match(COMPOSE_TOPIC_PATTERNS[i]);
    const normalized = normalizeMailSearchQuery(match?.[1]);
    if (normalized) {
      return normalized;
    }
  }

  // Tier 3: Use prefill subject directly — the LLM already picked a relevant subject
  return normalizeMailSearchQuery(prefillSubject);
}

export function resolveMailDiscussionReplyAllPlanFromToolCall(
  funcName: string | undefined,
  args: Record<string, unknown>
): ICompoundWorkflowPlan | undefined {
  if (funcName !== 'search_emails' && funcName !== 'show_compose_form' && funcName !== 'read_email_content') {
    return undefined;
  }

  if (!canPlanMailDiscussionReplyAll()) {
    return undefined;
  }

  const currentMessageId = getCurrentMailDiscussionMessageId();
  const query = funcName === 'search_emails' || funcName === 'read_email_content'
    ? getMailDiscussionSearchQueryFromToolArgs(args)
    : (() => {
        const preset = typeof args.preset === 'string' ? args.preset.trim() : '';
        if (preset !== 'email-compose') {
          return undefined;
        }
        return extractMailDiscussionSearchQueryFromComposeToolArgs(args);
      })();
  if (!currentMessageId && !query) {
    return undefined;
  }

  return {
    shouldPlan: true,
    familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
    confidence: 0.95,
    slots: {
      query: query || 'current mail discussion',
      selectionHint: 'none'
    },
    steps: [
      {
        id: 'resolve-mail-thread',
        kind: 'resolve-mail-thread',
        label: 'Resolve the mail discussion thread'
      },
      {
        id: 'compose-email',
        kind: 'compose-email',
        label: 'Open email draft'
      }
    ],
    label: 'Email visible recap to mail participants',
    userText: typeof args.query === 'string' ? args.query : (query || 'current mail discussion')
  };
}

export function setPendingMailDiscussionReplyAllPlan(plan: ICompoundWorkflowPlan | undefined): void {
  pendingMailDiscussionReplyAllWorkflow = plan
    ? {
        plan: clonePlan(plan),
        createdAt: Date.now()
      }
    : undefined;
}

export function clearPendingMailDiscussionReplyAllPlan(): void {
  pendingMailDiscussionReplyAllWorkflow = undefined;
}

export function hasPendingMailDiscussionReplyAllPlan(): boolean {
  return !!pendingMailDiscussionReplyAllWorkflow;
}

export function consumePendingMailDiscussionReplyAllPlan(options?: {
  latestUserMessageText?: string;
  requireHiePrompt?: boolean;
}): ICompoundWorkflowPlan | undefined {
  if (!pendingMailDiscussionReplyAllWorkflow) {
    return undefined;
  }

  if (options?.requireHiePrompt) {
    const latestUserMessageText = options.latestUserMessageText?.trim() || '';
    if (!latestUserMessageText || !isHiePromptMessage(latestUserMessageText)) {
      return undefined;
    }
  }

  if (!getCurrentMailDiscussionMessageId()) {
    return undefined;
  }

  const plan = clonePlan(pendingMailDiscussionReplyAllWorkflow.plan);
  pendingMailDiscussionReplyAllWorkflow = undefined;
  return plan;
}

export function resolveMailDiscussionReplyAllPlan(
  text: string | undefined,
  options?: {
    allowPending?: boolean;
    latestUserMessageText?: string;
    requireHiePromptForPending?: boolean;
  }
): ICompoundWorkflowPlan | undefined {
  if (options?.allowPending) {
    const pendingPlan = consumePendingMailDiscussionReplyAllPlan({
      latestUserMessageText: options.latestUserMessageText,
      requireHiePrompt: options.requireHiePromptForPending
    });
    if (pendingPlan) {
      return pendingPlan;
    }
  }

  if (!text) {
    return undefined;
  }

  return buildMailDiscussionReplyAllPlan(text);
}
