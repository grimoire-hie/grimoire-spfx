import type { HieTurnStartMode } from './HIETypes';

export interface IHieIngressTurnPolicyOptions {
  hasTaskContext: boolean;
  hasVisibleBlocks: boolean;
  visibleBlockTitles?: string[];
  visibleReferenceTitles?: string[];
}

export interface IHieIngressTurnPolicyDecision {
  mode: HieTurnStartMode;
  reason: string;
}

const EXPLICIT_RESET_PATTERNS: ReadonlyArray<RegExp> = [
  /\bnew topic\b/i,
  /\bstart over\b/i,
  /\bforget that\b/i,
  /\bsomething else\b/i,
  /\bdifferent topic\b/i,
  /\bchange subject\b/i
];

const CONTEXTUAL_REFERENCE_PATTERNS: ReadonlyArray<RegExp> = [
  /^(?:\d+|first|second|third|fourth|fifth|sixth|last)$/i,
  /^\s*what about\b/i,
  /\b(?:summarize|preview|open|read|share|send|email|mail|recap|chat about|explain)\b.*\b(?:this|that|these|those|it|them|results?|documents?|docs?|files?|emails?|items?)\b/i,
  /\b(?:this|that|these|those|it|them|results?|documents?|docs?|files?|emails?|items?)\b.*\b(?:summarize|preview|open|read|share|send|email|mail|recap|chat about|explain)\b/i,
  /\b(?:document|doc|item|file|result|email|message)\s+#?\d+\b/i,
  /\b(?:document|doc|item|file|result|email|message)\s+(?:first|second|third|fourth|fifth|sixth|last)\b/i
];

const TOP_LEVEL_REQUEST_PATTERNS: ReadonlyArray<RegExp> = [
  /^\s*(?:search|find|look up|lookup|research|browse|list|show)\b/i,
  /^\s*(?:open|read|inspect|review)\b/i,
  /^\s*(?:i am searching for|i'm searching for|i need|help me find)\b/i
];

const TITLE_MATCH_ACTION_PATTERNS: ReadonlyArray<RegExp> = [
  /\bsummarize\b/i,
  /\bsummary\b/i,
  /\brecap\b/i,
  /\bpreview\b/i,
  /\bopen\b/i,
  /\bread\b/i,
  /\bshow\b/i,
  /\bdetails?\b/i,
  /\bexplain\b/i,
  /\bchat about\b/i,
  /\bshare\b/i,
  /\bsend\b/i,
  /\bemail\b/i,
  /\bmail\b/i
];

function normalizeMatchText(value: string): string {
  return value
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function hasVisibleTitleMatch(
  text: string,
  titles: ReadonlyArray<string> | undefined
): boolean {
  if (!titles || titles.length === 0) {
    return false;
  }

  const normalizedText = normalizeMatchText(text);
  if (!normalizedText) {
    return false;
  }

  return titles.some((title) => {
    const normalizedTitle = normalizeMatchText(title);
    return normalizedTitle.length >= 3 && normalizedText.includes(normalizedTitle);
  });
}

function matchesAny(text: string, patterns: ReadonlyArray<RegExp>): boolean {
  return patterns.some((pattern) => pattern.test(text));
}

export function resolveIngressTurnStartPolicy(
  text: string,
  options: IHieIngressTurnPolicyOptions
): IHieIngressTurnPolicyDecision {
  const normalized = text.trim();
  const hasContext = options.hasTaskContext || options.hasVisibleBlocks;

  if (!normalized) {
    return { mode: 'auto', reason: 'empty-user-turn' };
  }

  if (matchesAny(normalized, EXPLICIT_RESET_PATTERNS)) {
    return { mode: 'new-root', reason: 'explicit-reset-phrase' };
  }

  if (!hasContext) {
    return { mode: 'new-root', reason: 'no-active-context' };
  }

  if (matchesAny(normalized, CONTEXTUAL_REFERENCE_PATTERNS)) {
    return { mode: 'inherit', reason: 'contextual-follow-up' };
  }

  if (
    matchesAny(normalized, TITLE_MATCH_ACTION_PATTERNS)
    && (
      hasVisibleTitleMatch(normalized, options.visibleReferenceTitles)
      || hasVisibleTitleMatch(normalized, options.visibleBlockTitles)
    )
  ) {
    return { mode: 'inherit', reason: 'visible-title-follow-up' };
  }

  if (matchesAny(normalized, TOP_LEVEL_REQUEST_PATTERNS)) {
    return { mode: 'new-root', reason: 'top-level-request' };
  }

  return { mode: 'auto', reason: 'defer-to-hie-auto' };
}
