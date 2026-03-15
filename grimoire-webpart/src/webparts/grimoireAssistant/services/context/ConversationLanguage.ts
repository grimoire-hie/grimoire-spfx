import { detectQueryLanguage, normalizeLanguageTag } from '../search/SearchLanguageUtils';

export type ConversationLanguage = 'en' | 'fr' | 'it' | 'de' | 'es';

const SUPPORTED_CONVERSATION_LANGUAGES = new Set([
  'en',
  'fr',
  'it',
  'de',
  'es'
] as const);

const LANGUAGE_LABELS: Record<ConversationLanguage, string> = {
  en: 'English',
  fr: 'French',
  it: 'Italian',
  de: 'German',
  es: 'Spanish'
};

const EXPLICIT_LANGUAGE_SWITCH_PATTERNS: ReadonlyArray<{
  language: ConversationLanguage;
  patterns: ReadonlyArray<RegExp>;
}> = [
  {
    language: 'en',
    patterns: [
      /\b(?:speak|reply|answer|respond|continue)\s+in\s+english\b/i,
      /\b(?:let'?s|lets)\s+(?:speak|continue)\s+in\s+english\b/i,
      /\benglish\s+please\b/i,
      /\bin english\b/i
    ]
  },
  {
    language: 'fr',
    patterns: [
      /\b(?:speak|reply|answer|respond|continue)\s+in\s+french\b/i,
      /\b(?:parle|parler|réponds|répondre|continue|continuons)\s+(?:en\s+)?fran[cç]ais\b/i,
      /\bfran[cç]ais\b/i
    ]
  },
  {
    language: 'it',
    patterns: [
      /\b(?:speak|reply|answer|respond|continue)\s+in\s+italian\b/i,
      /\b(?:parla|parliamo|rispondi|rispondere|continua|continuiamo)\s+(?:in\s+)?italiano\b/i,
      /\bitaliano\b/i
    ]
  },
  {
    language: 'de',
    patterns: [
      /\b(?:speak|reply|answer|respond|continue)\s+in\s+german\b/i,
      /\b(?:sprich|spreche|antwort[e]?)\s+(?:auf\s+)?deutsch\b/i,
      /\bdeutsch\b/i
    ]
  },
  {
    language: 'es',
    patterns: [
      /\b(?:speak|reply|answer|respond|continue)\s+in\s+spanish\b/i,
      /\b(?:habla|hablemos|responde|responder|contin[aú]a)\s+(?:en\s+)?espa[nñ]ol\b/i,
      /\bespa[nñ]ol\b/i
    ]
  }
];

export function normalizeConversationLanguage(language?: string): ConversationLanguage | undefined {
  const normalized = normalizeLanguageTag(language);
  if (!normalized || !SUPPORTED_CONVERSATION_LANGUAGES.has(normalized as ConversationLanguage)) {
    return undefined;
  }
  return normalized as ConversationLanguage;
}

export function getConversationLanguageLabel(language?: string): string {
  const normalized = normalizeConversationLanguage(language) || 'en';
  return LANGUAGE_LABELS[normalized];
}

export function detectExplicitConversationLanguageSwitch(text: string): ConversationLanguage | undefined {
  const trimmed = text.trim();
  if (!trimmed) return undefined;

  for (let i = 0; i < EXPLICIT_LANGUAGE_SWITCH_PATTERNS.length; i++) {
    const candidate = EXPLICIT_LANGUAGE_SWITCH_PATTERNS[i];
    for (let j = 0; j < candidate.patterns.length; j++) {
      if (candidate.patterns[j].test(trimmed)) {
        return candidate.language;
      }
    }
  }

  return undefined;
}

export function resolveConversationLanguage(
  text: string,
  currentLanguage?: string,
  fallbackLanguage?: string
): ConversationLanguage {
  const explicit = detectExplicitConversationLanguageSwitch(text);
  if (explicit) {
    return explicit;
  }

  const normalizedCurrent = normalizeConversationLanguage(currentLanguage);
  const normalizedFallback = normalizeConversationLanguage(fallbackLanguage) || 'en';
  const detected = normalizeConversationLanguage(
    detectQueryLanguage(text, normalizedCurrent || normalizedFallback)
  ) || normalizedCurrent || normalizedFallback;

  if (normalizedCurrent && normalizedCurrent !== 'en') {
    if (detected !== 'en' && detected !== normalizedCurrent) {
      return detected;
    }
    return normalizedCurrent;
  }

  return detected;
}

export function buildConversationLanguageContextMessage(language: string): string {
  const normalized = normalizeConversationLanguage(language) || 'en';
  const label = LANGUAGE_LABELS[normalized];
  return `[Conversation preference: Reply in ${label} until the user explicitly switches language.]`;
}
