const STOPWORDS: Record<string, string[]> = {
  en: ['the', 'and', 'for', 'with', 'from', 'last', 'this', 'that', 'into', 'about'],
  fr: ['le', 'la', 'les', 'de', 'des', 'pour', 'avec', 'dans', 'sur', 'une', 'un'],
  it: ['il', 'lo', 'la', 'gli', 'dei', 'delle', 'per', 'con', 'nel', 'una', 'uno'],
  de: ['der', 'die', 'das', 'und', 'mit', 'für', 'von', 'letzte', 'eine', 'einer'],
  es: ['el', 'la', 'los', 'las', 'de', 'para', 'con', 'una', 'uno', 'sobre']
};

const LCID_BY_LANGUAGE: Record<string, number> = {
  ar: 1025,
  de: 1031,
  en: 1033,
  es: 3082,
  fr: 1036,
  it: 1040,
  ja: 1041,
  ko: 1042,
  nl: 1043,
  pl: 1045,
  pt: 1046,
  ru: 1049,
  sv: 1053,
  tr: 1055,
  uk: 1058,
  zh: 2052
};

const LANGUAGE_BY_LCID: Record<string, string> = Object.keys(LCID_BY_LANGUAGE).reduce<Record<string, string>>((acc, key) => {
  acc[String(LCID_BY_LANGUAGE[key])] = key;
  return acc;
}, {});

function countStopwordHits(tokens: string[], language: string): number {
  const stopwords = STOPWORDS[language];
  if (!stopwords || tokens.length === 0) return 0;
  let score = 0;
  for (let i = 0; i < tokens.length; i++) {
    if (stopwords.indexOf(tokens[i]) !== -1) score++;
  }
  return score;
}

export function normalizeLanguageTag(language: string | undefined): string | undefined {
  if (!language) return undefined;
  const trimmed = language.trim();
  if (!trimmed) return undefined;
  if (/^\d+$/.test(trimmed)) {
    return LANGUAGE_BY_LCID[trimmed];
  }
  const normalized = trimmed.replace('_', '-').toLowerCase();
  if (normalized === 'zh-cn' || normalized === 'zh-hans') return 'zh';
  if (normalized === 'zh-tw' || normalized === 'zh-hant') return 'zh';
  return normalized.split('-')[0];
}

export function getLcidForLanguage(language: string | undefined): number {
  const normalized = normalizeLanguageTag(language);
  if (!normalized) return LCID_BY_LANGUAGE.en;
  return LCID_BY_LANGUAGE[normalized] || LCID_BY_LANGUAGE.en;
}

export function detectQueryLanguage(query: string, fallbackLanguage?: string): string {
  const normalizedFallback = normalizeLanguageTag(fallbackLanguage) || 'en';
  const trimmed = query.trim();
  if (!trimmed) return normalizedFallback;

  if (/[\u0600-\u06FF]/.test(trimmed)) return 'ar';
  if (/[\u3040-\u30FF]/.test(trimmed)) return 'ja';
  if (/[\u4E00-\u9FFF]/.test(trimmed)) return 'zh';
  if (/[\uAC00-\uD7AF]/.test(trimmed)) return 'ko';
  if (/[\u0400-\u04FF]/.test(trimmed)) return 'ru';

  const tokens = trimmed
    .toLowerCase()
    .split(/[^a-zA-Z\u00c0-\u024f\u0400-\u04ff\u0600-\u06ff\u0590-\u05ff\u0e00-\u0e7f]+/)
    .map((token) => token.trim())
    .filter(Boolean);

  let bestLanguage = normalizedFallback;
  let bestScore = 0;
  Object.keys(STOPWORDS).forEach((language) => {
    const score = countStopwordHits(tokens, language);
    if (score > bestScore) {
      bestScore = score;
      bestLanguage = language;
    }
  });

  if (bestScore > 0) return bestLanguage;
  if (/[àâçéèêëîïôùûüÿœ]/i.test(trimmed)) return 'fr';
  if (/[äöüß]/i.test(trimmed)) return 'de';
  if (/[ìòàù]/i.test(trimmed)) return 'it';
  if (/[ñáéíóú]/i.test(trimmed)) return 'es';

  return normalizedFallback;
}

export function shouldUseTranslationFallback(queryLanguage: string, userLanguage?: string): string | undefined {
  const normalizedQueryLanguage = normalizeLanguageTag(queryLanguage) || 'en';
  const normalizedUserLanguage = normalizeLanguageTag(userLanguage);

  if (normalizedUserLanguage && normalizedUserLanguage !== normalizedQueryLanguage) {
    return normalizedUserLanguage;
  }
  if (normalizedQueryLanguage !== 'en') {
    return 'en';
  }
  return undefined;
}
