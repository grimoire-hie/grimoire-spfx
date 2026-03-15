export type ExplicitPersonalOneDriveIntentKind =
  | 'browse-root'
  | 'search-by-name'
  | 'unsupported-filter';

export interface IExplicitPersonalOneDriveIntent {
  kind: ExplicitPersonalOneDriveIntentKind;
  searchQuery?: string;
}

const PERSONAL_ONEDRIVE_HINT = /\bmy\s+(?:personal\s+)?one\s*drive\b/i;
const PERSONAL_ONEDRIVE_FILES_HINT = /\b(?:file|files|folder|folders|content|document|documents)\b/i;
const BROWSE_HINT = /\b(?:show|list|browse|open)\b/i;
const NAME_SEARCH_PATTERNS: ReadonlyArray<RegExp> = [
  /\b(?:files?|filenames?)\s+(?:starting|beginning)\s+with\s+["']?([^"',.;!?]+)["']?/i,
  /\b(?:start|starts|starting|beginning)\s+with\s+["']?([^"',.;!?]+)["']?/i,
  /\b(?:named|called)\s+["']?([^"',.;!?]+)["']?/i,
  /\bfind\s+(.+?)\s+in\s+my\s+(?:personal\s+)?one\s*drive\b/i,
  /\bsearch\s+my\s+(?:personal\s+)?one\s*drive\s+for\s+(.+?)$/i,
  /\blook\s+up\s+(.+?)\s+in\s+my\s+(?:personal\s+)?one\s*drive\b/i
];
const UNSUPPORTED_FILTER_PATTERNS: ReadonlyArray<RegExp> = [
  /\b(?:recent|recently|latest|newest|oldest)\b/i,
  /\b(?:today|yesterday)\b/i,
  /\blast\s+(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+(?:day|days|week|weeks|month|months|hour|hours)\b/i,
  /\b(?:created|modified|opened|used)\b.*\b(?:last|recent|today|yesterday)\b/i,
  /\b(?:pdfs?|docx?|xlsx?|pptx?|csv|txt|json|images?|photos?)\b/i,
  /\bfile\s+type\b/i
];

function normalizeSearchQuery(value: string): string | undefined {
  const trimmed = value
    .trim()
    .replace(/^["']+|["']+$/g, '')
    .replace(/\b(?:files?|filenames?|folders?|documents?)\b$/i, '')
    .trim();
  return trimmed || undefined;
}

function extractSearchQuery(text: string): string | undefined {
  for (let i = 0; i < NAME_SEARCH_PATTERNS.length; i++) {
    const match = text.match(NAME_SEARCH_PATTERNS[i]);
    const captured = match?.[1];
    const normalized = captured ? normalizeSearchQuery(captured) : undefined;
    if (normalized) {
      return normalized;
    }
  }
  return undefined;
}

function hasUnsupportedFilter(text: string): boolean {
  for (let i = 0; i < UNSUPPORTED_FILTER_PATTERNS.length; i++) {
    if (UNSUPPORTED_FILTER_PATTERNS[i].test(text)) {
      return true;
    }
  }
  return false;
}

export function classifyExplicitPersonalOneDriveIntent(text: string): IExplicitPersonalOneDriveIntent | undefined {
  const normalizedText = text.trim();
  if (!normalizedText || !PERSONAL_ONEDRIVE_HINT.test(normalizedText)) {
    return undefined;
  }

  if (hasUnsupportedFilter(normalizedText)) {
    return { kind: 'unsupported-filter' };
  }

  const searchQuery = extractSearchQuery(normalizedText);
  if (searchQuery) {
    return {
      kind: 'search-by-name',
      searchQuery
    };
  }

  if (BROWSE_HINT.test(normalizedText) && PERSONAL_ONEDRIVE_FILES_HINT.test(normalizedText)) {
    return { kind: 'browse-root' };
  }

  return undefined;
}
