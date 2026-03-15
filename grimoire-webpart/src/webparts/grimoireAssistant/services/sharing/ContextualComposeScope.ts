import type {
  IBlock,
  IDocumentLibraryData,
  ISearchResultsData
} from '../../models/IBlock';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';

const COUNT_WORDS: Readonly<Record<string, number>> = {
  // English
  one: 1, two: 2, three: 3, four: 4, five: 5,
  six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
  // French
  un: 1, une: 1, deux: 2, trois: 3, quatre: 4, cinq: 5,
  sept: 7, huit: 8, neuf: 9, dix: 10,
  // Italian
  uno: 1, una: 1, due: 2, tre: 3, quattro: 4, cinque: 5,
  sei: 6, sette: 7, otto: 8, nove: 9, dieci: 10,
  // German
  eins: 1, zwei: 2, drei: 3, vier: 4,
  sechs: 6, sieben: 7, acht: 8, neun: 9, zehn: 10,
  // Spanish
  dos: 2, cuatro: 4, cinco: 5,
  seis: 6, siete: 7, ocho: 8, nueve: 9, diez: 10
};

const ORDINAL_INDEX_PATTERNS: ReadonlyArray<{ pattern: RegExp; index: number }> = [
  // en: first, fr: premier/première, it: primo/prima, de: erste/erster/erstes, es: primero/primera
  { pattern: /\b(?:first|premi(?:er|[eè]re)|prim[oa]|erst(?:e[rns]?)?|primer[oa]?)\b/i, index: 1 },
  { pattern: /\b(?:second|deuxi[eè]me|second[oa]|zweit(?:e[rns]?)?|segund[oa]?)\b/i, index: 2 },
  { pattern: /\b(?:third|troisi[eè]me|terz[oa]|dritt(?:e[rns]?)?|tercer[oa]?)\b/i, index: 3 },
  { pattern: /\b(?:fourth|quatri[eè]me|quart[oa]|viert(?:e[rns]?)?|cuart[oa]?)\b/i, index: 4 },
  { pattern: /\b(?:fifth|cinqui[eè]me|quint[oa]|f[uü]nft(?:e[rns]?)?|quint[oa]?)\b/i, index: 5 },
  { pattern: /\b(?:sixth|sixi[eè]me|sest[oa]|sechst(?:e[rns]?)?|sext[oa]?)\b/i, index: 6 },
  { pattern: /\b(?:seventh|septi[eè]me|settim[oa]|siebt(?:e[rns]?)?|s[eé]ptim[oa]?)\b/i, index: 7 },
  { pattern: /\b(?:eighth|huiti[eè]me|ottav[oa]|acht(?:e[rns]?)?|octav[oa]?)\b/i, index: 8 },
  { pattern: /\b(?:ninth|neuvi[eè]me|non[oa]|neunt(?:e[rns]?)?|noven[oa]?)\b/i, index: 9 },
  { pattern: /\b(?:tenth|dixi[eè]me|decim[oa]|zehnt(?:e[rns]?)?|d[eé]cim[oa]?)\b/i, index: 10 }
];

const ALL_SCOPE_PATTERNS: ReadonlyArray<RegExp> = [
  /\ball\b/i,
  /\bevery(?:thing|one)?\b/i,
  /\btou(?:s|tes)\b/i,          // fr: tous/toutes
  /\btutt[ie]\b/i,              // it: tutti/tutte
  /\balle[ns]?\b/i,             // de: alle/allen/alles
  /\btod[oa]s\b/i               // es: todos/todas
];

const SELECTED_SCOPE_PATTERNS: ReadonlyArray<RegExp> = [
  /\bselected\b/i,
  /\bchecked\b/i,
  /\bchosen\b/i,
  /\bhighlighted\b/i,
  /\bcurrent selection\b/i,
  /\bs[eé]lectionn[eé][es]?\b/i,   // fr: sélectionnés/sélectionnées
  /\bselezionat[ie]\b/i,            // it: selezionati/selezionate
  /\bausgew[aä]hlt(?:e[rns]?)?\b/i, // de: ausgewählt/ausgewählte
  /\bseleccionad[oa]s?\b/i          // es: seleccionado/seleccionados
];

type ParsedComposeScopeRequest =
  | { kind: 'all-visible' }
  | { kind: 'selected' }
  | { kind: 'range-from-start'; count: number }
  | { kind: 'range-from-end'; count: number }
  | { kind: 'indices'; indices: number[] };

export type ContextualComposeScopeMode =
  | 'single'
  | 'multiple'
  | 'selected'
  | 'all-visible';

export interface IResolvedContextualComposeScope {
  blockId?: string;
  selectedIndices?: number[];
  attachmentUris?: string[];
  itemTitle?: string;
  itemUrl?: string;
  scopeMode: ContextualComposeScopeMode;
  explicit: boolean;
  resolved: boolean;
}

export interface IContextualComposeScopeOptions {
  text?: string;
  blocks: IBlock[];
  activeBlockId?: string;
  selectedActionIndices?: number[];
}

// Static alternation of all count-word keys (must stay in sync with COUNT_WORDS above).
// eslint-disable-next-line @rushstack/security/no-unsafe-regexp
const COUNT_WORD_ALT = 'one|two|three|four|five|six|seven|eight|nine|ten|un|une|deux|trois|quatre|cinq|sept|huit|neuf|dix|uno|una|due|tre|quattro|cinque|sei|sette|otto|nove|dieci|eins|zwei|drei|vier|sechs|sieben|acht|neun|zehn|dos|cuatro|cinco|seis|siete|ocho|nueve|diez';

// en: first/top, fr: premiers/premières, it: primi/prime, de: ersten, es: primeros/primeras
const LEADING_COUNT_PATTERN = new RegExp( // eslint-disable-line @rushstack/security/no-unsafe-regexp
  '\\b(?:first|top|premi(?:ers?|[e\u00e8]res?)|prim[ie]|ersten?|primer[oa]s?)\\s+(\\d{1,2}|' + COUNT_WORD_ALT + ')\\b', 'i'
);

// en: last, fr: derniers/dernières, it: ultimi/ultime, de: letzten, es: últimos/últimas
const TRAILING_COUNT_PATTERN = new RegExp( // eslint-disable-line @rushstack/security/no-unsafe-regexp
  '\\b(?:last|derni(?:ers?|[e\u00e8]res?)|ultim[ie]|letzten?|[\u00faú]ltim[oa]s?)\\s+(\\d{1,2}|' + COUNT_WORD_ALT + ')\\b', 'i'
);

// Single "last" keyword (any language)
const LAST_KEYWORD_PATTERN = /\b(?:last|dernier|ultimo|letzte[rns]?|[uú]ltimo)\b/i;

function normalizeContextualText(text: string): string {
  return text
    .toLowerCase()
    .replace(/[\u201C\u201D\u00AB\u00BB\u300C\u300D\u300E\u300F]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function dedupePositiveIndices(indices: number[]): number[] {
  return Array.from(new Set(
    indices.filter((index) => Number.isFinite(index) && index > 0)
  ));
}

function isContextualVisibleItemBlock(block: IBlock): boolean {
  return block.type === 'search-results' || block.type === 'document-library';
}

function isAttachableFileUrl(url: string): boolean {
  if (!url.trim()) {
    return false;
  }

  try {
    const parsed = new URL(url);
    if (!/(?:sharepoint|onedrive)\.com$/i.test(parsed.hostname)) {
      return false;
    }

    return /\.[a-z0-9]{2,8}$/i.test(parsed.pathname);
  } catch {
    return false;
  }
}

function parseCountToken(rawValue: string | undefined): number | undefined {
  const normalized = rawValue?.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }

  if (/^\d{1,2}$/.test(normalized)) {
    const parsed = parseInt(normalized, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
  }

  return COUNT_WORDS[normalized];
}

function parseComposeScopeRequest(text: string): ParsedComposeScopeRequest | undefined {
  const normalizedText = normalizeContextualText(text);
  if (!normalizedText) {
    return undefined;
  }

  if (ALL_SCOPE_PATTERNS.some((pattern) => pattern.test(normalizedText))) {
    return { kind: 'all-visible' };
  }

  if (SELECTED_SCOPE_PATTERNS.some((pattern) => pattern.test(normalizedText))) {
    return { kind: 'selected' };
  }

  const leadingCountMatch = normalizedText.match(LEADING_COUNT_PATTERN);
  const leadingCount = parseCountToken(leadingCountMatch?.[1]);
  if (leadingCount) {
    return { kind: 'range-from-start', count: leadingCount };
  }

  const trailingCountMatch = normalizedText.match(TRAILING_COUNT_PATTERN);
  const trailingCount = parseCountToken(trailingCountMatch?.[1]);
  if (trailingCount) {
    return { kind: 'range-from-end', count: trailingCount };
  }

  if (LAST_KEYWORD_PATTERN.test(normalizedText)) {
    return { kind: 'range-from-end', count: 1 };
  }

  const explicitMatch = normalizedText.match(
    /\b(?:item|result|document|doc|file|number|#|r[eé]sultat|documento|fichier|datei|archivo|elemento|ergebnis|num[eé]ro)\s*(\d{1,2})(?:st|nd|rd|th|[eè]me|[oa])?\b/i
  ) || normalizedText.match(/\b(\d{1,2})(?:st|nd|rd|th|[eè]me|[oa])?\b/);
  if (explicitMatch) {
    const parsed = parseInt(explicitMatch[1], 10);
    if (Number.isFinite(parsed) && parsed > 0) {
      return { kind: 'indices', indices: [parsed] };
    }
  }

  for (let i = 0; i < ORDINAL_INDEX_PATTERNS.length; i++) {
    if (ORDINAL_INDEX_PATTERNS[i].pattern.test(normalizedText)) {
      return { kind: 'indices', indices: [ORDINAL_INDEX_PATTERNS[i].index] };
    }
  }

  return undefined;
}

function getPreferredContextualBlockIds(activeBlockId?: string): string[] {
  const sourceContext = hybridInteractionEngine.captureCurrentSourceContext();
  const taskContext = hybridInteractionEngine.getCurrentTaskContext();
  const candidates = [
    sourceContext?.sourceBlockId,
    taskContext?.sourceBlockId,
    taskContext?.derivedBlockId,
    activeBlockId
  ];

  return candidates.filter((value, index, values): value is string => (
    typeof value === 'string' && value.trim().length > 0 && values.indexOf(value) === index
  ));
}

function resolveContextualShareBlock(blocks: IBlock[], activeBlockId?: string): IBlock | undefined {
  const byId = new Map(blocks.map((block) => [block.id, block]));
  const preferredIds = getPreferredContextualBlockIds(activeBlockId);

  for (let i = 0; i < preferredIds.length; i++) {
    const block = byId.get(preferredIds[i]);
    if (block && isContextualVisibleItemBlock(block)) {
      return block;
    }
  }

  for (let i = blocks.length - 1; i >= 0; i--) {
    const block = blocks[i];
    if (isContextualVisibleItemBlock(block)) {
      return block;
    }
  }

  return undefined;
}

function getBlockItemCount(block: IBlock): number {
  if (block.type === 'search-results') {
    return (block.data as ISearchResultsData).results.length;
  }

  if (block.type === 'document-library') {
    return (block.data as IDocumentLibraryData).items.length;
  }

  return 0;
}

function getBlockItemTitle(block: IBlock, index: number): string | undefined {
  const zeroBased = index - 1;
  if (zeroBased < 0) {
    return undefined;
  }

  if (block.type === 'search-results') {
    return (block.data as ISearchResultsData).results[zeroBased]?.title;
  }

  if (block.type === 'document-library') {
    return (block.data as IDocumentLibraryData).items[zeroBased]?.name;
  }

  return undefined;
}

function getBlockItemUrl(block: IBlock, index: number): string | undefined {
  const zeroBased = index - 1;
  if (zeroBased < 0) {
    return undefined;
  }

  if (block.type === 'search-results') {
    return (block.data as ISearchResultsData).results[zeroBased]?.url;
  }

  if (block.type === 'document-library') {
    return (block.data as IDocumentLibraryData).items[zeroBased]?.url;
  }

  return undefined;
}

function getBlockAttachmentUris(block: IBlock, indices: number[]): string[] {
  const uris = indices
    .map((index) => {
      const zeroBased = index - 1;
      if (zeroBased < 0) {
        return undefined;
      }

      if (block.type === 'search-results') {
        const item = (block.data as ISearchResultsData).results[zeroBased];
        return item?.url;
      }

      if (block.type === 'document-library') {
        const item = (block.data as IDocumentLibraryData).items[zeroBased];
        if (!item || item.type !== 'file') {
          return undefined;
        }
        return item.url;
      }

      return undefined;
    })
    .filter((url): url is string => typeof url === 'string' && isAttachableFileUrl(url));

  return Array.from(new Set(uris));
}

function getSelectedIndicesFromUiState(selectedActionIndices?: number[]): number[] | undefined {
  const sourceContext = hybridInteractionEngine.captureCurrentSourceContext();
  const taskContext = hybridInteractionEngine.getCurrentTaskContext();
  const indices = dedupePositiveIndices([
    ...(sourceContext?.selectedItems || []).map((item) => item.index || 0),
    ...(taskContext?.selectedItems || []).map((item) => item.index || 0),
    ...(selectedActionIndices || [])
  ]);

  return indices.length > 0 ? indices : undefined;
}

function resolveRequestedIndices(
  request: ParsedComposeScopeRequest,
  block: IBlock,
  selectedActionIndices?: number[]
): number[] | undefined {
  const itemCount = getBlockItemCount(block);
  if (itemCount <= 0) {
    return undefined;
  }

  if (request.kind === 'all-visible') {
    return Array.from({ length: itemCount }, (_, index) => index + 1);
  }

  if (request.kind === 'selected') {
    const selectedIndices = getSelectedIndicesFromUiState(selectedActionIndices);
    if (!selectedIndices) {
      return undefined;
    }

    return selectedIndices.filter((index) => index <= itemCount);
  }

  if (request.kind === 'range-from-start') {
    return Array.from({ length: Math.min(request.count, itemCount) }, (_, index) => index + 1);
  }

  if (request.kind === 'range-from-end') {
    const count = Math.min(request.count, itemCount);
    const start = Math.max(1, itemCount - count + 1);
    return Array.from({ length: count }, (_, index) => start + index);
  }

  return request.indices.filter((index) => index <= itemCount);
}

function inferScopeMode(request: ParsedComposeScopeRequest, resolvedCount: number): ContextualComposeScopeMode {
  if (request.kind === 'all-visible') {
    return 'all-visible';
  }

  if (request.kind === 'selected') {
    return 'selected';
  }

  return resolvedCount <= 1 ? 'single' : 'multiple';
}

export function resolveContextualComposeScope(
  options: IContextualComposeScopeOptions
): IResolvedContextualComposeScope | undefined {
  const request = parseComposeScopeRequest(options.text || '');
  if (!request) {
    return undefined;
  }

  const block = resolveContextualShareBlock(options.blocks, options.activeBlockId);
  if (!block) {
    return {
      scopeMode: inferScopeMode(request, 0),
      explicit: true,
      resolved: false
    };
  }

  const selectedIndices = dedupePositiveIndices(
    resolveRequestedIndices(request, block, options.selectedActionIndices) || []
  );
  const itemUrl = selectedIndices.length === 1 ? getBlockItemUrl(block, selectedIndices[0]) : undefined;
  const itemTitle = selectedIndices.length === 1 ? getBlockItemTitle(block, selectedIndices[0]) : undefined;

  return {
    blockId: block.id,
    selectedIndices: selectedIndices.length > 0 ? selectedIndices : undefined,
    attachmentUris: selectedIndices.length > 0 ? getBlockAttachmentUris(block, selectedIndices) : undefined,
    itemTitle,
    itemUrl,
    scopeMode: inferScopeMode(request, selectedIndices.length),
    explicit: true,
    resolved: selectedIndices.length > 0
  };
}

function parseShareSelectionIndices(staticArgs: Record<string, unknown>): number[] | undefined {
  const raw = staticArgs.shareSelectionIndices;
  if (!Array.isArray(raw)) {
    return undefined;
  }

  const parsed = dedupePositiveIndices(
    raw.map((value) => typeof value === 'number' ? value : parseInt(String(value), 10))
  );
  return parsed.length > 0 ? parsed : undefined;
}

export function readComposeShareScope(staticArgs: Record<string, unknown>): IResolvedContextualComposeScope | undefined {
  const explicit = staticArgs.shareScopeExplicit === true
    || typeof staticArgs.shareScopeMode === 'string'
    || Array.isArray(staticArgs.shareSelectionIndices);
  if (!explicit) {
    return undefined;
  }

  const parsedMode = typeof staticArgs.shareScopeMode === 'string'
    ? staticArgs.shareScopeMode
    : undefined;
  const selectedIndices = parseShareSelectionIndices(staticArgs);
  const attachmentUris = Array.isArray(staticArgs.attachmentUris)
    ? staticArgs.attachmentUris.filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    : undefined;
  const resolved = staticArgs.shareScopeResolved === false
    ? false
    : !!(selectedIndices && selectedIndices.length > 0);
  const inferredMode: ContextualComposeScopeMode = (
    parsedMode === 'single'
    || parsedMode === 'multiple'
    || parsedMode === 'selected'
    || parsedMode === 'all-visible'
  )
    ? parsedMode
    : ((selectedIndices?.length || 0) <= 1 ? 'single' : 'multiple');

  return {
    blockId: typeof staticArgs.shareBlockId === 'string' ? staticArgs.shareBlockId : undefined,
    selectedIndices,
    attachmentUris,
    itemTitle: typeof staticArgs.shareItemTitle === 'string' ? staticArgs.shareItemTitle : undefined,
    itemUrl: typeof staticArgs.fileOrFolderUrl === 'string' ? staticArgs.fileOrFolderUrl : undefined,
    scopeMode: inferredMode,
    explicit: true,
    resolved
  };
}

export function buildComposeStaticArgsFromScope(scope: IResolvedContextualComposeScope): Record<string, unknown> {
  const nextStaticArgs: Record<string, unknown> = {
    shareScopeMode: scope.scopeMode,
    shareScopeExplicit: scope.explicit,
    shareScopeResolved: scope.resolved
  };

  if (scope.blockId) {
    nextStaticArgs.shareBlockId = scope.blockId;
  }
  if (scope.selectedIndices && scope.selectedIndices.length > 0) {
    nextStaticArgs.shareSelectionIndices = scope.selectedIndices;
  }
  if (scope.attachmentUris && scope.attachmentUris.length > 0) {
    nextStaticArgs.attachmentUris = scope.attachmentUris;
  }
  if (scope.itemTitle) {
    nextStaticArgs.shareItemTitle = scope.itemTitle;
    nextStaticArgs.fileOrFolderName = scope.itemTitle;
  }
  if (scope.itemUrl) {
    nextStaticArgs.fileOrFolderUrl = scope.itemUrl;
  }

  return nextStaticArgs;
}
