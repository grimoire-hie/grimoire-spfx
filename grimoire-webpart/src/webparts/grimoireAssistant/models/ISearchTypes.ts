/**
 * ISearchTypes — Types for semantic search, planning, and fusion.
 */

// ─── Copilot Search API (/beta/copilot/search) ────────────────────

export interface ICopilotSearchRequest {
  query: string;
  pageSize?: number; // 1-100, default 25
  dataSources?: {
    oneDrive?: {
      filterExpression?: string;
      resourceMetadataNames?: string[];
    };
  };
}

/** Raw hit from the Copilot Search API */
export interface ICopilotSearchHit {
  webUrl: string;
  preview: string;
  resourceType: string;
  resourceMetadata?: Record<string, string>;
}

/** Raw response from the Copilot Search API */
export interface ICopilotSearchApiResponse {
  totalCount: number;
  searchHits: ICopilotSearchHit[];
  '@odata.nextLink'?: string;
}

// ─── Copilot Retrieval API (/beta/copilot/retrieval) ──────────────

export type RetrievalDataSource = 'sharePoint' | 'oneDriveBusiness' | 'externalItem';

export interface ICopilotRetrievalRequest {
  queryString: string;
  dataSource: RetrievalDataSource;
  maximumNumberOfResults?: number; // 1-25, default 10
  resourceMetadata?: string[];
  filterExpression?: string;
}

export interface IRetrievalExtract {
  text: string;
  relevanceScore?: number;
}

/** Raw hit from the Copilot Retrieval API */
export interface ICopilotRetrievalHit {
  webUrl: string;
  resourceReference?: {
    id: string;
    type: string;
  };
  resourceType?: 'listItem' | 'externalItem';
  resourceMetadata?: Record<string, string>;
  extracts?: IRetrievalExtract[];
}

export interface ICopilotRetrievalApiResponse {
  retrievalHits: ICopilotRetrievalHit[];
}

// ─── Search Planning ──────────────────────────────────────────────

export type SearchQueryVariantKind =
  | 'raw'
  | 'semantic-rewrite'
  | 'corrected'
  | 'translation'
  | 'keyword-fallback';

export interface ISearchIntentPlan {
  rawQuery: string;
  semanticRewriteQuery?: string;
  semanticRewriteConfidence?: number;
  sharePointLexicalQuery?: string;
  sharePointLexicalConfidence?: number;
  correctedQuery?: string;
  queryLanguage: string;
  translationFallbackQuery?: string;
  translationFallbackLanguage?: string;
  keywordFallbackQuery?: string;
  usedCorrection: boolean;
  usedTranslation: boolean;
}

// ─── Normalized Result (unified across all search APIs) ───────────

export type SearchResultSource =
  | 'copilot-search'
  | 'copilot-retrieval'
  | 'sharepoint-search'
  | 'insights';

export interface ICopilotSearchResult {
  /** Document title */
  title: string;
  /** Preview snippet */
  summary: string;
  /** Full document URL */
  url: string;
  /** File type (docx, xlsx, pptx, pdf, etc.) */
  fileType?: string;
  /** Last modified timestamp */
  lastModified?: string;
  /** Author display name */
  author?: string;
  /** Site name where the document lives */
  siteName?: string;
  /** Relevance score (raw API value when available) */
  score?: number;
  /** Source system that produced the result */
  source: SearchResultSource | string;
  /** 1-based position within the source result set */
  sourceRank?: number;
  /** Native score/rank returned by the source when available */
  sourceNativeScore?: number;
  /** Query text used for this result set */
  queryText?: string;
  /** Which query variant produced the hit */
  variantKind?: SearchQueryVariantKind;
  /** Language used for the query variant */
  variantLanguage?: string;
  /** Document/content language when the API returns it */
  language?: string;
}

// ─── RRF Fusion ───────────────────────────────────────────────────

export interface IRRFResult {
  /** Document URL (dedup key) */
  url: string;
  /** Document title */
  title: string;
  /** Best summary across sources */
  summary: string;
  /** Fused RRF score */
  rrfScore: number;
  /** Which sources contributed to this result */
  sources: string[];
  /** Which query variants contributed to this result */
  variantKinds: SearchQueryVariantKind[];
  /** Languages of the contributing query variants */
  variantLanguages: string[];
  /** Original result details */
  fileType?: string;
  lastModified?: string;
  author?: string;
  siteName?: string;
  language?: string;
}

// ─── API Response Wrapper ─────────────────────────────────────────

export interface ISearchApiError {
  code: string;
  message: string;
  innerError?: {
    code?: string;
    message?: string;
  };
}

export interface ISearchApiResponse<T> {
  success: boolean;
  data?: T;
  error?: ISearchApiError;
  durationMs?: number;
}
