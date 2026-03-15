import type { ISearchQuery, ISearchResult } from '@pnp/sp/search';
import type { ICopilotSearchResult, ISearchApiResponse, SearchQueryVariantKind } from '../../models/ISearchTypes';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { getSP } from '../pnp/pnpConfig';
import { logService } from '../logging/LogService';
import { getLcidForLanguage, normalizeLanguageTag } from './SearchLanguageUtils';

const DOCUMENTS_SOURCE_ID = 'e7ec8cee-ded8-43c9-beb5-436b54b31e84';

export interface ISharePointSearchOptions {
  maxResults?: number;
  queryLanguage?: string;
  variantKind?: SearchQueryVariantKind;
  variantLanguage?: string;
  requestQueryText?: string;
}

function normalizeMatchText(value: string): string {
  return value
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^0-9a-z\u00c0-\u024f\u0400-\u04ff\u0600-\u06ff\u0590-\u05ff\u3040-\u30ff\u4e00-\u9fff\uac00-\ud7af\u0e00-\u0e7f]+/gi, ' ')
    .trim();
}

function tokenizeLexicalQuery(queryText: string): string[] {
  const normalized = normalizeMatchText(queryText);
  if (!normalized) return [];
  return normalized
    .split(/\s+/)
    .map((term) => term.trim())
    .filter((term) => term.length >= 2);
}

function containsLexicalTerm(text: string | undefined, term: string): boolean {
  if (!text) return false;
  const normalized = normalizeMatchText(text);
  if (!normalized) return false;
  return ` ${normalized} `.includes(` ${term} `);
}

function extractUrlMatchText(url: string): string {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    return decodeURIComponent(parsed.pathname || '');
  } catch {
    return url.replace(/^https?:\/\/[^/]+/i, '');
  }
}

function cleanSummary(summary: string | undefined): string {
  if (!summary) return '';
  return summary
    .replace(/<c\d+>/g, '')
    .replace(/<\/c\d+>/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeDate(value: Date | string | undefined): string | undefined {
  if (!value) return undefined;
  if (value instanceof Date) return value.toISOString();
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? undefined : parsed.toISOString();
}

function normalizeAuthor(author: string | undefined): string | undefined {
  if (!author) return undefined;
  const trimmed = author.trim();
  if (!trimmed) return undefined;
  return trimmed.split('|')[0].trim();
}

function extractSiteName(result: ISearchResult, url: string): string | undefined {
  if (result.SiteName) return result.SiteName;
  try {
    const parsed = new URL(url);
    const parts = parsed.pathname.split('/').filter(Boolean);
    if (parts.length >= 2 && (parts[0] === 'sites' || parts[0] === 'teams')) {
      return decodeURIComponent(parts[1]);
    }
    return parsed.hostname;
  } catch {
    return undefined;
  }
}

function extractTitle(result: ISearchResult, url: string): string {
  if (result.Title && result.Title.trim()) return result.Title.trim();
  try {
    const parts = new URL(url).pathname.split('/');
    const last = parts[parts.length - 1] || url;
    return decodeURIComponent(last).replace(/\.[^.]+$/, '');
  } catch {
    return url;
  }
}

export class SharePointSearchService {
  public async search(
    query: string,
    options: ISharePointSearchOptions = {}
  ): Promise<ISearchApiResponse<ICopilotSearchResult[]>> {
    const startTime = performance.now();
    const trimmedQuery = query.trim();
    const searchTuning = getRuntimeTuningConfig().search;
    const maxResults = Math.min(Math.max(1, options.maxResults || searchTuning.sharePointMaxResultsCap), searchTuning.sharePointMaxResultsCap);
    const queryLanguage = normalizeLanguageTag(options.queryLanguage) || 'en';
    const lcid = getLcidForLanguage(queryLanguage);
    const queryText = (options.requestQueryText || trimmedQuery).trim() || trimmedQuery;
    const lexicalTerms = tokenizeLexicalQuery(queryText);

    logService.info(
      'search',
      queryText === trimmedQuery
        ? `SharePoint Search: "${trimmedQuery}" (max: ${maxResults}, lang: ${queryLanguage})`
        : `SharePoint Search: "${trimmedQuery}" → "${queryText}" (max: ${maxResults}, lang: ${queryLanguage})`,
      JSON.stringify({
        rawQuery: trimmedQuery,
        requestQueryText: queryText,
        lexicalTerms,
        variantKind: options.variantKind || 'raw',
        variantLanguage: options.variantLanguage || queryLanguage
      }, null, 2)
    );

    const request: ISearchQuery = {
      Querytext: queryText,
      SourceId: DOCUMENTS_SOURCE_ID,
      TrimDuplicates: true,
      EnableQueryRules: false,
      EnableSorting: true,
      RowLimit: maxResults,
      RowsPerPage: maxResults,
      Culture: lcid,
      ClientType: 'GrimoireSemanticSearch',
      DesiredSnippetLength: 180,
      MaxSnippetLength: 240,
      SummaryLength: 180,
      SelectProperties: [
        'Title',
        'Path',
        'OriginalPath',
        'Description',
        'HitHighlightedSummary',
        'Author',
        'LastModifiedTime',
        'Write',
        'FileExtension',
        'FileType',
        'SiteName',
        'Culture',
        'Rank'
      ]
    };

    try {
      const searchResults = await getSP().search(request);
      const durationMs = Math.round(performance.now() - startTime);
      const rawItems = searchResults.PrimarySearchResults || [];
      const results = this.normalizeResults(
        rawItems,
        queryText,
        options.variantKind || 'raw',
        options.variantLanguage || queryLanguage,
        lexicalTerms
      );

      logService.info(
        'search',
        `SharePoint Search: ${results.length} results`,
        JSON.stringify({
          rawResultCount: rawItems.length,
          keptResultCount: results.length,
          filteredOutCount: rawItems.length - results.length,
          topRawTitles: rawItems.slice(0, 10).map((item) => item.Title || item.Path || ''),
          topKeptTitles: results.slice(0, 10).map((item) => item.title)
        }, null, 2),
        durationMs
      );

      return {
        success: true,
        data: results,
        durationMs
      };
    } catch (error) {
      const durationMs = Math.round(performance.now() - startTime);
      const message = error instanceof Error ? error.message : 'SharePoint search failed';
      logService.error('search', `SharePoint Search failed: ${message}`, undefined, durationMs);
      return {
        success: false,
        error: { code: 'SharePointSearchError', message },
        durationMs
      };
    }
  }

  private normalizeResults(
    items: ISearchResult[],
    queryText: string,
    variantKind: SearchQueryVariantKind,
    variantLanguage: string,
    lexicalTerms: string[]
  ): ICopilotSearchResult[] {
    return items
      .filter((item) => typeof item.Path === 'string' && item.Path.trim().length > 0)
      .filter((item) => {
        if (lexicalTerms.length === 0) return true;
        const url = (item.OriginalPath || item.Path || '').trim();
        const urlMatchText = extractUrlMatchText(url);
        return lexicalTerms.some((term) =>
          containsLexicalTerm(item.Title, term)
          || containsLexicalTerm(item.HitHighlightedSummary, term)
          || containsLexicalTerm(item.Description, term)
          || containsLexicalTerm(urlMatchText, term)
        );
      })
      .map((item, index) => {
        const url = (item.OriginalPath || item.Path || '').trim();
        const nativeRank = typeof item.Rank === 'number' ? item.Rank : undefined;
        return {
          title: extractTitle(item, url),
          summary: cleanSummary(item.HitHighlightedSummary || item.Description),
          url,
          fileType: item.FileExtension || item.FileType,
          lastModified: normalizeDate(item.LastModifiedTime || item.Write),
          author: normalizeAuthor(item.Author),
          siteName: extractSiteName(item, url),
          score: nativeRank,
          source: 'sharepoint-search',
          sourceRank: index + 1,
          sourceNativeScore: nativeRank,
          queryText,
          variantKind,
          variantLanguage,
          language: normalizeLanguageTag(item.Culture)
        };
      });
  }
}
