/**
 * CopilotRetrievalService
 * Calls Microsoft 365 Copilot Retrieval API: POST /beta/copilot/retrieval
 * Returns unordered text chunks with relevance scores (semantic search / RAG).
 * Maximum 25 results per request.
 *
 * Uses SPFx AadHttpClient for authentication.
 */

import { AadHttpClient } from '@microsoft/sp-http';
import type {
  ICopilotRetrievalRequest,
  ICopilotRetrievalHit,
  ICopilotRetrievalApiResponse,
  RetrievalDataSource,
  ICopilotSearchResult,
  ISearchApiResponse
} from '../../models/ISearchTypes';
import { logService } from '../logging/LogService';

const RETRIEVAL_ENDPOINT = 'https://graph.microsoft.com/beta/copilot/retrieval';

export class CopilotRetrievalService {
  private aadClient: AadHttpClient;

  constructor(aadClient: AadHttpClient) {
    this.aadClient = aadClient;
  }

  /**
   * Perform semantic retrieval search.
   *
   * @param query - Natural language search query
   * @param dataSource - Data source to search (default: 'sharePoint')
   * @param maxResults - Maximum results (1-25, default 10)
   * @param filterExpression - Optional KQL filter expression (e.g. 'filetype:pdf AND author:"John"')
   * @returns Normalized search results (same format as CopilotSearchService for RRF fusion)
   */
  public async search(
    query: string,
    dataSource: RetrievalDataSource = 'sharePoint',
    maxResults: number = 10,
    filterExpression?: string
  ): Promise<ISearchApiResponse<ICopilotSearchResult[]>> {
    const startTime = performance.now();

    const request: ICopilotRetrievalRequest = {
      queryString: query,
      dataSource,
      maximumNumberOfResults: Math.min(Math.max(1, maxResults), 25),
      resourceMetadata: ['title', 'createdBy', 'lastModifiedDateTime', 'fileExtension', 'siteName']
    };

    if (filterExpression) {
      request.filterExpression = filterExpression;
    }

    logService.info(
      'search',
      `Copilot Retrieval: "${query}" (source: ${dataSource}, max: ${maxResults}${filterExpression ? ', filter: ' + filterExpression : ''})`,
      JSON.stringify({
        endpoint: RETRIEVAL_ENDPOINT,
        request
      }, null, 2)
    );

    try {
      const response = await this.aadClient.post(
        RETRIEVAL_ENDPOINT,
        AadHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(request)
        }
      );

      const durationMs = Math.round(performance.now() - startTime);

      if (!response.ok) {
        const errorBody = await response.json();
        const errorMsg = errorBody?.error?.message || `HTTP ${response.status}`;
        logService.error('search', `Copilot Retrieval failed: ${errorMsg}`, undefined, durationMs);
        return {
          success: false,
          error: { code: 'RetrievalError', message: errorMsg },
          durationMs
        };
      }

      const apiResponse = (await response.json()) as ICopilotRetrievalApiResponse;
      const results = this.normalizeResults(apiResponse.retrievalHits || []);

      logService.info('search', `Copilot Retrieval: ${results.length} chunks`, undefined, durationMs);

      return {
        success: true,
        data: results,
        durationMs
      };
    } catch (error) {
      const durationMs = Math.round(performance.now() - startTime);
      const msg = error instanceof Error ? error.message : 'Network error';
      logService.error('search', `Copilot Retrieval network error: ${msg}`, undefined, durationMs);
      return {
        success: false,
        error: { code: 'NetworkError', message: msg },
        durationMs
      };
    }
  }

  /**
   * Normalize retrieval hits into the unified result format.
   * Each hit becomes a result; best extract text is used as summary.
   */
  private normalizeResults(hits: ICopilotRetrievalHit[]): ICopilotSearchResult[] {
    return hits.map((hit, index) => {
      const metadata = hit.resourceMetadata || {};
      const bestExtract = this.getBestExtract(hit);

      return {
        title: metadata.title || this.extractTitleFromUrl(hit.webUrl),
        summary: bestExtract.text,
        url: hit.webUrl,
        fileType: metadata.fileExtension || this.extractFileType(hit.webUrl),
        lastModified: metadata.lastModifiedDateTime,
        author: metadata.createdBy,
        siteName: metadata.siteName,
        score: bestExtract.score,
        sourceRank: index + 1,
        sourceNativeScore: bestExtract.score,
        source: 'copilot-retrieval'
      };
    });
  }

  private getBestExtract(hit: ICopilotRetrievalHit): { text: string; score?: number } {
    if (!hit.extracts || hit.extracts.length === 0) {
      return { text: '' };
    }
    // Pick the extract with the highest relevance score
    let best = hit.extracts[0];
    for (let i = 1; i < hit.extracts.length; i++) {
      if ((hit.extracts[i].relevanceScore || 0) > (best.relevanceScore || 0)) {
        best = hit.extracts[i];
      }
    }
    return { text: best.text, score: best.relevanceScore };
  }

  private extractTitleFromUrl(url: string): string {
    try {
      const parts = url.split('/');
      const last = parts[parts.length - 1];
      return decodeURIComponent(last).replace(/\.[^.]+$/, '');
    } catch {
      return url;
    }
  }

  private extractFileType(url: string): string | undefined {
    const match = url.match(/\.(\w{2,5})(?:\?|$)/);
    return match ? match[1] : undefined;
  }
}
