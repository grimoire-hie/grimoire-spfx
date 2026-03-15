/**
 * CopilotSearchService
 * Calls Microsoft 365 Copilot Search API: POST /beta/copilot/search
 * Returns ordered documents ranked by relevance (hybrid semantic + lexical search).
 * Maximum 100 results per request.
 *
 * Uses SPFx AadHttpClient for authentication — no extra Entra app registration needed.
 */

import { AadHttpClient } from '@microsoft/sp-http';
import type {
  ICopilotSearchRequest,
  ICopilotSearchHit,
  ICopilotSearchApiResponse,
  ICopilotSearchResult,
  ISearchApiResponse
} from '../../models/ISearchTypes';
import { logService } from '../logging/LogService';

const SEARCH_ENDPOINT = 'https://graph.microsoft.com/beta/copilot/search';

export class CopilotSearchService {
  private aadClient: AadHttpClient;

  constructor(aadClient: AadHttpClient) {
    this.aadClient = aadClient;
  }

  /**
   * Search for documents using the Copilot Search API.
   *
   * @param query - Natural language search query
   * @param pageSize - Number of results (1-100, default 25)
   * @param pathFilter - Optional path filter (only `path` is documented as working in Copilot Search)
   * @returns Normalized search results
   */
  public async search(
    query: string,
    pageSize: number = 25,
    pathFilter?: string
  ): Promise<ISearchApiResponse<ICopilotSearchResult[]>> {
    const startTime = performance.now();

    const metadataNames = ['title', 'createdBy', 'lastModifiedDateTime', 'fileExtension', 'siteName'];

    const request: ICopilotSearchRequest = {
      query,
      pageSize: Math.min(Math.max(1, pageSize), 100),
      dataSources: {
        oneDrive: {
          resourceMetadataNames: metadataNames
        }
      }
    };

    // Only `path` filter is documented as working in Copilot Search API
    if (pathFilter) {
      request.dataSources = {
        oneDrive: {
          filterExpression: pathFilter,
          resourceMetadataNames: metadataNames
        }
      };
    }

    logService.info(
      'search',
      `Copilot Search: "${query}" (pageSize: ${pageSize}${pathFilter ? ', path: ' + pathFilter : ''})`,
      JSON.stringify({
        endpoint: SEARCH_ENDPOINT,
        request
      }, null, 2)
    );

    try {
      const response = await this.aadClient.post(
        SEARCH_ENDPOINT,
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
        logService.error('search', `Copilot Search failed: ${errorMsg}`, undefined, durationMs);
        return {
          success: false,
          error: { code: 'SearchError', message: errorMsg },
          durationMs
        };
      }

      const apiResponse = (await response.json()) as ICopilotSearchApiResponse;
      const results = this.normalizeResults(apiResponse.searchHits || []);

      logService.info('search', `Copilot Search: ${results.length} of ${apiResponse.totalCount} results`, undefined, durationMs);

      return {
        success: true,
        data: results,
        durationMs
      };
    } catch (error) {
      const durationMs = Math.round(performance.now() - startTime);
      const msg = error instanceof Error ? error.message : 'Network error';
      logService.error('search', `Copilot Search network error: ${msg}`, undefined, durationMs);
      return {
        success: false,
        error: { code: 'NetworkError', message: msg },
        durationMs
      };
    }
  }

  /**
   * Normalize raw API hits into our unified result format.
   */
  private normalizeResults(hits: ICopilotSearchHit[]): ICopilotSearchResult[] {
    return hits.map((hit, index) => {
      const metadata = hit.resourceMetadata || {};
      return {
        title: metadata.title || this.extractTitleFromUrl(hit.webUrl),
        summary: hit.preview || '',
        url: hit.webUrl,
        fileType: metadata.fileExtension || this.extractFileType(hit.webUrl),
        lastModified: metadata.lastModifiedDateTime,
        author: metadata.createdBy,
        siteName: metadata.siteName,
        source: 'copilot-search',
        sourceRank: index + 1
      };
    });
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
