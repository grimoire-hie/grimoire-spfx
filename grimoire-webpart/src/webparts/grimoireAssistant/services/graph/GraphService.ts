/**
 * GraphService
 * Core Microsoft Graph API wrapper using SPFx AadHttpClient.
 * All Graph calls go through AadHttpClient — no extra auth setup needed.
 */

import { AadHttpClient } from '@microsoft/sp-http';
import { logService } from '../logging/LogService';
import { normalizeError } from '../utils/errorUtils';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_BETA = 'https://graph.microsoft.com/beta';

export interface IGraphResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  durationMs?: number;
}

export class GraphService {
  constructor(protected readonly client: AadHttpClient) {}

  /**
   * GET request to Graph API (v1.0).
   */
  public async get<T>(path: string): Promise<IGraphResponse<T>> {
    return this.request<T>('GET', `${GRAPH_BASE}${path}`);
  }

  /**
   * GET request to Graph API (beta).
   */
  public async getBeta<T>(path: string): Promise<IGraphResponse<T>> {
    return this.request<T>('GET', `${GRAPH_BETA}${path}`);
  }

  /**
   * POST request to Graph API (v1.0).
   */
  public async post<T>(path: string, body: unknown): Promise<IGraphResponse<T>> {
    return this.request<T>('POST', `${GRAPH_BASE}${path}`, body);
  }

  /**
   * Core request method with logging and error handling.
   */
  private async request<T>(method: string, url: string, body?: unknown): Promise<IGraphResponse<T>> {
    const startTime = performance.now();
    const shortPath = url.replace(/https:\/\/graph\.microsoft\.com\/(v1\.0|beta)/, '');

    try {
      let response;
      if (method === 'POST' && body) {
        response = await this.client.post(
          url,
          AadHttpClient.configurations.v1,
          {
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
          }
        );
      } else {
        response = await this.client.get(url, AadHttpClient.configurations.v1);
      }

      const durationMs = Math.round(performance.now() - startTime);

      if (!response.ok) {
        const errorText = await response.text().catch(() => `HTTP ${response.status}`);
        logService.error('graph', `${method} ${shortPath}: ${response.status}`, errorText, durationMs);
        return { success: false, error: `HTTP ${response.status}: ${errorText}`, durationMs };
      }

      const data = await response.json() as T;
      logService.info('graph', `${method} ${shortPath}`, undefined, durationMs);
      return { success: true, data, durationMs };
    } catch (error) {
      const durationMs = Math.round(performance.now() - startTime);
      const normalizedError = normalizeError(error, `Graph ${method} request failed`);
      logService.error('graph', `${method} ${shortPath}: ${normalizedError.message}`, undefined, durationMs);
      return { success: false, error: normalizedError.message, durationMs };
    }
  }
}
