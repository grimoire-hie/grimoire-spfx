/**
 * PersistenceService — HTTP client for Grimoire backend persistence API.
 * Handles notes and preferences CRUD via grimoire-backend Azure Table Storage.
 */

import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import type { IProxyConfig } from '../../store/useGrimoireStore';
import { logService } from '../logging/LogService';
import { normalizeError } from '../utils/errorUtils';

const PERSISTENCE_TIMEOUT_MS = 10_000;

type SecuredRequestFactory = () => Promise<Response | HttpClientResponse>;

// ─── Types ──────────────────────────────────────────────────────

export interface INote {
  id: string;
  text: string;
  tags: string[];
  createdAt: string;
}

// ─── Service ────────────────────────────────────────────────────

export class PersistenceService {
  /**
   * Save a note to persistent storage.
   */
  static async saveNote(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    text: string,
    tags: string[]
  ): Promise<{ id: string }> {
    const result = await PersistenceService.post<{ id: string }>(proxyConfig, userApiClient, undefined, '/user/notes', {
      action: 'save',
      text,
      tags
    });
    logService.info('system', `Note saved: "${text.substring(0, 40)}..." (${tags.join(', ')})`);
    return result;
  }

  /**
   * List notes, optionally filtered by keyword.
   */
  static async listNotes(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    keyword?: string
  ): Promise<INote[]> {
    const result = await PersistenceService.post<{ notes: INote[] }>(proxyConfig, userApiClient, undefined, '/user/notes', {
      action: 'list',
      keyword
    });
    return result.notes || [];
  }

  /**
   * Delete a note by ID.
   */
  static async deleteNote(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    noteId: string
  ): Promise<void> {
    await PersistenceService.post(proxyConfig, userApiClient, undefined, '/user/notes', {
      action: 'delete',
      noteId
    });
    logService.info('system', `Note deleted: ${noteId}`);
  }

  /**
   * Get all preferences for a user.
   */
  static async getPreferences(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    getToken?: ((resource: string) => Promise<string>) | undefined
  ): Promise<Record<string, string>> {
    const result = await PersistenceService.post<{ preferences: Record<string, string> }>(
      proxyConfig, userApiClient, getToken, '/user/preferences', { action: 'get' }
    );
    return result.preferences || {};
  }

  /**
   * Set a single preference.
   */
  static async setPreference(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    getToken: ((resource: string) => Promise<string>) | undefined,
    key: string,
    value: string
  ): Promise<void> {
    await PersistenceService.post(proxyConfig, userApiClient, getToken, '/user/preferences', {
      action: 'set',
      key,
      value
    });
  }

  // ─── Private ──────────────────────────────────────────────────

  private static async post<T>(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    getToken: ((resource: string) => Promise<string>) | undefined,
    path: string,
    body: Record<string, unknown>
  ): Promise<T> {
    if (!proxyConfig.backendApiResource) {
      throw new Error('Backend API Resource is not configured for secured user routes.');
    }

    const url = `${proxyConfig.proxyUrl}${path}`;
    const requestFactory = await PersistenceService.createSecuredRequestFactory(
      proxyConfig,
      userApiClient,
      getToken,
      url,
      body
    );
    const response = await PersistenceService.fetchWithTimeout(
      requestFactory,
      PERSISTENCE_TIMEOUT_MS
    );

    if (!response.ok) {
      const errText = await response.text().catch(() => `HTTP ${response.status}`);
      throw new Error(`Persistence API error (${response.status}): ${errText.slice(0, 200)}`);
    }

    return response.json() as Promise<T>;
  }

  private static async createSecuredRequestFactory(
    proxyConfig: IProxyConfig,
    userApiClient: AadHttpClient | undefined,
    getToken: ((resource: string) => Promise<string>) | undefined,
    url: string,
    body: Record<string, unknown>
  ): Promise<SecuredRequestFactory> {
    if (userApiClient) {
      return () => userApiClient.post(
        url,
        AadHttpClient.configurations.v1,
        {
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(body)
        }
      );
    }

    if (!getToken) {
      throw new Error('Secured user API client and token provider are not available.');
    }

    const token = await getToken(proxyConfig.backendApiResource!);
    return () => fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`
      },
      body: JSON.stringify(body)
    });
  }

  private static async fetchWithTimeout(
    requestFactory: SecuredRequestFactory,
    timeoutMs: number
  ): Promise<Response | HttpClientResponse> {
    let timeoutId: ReturnType<typeof setTimeout> | undefined;
    const timeoutMarker = Symbol('persistence-timeout');
    const timeoutPromise = new Promise<typeof timeoutMarker>((resolve) => {
      timeoutId = setTimeout(() => resolve(timeoutMarker), timeoutMs);
    });

    try {
      const result = await Promise.race([requestFactory(), timeoutPromise]);
      if (result === timeoutMarker) {
        throw new Error('__PERSISTENCE_TIMEOUT__');
      }
      return result;
    } catch (error) {
      const normalizedError = normalizeError(error, 'Persistence API request failed');
      if (error instanceof Error && error.message === '__PERSISTENCE_TIMEOUT__') {
        throw new Error(`Persistence API timed out after ${Math.round(timeoutMs / 1000)}s`);
      }
      if (normalizedError.name === 'AbortError') {
        throw new Error(`Persistence API timed out after ${Math.round(timeoutMs / 1000)}s`);
      }
      throw error;
    } finally {
      if (timeoutId) {
        clearTimeout(timeoutId);
      }
    }
  }
}
