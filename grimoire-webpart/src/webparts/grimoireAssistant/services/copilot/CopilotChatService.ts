/**
 * CopilotChatService
 * Calls Microsoft 365 Copilot Chat API:
 *  - POST /beta/copilot/conversations
 *  - POST /beta/copilot/conversations/{id}/chat
 *
 * Uses ephemeral conversations per request.
 */

import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { logService } from '../logging/LogService';

const COPILOT_CREATE_CONVERSATION_ENDPOINT = 'https://graph.microsoft.com/beta/copilot/conversations';
const COPILOT_CONVERSATIONS_ENDPOINT = 'https://graph.microsoft.com/beta/copilot/conversations';

export interface ICopilotChatReference {
  title?: string;
  url?: string;
  attributionType?: string;
  attributionSource?: string;
}

export interface ICopilotChatRequest {
  prompt: string;
  fileUris?: string[];
  additionalContext?: string[];
  enableWebGrounding?: boolean;
  timeZone?: string;
}

export interface ICopilotChatResult {
  success: boolean;
  text?: string;
  references?: ICopilotChatReference[];
  conversationId?: string;
  error?: { code: string; message: string };
  durationMs: number;
}

interface ICopilotCreateConversationResponse {
  id?: string;
}

interface ICopilotConversationAttribution {
  providerDisplayName?: string;
  seeMoreWebUrl?: string;
  attributionType?: string;
  attributionSource?: string;
}

interface ICopilotConversationResponseMessage {
  text?: string;
  attributions?: ICopilotConversationAttribution[];
}

interface ICopilotChatResponse {
  messages?: ICopilotConversationResponseMessage[];
}

interface ICopilotConversationRequestBody {
  message: { text: string };
  locationHint: { timeZone: string };
  additionalContext?: Array<{ text: string }>;
  contextualResources?: {
    files?: Array<{ uri: string }>;
    webContext?: { isWebEnabled: boolean };
  };
}

export function isM365ContextUri(uri: string): boolean {
  try {
    const url = new URL(uri);
    const host = url.hostname.toLowerCase();
    return (
      host.endsWith('.sharepoint.com') ||
      host.endsWith('.sharepoint.us') ||
      host.endsWith('.sharepoint-mil.us') ||
      host.endsWith('.sharepoint.cn') ||
      host.endsWith('.onedrive.live.com')
    );
  } catch {
    return false;
  }
}

function getDefaultTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
}

export class CopilotChatService {
  private readonly aadClient: AadHttpClient;

  constructor(aadClient: AadHttpClient) {
    this.aadClient = aadClient;
  }

  public async chat(request: ICopilotChatRequest): Promise<ICopilotChatResult> {
    const start = performance.now();
    const prompt = request.prompt?.trim();
    if (!prompt) {
      return {
        success: false,
        error: { code: 'ValidationError', message: 'Prompt cannot be empty.' },
        durationMs: 0
      };
    }

    try {
      const createResp = await this.aadClient.post(
        COPILOT_CREATE_CONVERSATION_ENDPOINT,
        AadHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: '{}'
        }
      );

      if (!createResp.ok) {
        const message = await this.readErrorMessage(createResp);
        const durationMs = Math.round(performance.now() - start);
        logService.error('graph', `Create Copilot conversation failed: ${message}`, undefined, durationMs);
        return {
          success: false,
          error: { code: 'CreateConversationError', message },
          durationMs
        };
      }

      const created = (await createResp.json()) as ICopilotCreateConversationResponse;
      const conversationId = (created.id || '').trim();
      if (!conversationId) {
        const durationMs = Math.round(performance.now() - start);
        return {
          success: false,
          error: { code: 'CreateConversationError', message: 'Copilot conversation id missing in response.' },
          durationMs
        };
      }

      const body: ICopilotConversationRequestBody = {
        message: { text: prompt },
        locationHint: { timeZone: request.timeZone || getDefaultTimeZone() }
      };

      if (request.additionalContext && request.additionalContext.length > 0) {
        const cleaned = request.additionalContext
          .map((t) => t.trim())
          .filter((t) => t.length > 0)
          .map((text) => ({ text }));
        if (cleaned.length > 0) {
          body.additionalContext = cleaned;
        }
      }

      const hasFiles = !!request.fileUris?.length;
      const hasWebToggle = typeof request.enableWebGrounding === 'boolean';
      if (hasFiles || hasWebToggle) {
        body.contextualResources = {};
        if (hasFiles) {
          body.contextualResources.files = request.fileUris!.map((uri) => ({ uri }));
        }
        if (hasWebToggle) {
          body.contextualResources.webContext = { isWebEnabled: !!request.enableWebGrounding };
        }
      }

      const chatResp = await this.aadClient.post(
        `${COPILOT_CONVERSATIONS_ENDPOINT}/${encodeURIComponent(conversationId)}/chat`,
        AadHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body)
        }
      );

      const durationMs = Math.round(performance.now() - start);

      if (!chatResp.ok) {
        const message = await this.readErrorMessage(chatResp);
        logService.error('graph', `Copilot chat failed: ${message}`, undefined, durationMs);
        return {
          success: false,
          conversationId,
          error: { code: 'ChatError', message },
          durationMs
        };
      }

      const data = (await chatResp.json()) as ICopilotChatResponse;
      const responseMessage = this.pickBestResponseMessage(data.messages || [], prompt);
      const responseText = responseMessage?.text?.trim() || '';
      if (!responseText) {
        return {
          success: false,
          conversationId,
          error: { code: 'ChatError', message: 'Copilot chat response did not include text.' },
          durationMs
        };
      }

      const references = this.extractReferences(responseMessage?.attributions || []);

      logService.info(
        'graph',
        `Copilot chat completed (${references.length} refs)`,
        undefined,
        durationMs
      );

      return {
        success: true,
        text: responseText,
        references,
        conversationId,
        durationMs
      };
    } catch (error) {
      const durationMs = Math.round(performance.now() - start);
      const message = error instanceof Error ? error.message : 'Network error';
      logService.error('graph', `Copilot chat network error: ${message}`, undefined, durationMs);
      return {
        success: false,
        error: { code: 'NetworkError', message },
        durationMs
      };
    }
  }

  private pickBestResponseMessage(
    messages: ICopilotConversationResponseMessage[],
    prompt: string
  ): ICopilotConversationResponseMessage | undefined {
    const reversed = [...messages].reverse();
    const nonPrompt = reversed.find((m) => {
      const text = (m.text || '').trim();
      return !!text && text !== prompt;
    });
    if (nonPrompt) return nonPrompt;
    return reversed.find((m) => !!(m.text || '').trim());
  }

  private extractReferences(attributions: ICopilotConversationAttribution[]): ICopilotChatReference[] {
    const dedup = new Set<string>();
    const refs: ICopilotChatReference[] = [];
    attributions.forEach((attr) => {
      const title = (attr.providerDisplayName || '').trim() || undefined;
      const url = (attr.seeMoreWebUrl || '').trim() || undefined;
      const attributionType = (attr.attributionType || '').trim() || undefined;
      const attributionSource = (attr.attributionSource || '').trim() || undefined;
      if (!title && !url) return;
      const key = `${url || ''}|${title || ''}|${attributionType || ''}`;
      if (dedup.has(key)) return;
      dedup.add(key);
      refs.push({ title, url, attributionType, attributionSource });
    });
    return refs;
  }

  private async readErrorMessage(response: HttpClientResponse): Promise<string> {
    try {
      const data = await response.json() as {
        error?: { message?: string };
        message?: string;
      };
      return data?.error?.message || data?.message || `HTTP ${response.status}`;
    } catch {
      return `HTTP ${response.status}`;
    }
  }
}
