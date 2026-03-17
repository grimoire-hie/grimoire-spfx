import type { IProxyConfig, PublicWebSearchCapabilityStatus } from '../../store/useGrimoireStore';
import { logService } from '../logging/LogService';

const RESPONSES_API_VERSION = 'preview';
const PUBLIC_WEB_SEARCH_TIMEOUT_MS = 25_000;
const PUBLIC_WEB_PROBE_TIMEOUT_MS = 12_000;

export interface IPublicWebReference {
  title?: string;
  url: string;
}

export interface IPublicWebSearchProbeResult {
  status: PublicWebSearchCapabilityStatus;
  detail?: string;
}

export interface IPublicWebSearchResult {
  success: boolean;
  text?: string;
  references: IPublicWebReference[];
  capabilityStatus: PublicWebSearchCapabilityStatus;
  error?: { code: string; message: string };
  durationMs: number;
}

interface IResponsesApiErrorPayload {
  error?: {
    code?: string;
    message?: string;
  };
  message?: string;
}

interface IResponsesUrlCitation {
  url?: string;
  title?: string;
}

interface IResponsesAnnotation {
  type?: string;
  title?: string;
  url?: string;
  url_citation?: IResponsesUrlCitation;
}

interface IResponsesContentPart {
  type?: string;
  text?: string;
  annotations?: IResponsesAnnotation[];
}

interface IResponsesOutputItem {
  type?: string;
  content?: IResponsesContentPart[];
  annotations?: IResponsesAnnotation[];
}

interface IResponsesApiPayload {
  output_text?: string;
  output?: IResponsesOutputItem[];
}

function dedupeReferences(references: IPublicWebReference[]): IPublicWebReference[] {
  const seen = new Set<string>();
  const deduped: IPublicWebReference[] = [];

  references.forEach((reference) => {
    const url = (reference.url || '').trim();
    if (!url) return;
    const title = (reference.title || '').trim() || undefined;
    const key = `${url.toLowerCase()}|${(title || '').toLowerCase()}`;
    if (seen.has(key)) return;
    seen.add(key);
    deduped.push({ title, url });
  });

  return deduped;
}

function collectCitation(annotation: IResponsesAnnotation | undefined): IPublicWebReference | undefined {
  if (!annotation || annotation.type !== 'url_citation') {
    return undefined;
  }

  const url = (annotation.url_citation?.url || annotation.url || '').trim();
  if (!url) {
    return undefined;
  }

  const title = (annotation.url_citation?.title || annotation.title || '').trim() || undefined;
  return { title, url };
}

function parseResponsePayload(payload: IResponsesApiPayload): { text: string; references: IPublicWebReference[] } {
  const collectedReferences: IPublicWebReference[] = [];

  (payload.output || []).forEach((item) => {
    (item.annotations || []).forEach((annotation) => {
      const reference = collectCitation(annotation);
      if (reference) collectedReferences.push(reference);
    });

    (item.content || []).forEach((part) => {
      (part.annotations || []).forEach((annotation) => {
        const reference = collectCitation(annotation);
        if (reference) collectedReferences.push(reference);
      });
    });
  });

  const topLevelText = typeof payload.output_text === 'string'
    ? payload.output_text.trim()
    : '';
  if (topLevelText) {
    return {
      text: topLevelText,
      references: dedupeReferences(collectedReferences)
    };
  }

  const contentParts: string[] = [];
  (payload.output || []).forEach((item) => {
    (item.content || []).forEach((part) => {
      const text = typeof part.text === 'string' ? part.text.trim() : '';
      if (!text) return;
      contentParts.push(text);
    });
  });

  const seenText = new Set<string>();
  const dedupedText = contentParts.filter((part) => {
    if (seenText.has(part)) return false;
    seenText.add(part);
    return true;
  });

  return {
    text: dedupedText.join('\n\n').trim(),
    references: dedupeReferences(collectedReferences)
  };
}

function classifyCapabilityFailure(status: number | undefined, message: string): PublicWebSearchCapabilityStatus {
  const normalized = message.toLowerCase();

  if (
    status === 403
    || /forbidden|blocked|not allowed|disabled|denied|policy|access denied|organization/i.test(normalized)
  ) {
    return 'blocked';
  }

  if (
    status === 404
    || /unsupported|not supported|does not support|invalid tool|unknown tool|tool type|web_search_preview|responses api/i.test(normalized)
  ) {
    return 'unsupported';
  }

  return 'error';
}

function buildResearchInput(query: string, targetUrl?: string): string {
  if (!targetUrl) {
    return query;
  }

  return [
    'Check the public webpage at the target URL and summarize the page contents for the user.',
    `Target URL: ${targetUrl}`,
    `User request: ${query}`
  ].join('\n');
}

export class PublicWebSearchService {
  private readonly proxyConfig: IProxyConfig;

  constructor(proxyConfig: IProxyConfig) {
    this.proxyConfig = proxyConfig;
  }

  public async probeAvailability(): Promise<IPublicWebSearchProbeResult> {
    const result = await this.execute({
      input: 'Reply with the single word "available".',
      max_output_tokens: 16
    }, PUBLIC_WEB_PROBE_TIMEOUT_MS);

    if (!result.ok) {
      return {
        status: result.capabilityStatus,
        detail: result.message
      };
    }

    return { status: 'available' };
  }

  public async research(query: string, targetUrl?: string): Promise<IPublicWebSearchResult> {
    const start = performance.now();
    const normalizedQuery = query.trim();

    if (!normalizedQuery) {
      return {
        success: false,
        references: [],
        capabilityStatus: 'error',
        error: {
          code: 'ValidationError',
          message: 'Public web search query cannot be empty.'
        },
        durationMs: 0
      };
    }

    const result = await this.execute({
      input: buildResearchInput(normalizedQuery, targetUrl),
      max_output_tokens: 900
    }, PUBLIC_WEB_SEARCH_TIMEOUT_MS);
    const durationMs = Math.round(performance.now() - start);

    if (!result.ok) {
      logService.warning(
        'llm',
        `Public web search failed (${result.capabilityStatus}): ${result.message}`
      );
      return {
        success: false,
        references: [],
        capabilityStatus: result.capabilityStatus,
        error: {
          code: result.capabilityStatus === 'unsupported' ? 'UnsupportedError' : 'PublicWebSearchError',
          message: result.message
        },
        durationMs
      };
    }

    const parsed = parseResponsePayload(result.payload);
    if (!parsed.text && parsed.references.length === 0) {
      return {
        success: false,
        references: [],
        capabilityStatus: 'available',
        error: {
          code: 'EmptyResponseError',
          message: 'Public web search returned an empty response.'
        },
        durationMs
      };
    }

    logService.info(
      'llm',
      `Public web search completed (${parsed.references.length} refs)`,
      undefined,
      durationMs
    );

    return {
      success: true,
      text: parsed.text,
      references: parsed.references,
      capabilityStatus: 'available',
      durationMs
    };
  }

  private async execute(
    body: { input: string; max_output_tokens: number },
    timeoutMs: number
  ): Promise<
    | { ok: true; payload: IResponsesApiPayload }
    | { ok: false; message: string; capabilityStatus: PublicWebSearchCapabilityStatus }
  > {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const url = `${this.proxyConfig.proxyUrl}/${this.proxyConfig.backend}/openai/v1/responses?api-version=${RESPONSES_API_VERSION}`;

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-functions-key': this.proxyConfig.proxyApiKey
        },
        body: JSON.stringify({
          model: this.proxyConfig.deployment,
          input: body.input,
          max_output_tokens: body.max_output_tokens,
          tools: [{ type: 'web_search_preview' }]
        }),
        signal: controller.signal
      });

      if (!response.ok) {
        const { message } = await this.readErrorMessage(response);
        return {
          ok: false,
          message,
          capabilityStatus: classifyCapabilityFailure(response.status, message)
        };
      }

      const payload = await response.json() as IResponsesApiPayload;
      return { ok: true, payload };
    } catch (error) {
      const message = (error as Error).name === 'AbortError'
        ? `Public web search timed out after ${Math.round(timeoutMs / 1000)}s.`
        : (error as Error).message || 'Network error while calling public web search.';
      return {
        ok: false,
        message,
        capabilityStatus: 'error'
      };
    } finally {
      clearTimeout(timer);
    }
  }

  private async readErrorMessage(response: Response): Promise<{ message: string }> {
    try {
      const payload = await response.json() as IResponsesApiErrorPayload;
      return {
        message: payload.error?.message || payload.message || `HTTP ${response.status}`
      };
    } catch {
      const text = await response.text().catch(() => '');
      return {
        message: text.trim() || `HTTP ${response.status}`
      };
    }
  }
}
