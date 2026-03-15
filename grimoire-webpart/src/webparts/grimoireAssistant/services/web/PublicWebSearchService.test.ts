jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

import type { IProxyConfig } from '../../store/useGrimoireStore';
import { PublicWebSearchService } from './PublicWebSearchService';

function createJsonResponse(ok: boolean, body: unknown, status: number = 200): Response {
  return {
    ok,
    status,
    json: jest.fn(async () => body),
    text: jest.fn(async () => JSON.stringify(body))
  } as unknown as Response;
}

describe('PublicWebSearchService', () => {
  const proxyConfig: IProxyConfig = {
    proxyUrl: 'https://example.com/api',
    proxyApiKey: 'test-key',
    backend: 'reasoning',
    deployment: 'grimoire-reasoning',
    apiVersion: '2025-01-01-preview'
  };

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('parses response text and url citations from the Responses API payload', async () => {
    const fetchMock = jest.fn().mockResolvedValue(createJsonResponse(true, {
      output: [
        {
          type: 'message',
          content: [
            {
              type: 'output_text',
              text: 'Microsoft is a multinational technology company.',
              annotations: [
                {
                  type: 'url_citation',
                  title: 'Microsoft - Wikipedia',
                  url: 'https://en.wikipedia.org/wiki/Microsoft'
                },
                {
                  type: 'url_citation',
                  url_citation: {
                    title: 'Microsoft',
                    url: 'https://www.microsoft.com/'
                  }
                }
              ]
            }
          ]
        }
      ]
    }));

    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      writable: true,
      value: fetchMock
    });

    const service = new PublicWebSearchService(proxyConfig);
    const result = await service.research('Summarize Microsoft');

    expect(result.success).toBe(true);
    expect(result.text).toContain('technology company');
    expect(result.references).toEqual([
      {
        title: 'Microsoft - Wikipedia',
        url: 'https://en.wikipedia.org/wiki/Microsoft'
      },
      {
        title: 'Microsoft',
        url: 'https://www.microsoft.com/'
      }
    ]);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0][0]).toBe('https://example.com/api/reasoning/openai/v1/responses?api-version=preview');
  });

  it('classifies unsupported tool responses from Azure', async () => {
    const fetchMock = jest.fn().mockResolvedValue(createJsonResponse(false, {
      error: {
        message: 'Invalid tool type: web_search_preview is not supported by this deployment.'
      }
    }, 400));

    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      writable: true,
      value: fetchMock
    });

    const service = new PublicWebSearchService(proxyConfig);
    const result = await service.research('Summarize Microsoft');

    expect(result.success).toBe(false);
    expect(result.capabilityStatus).toBe('unsupported');
    expect(result.error?.message).toContain('web_search_preview');
  });

  it('classifies blocked capability probe responses distinctly', async () => {
    const fetchMock = jest.fn().mockResolvedValue(createJsonResponse(false, {
      error: {
        message: 'Access to this tool is blocked by organization policy.'
      }
    }, 403));

    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      writable: true,
      value: fetchMock
    });

    const service = new PublicWebSearchService(proxyConfig);
    const result = await service.probeAvailability();

    expect(result).toEqual({
      status: 'blocked',
      detail: 'Access to this tool is blocked by organization policy.'
    });
  });
});
