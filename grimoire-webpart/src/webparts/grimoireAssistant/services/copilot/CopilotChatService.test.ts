jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('@microsoft/sp-http', () => ({
  AadHttpClient: {
    configurations: {
      v1: {}
    }
  }
}));

import { CopilotChatService, isM365ContextUri } from './CopilotChatService';

interface IFakeResponse {
  ok: boolean;
  status: number;
  json: jest.Mock<Promise<unknown>, []>;
}

function createResponse(ok: boolean, body: unknown, status: number = 200): IFakeResponse {
  return {
    ok,
    status,
    json: jest.fn(async () => body)
  };
}

describe('CopilotChatService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('creates an ephemeral conversation and sends chat with files + context + web toggle', async () => {
    const post = jest
      .fn()
      .mockResolvedValueOnce(createResponse(true, { id: 'conv-123' }))
      .mockResolvedValueOnce(createResponse(true, {
        messages: [
          { text: 'Summarize this file.' },
          {
            text: 'Final summary',
            attributions: [
              {
                providerDisplayName: 'SharePoint',
                seeMoreWebUrl: 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/a.pdf',
                attributionType: 'file',
                attributionSource: 'm365'
              }
            ]
          }
        ]
      }));

    const service = new CopilotChatService({ post } as never);
    const result = await service.chat({
      prompt: 'Summarize this file.',
      fileUris: ['https://contoso.sharepoint.com/sites/eng/Shared%20Documents/a.pdf'],
      additionalContext: ['Use concise language'],
      enableWebGrounding: false,
      timeZone: 'Europe/Zurich'
    });

    expect(result.success).toBe(true);
    expect(result.text).toBe('Final summary');
    expect(result.references?.[0]?.url).toBe('https://contoso.sharepoint.com/sites/eng/Shared%20Documents/a.pdf');
    expect(post).toHaveBeenCalledTimes(2);
    expect(post.mock.calls[0][0]).toBe('https://graph.microsoft.com/beta/copilot/conversations');
    expect(post.mock.calls[1][0]).toBe('https://graph.microsoft.com/beta/copilot/conversations/conv-123/chat');

    const body = JSON.parse(post.mock.calls[1][2].body as string) as {
      contextualResources?: { files?: Array<{ uri: string }>; webContext?: { isWebEnabled: boolean } };
      additionalContext?: Array<{ text: string }>;
      locationHint?: { timeZone: string };
    };
    expect(body.contextualResources?.files?.[0]?.uri).toBe('https://contoso.sharepoint.com/sites/eng/Shared%20Documents/a.pdf');
    expect(body.contextualResources?.webContext?.isWebEnabled).toBe(false);
    expect(body.additionalContext?.[0]?.text).toBe('Use concise language');
    expect(body.locationHint?.timeZone).toBe('Europe/Zurich');
  });

  it('surfaces create-conversation failures', async () => {
    const post = jest.fn().mockResolvedValueOnce(createResponse(false, { error: { message: 'Forbidden' } }, 403));
    const service = new CopilotChatService({ post } as never);

    const result = await service.chat({ prompt: 'hi' });
    expect(result.success).toBe(false);
    expect(result.error?.code).toBe('CreateConversationError');
    expect(result.error?.message).toContain('Forbidden');
    expect(post).toHaveBeenCalledTimes(1);
  });

  it('surfaces chat failures', async () => {
    const post = jest
      .fn()
      .mockResolvedValueOnce(createResponse(true, { id: 'conv-123' }))
      .mockResolvedValueOnce(createResponse(false, { error: { message: 'Bad Request' } }, 400));
    const service = new CopilotChatService({ post } as never);

    const result = await service.chat({ prompt: 'hi' });
    expect(result.success).toBe(false);
    expect(result.error?.code).toBe('ChatError');
    expect(result.error?.message).toContain('Bad Request');
    expect(result.conversationId).toBe('conv-123');
  });

  it('creates a fresh conversation for each chat request', async () => {
    const post = jest
      .fn()
      .mockResolvedValueOnce(createResponse(true, { id: 'conv-1' }))
      .mockResolvedValueOnce(createResponse(true, { messages: [{ text: 'A1' }] }))
      .mockResolvedValueOnce(createResponse(true, { id: 'conv-2' }))
      .mockResolvedValueOnce(createResponse(true, { messages: [{ text: 'A2' }] }));
    const service = new CopilotChatService({ post } as never);

    const first = await service.chat({ prompt: 'Q1' });
    const second = await service.chat({ prompt: 'Q2' });

    expect(first.success).toBe(true);
    expect(second.success).toBe(true);
    expect(first.conversationId).toBe('conv-1');
    expect(second.conversationId).toBe('conv-2');
    expect(post).toHaveBeenCalledTimes(4);
    expect(post.mock.calls[0][0]).toBe('https://graph.microsoft.com/beta/copilot/conversations');
    expect(post.mock.calls[2][0]).toBe('https://graph.microsoft.com/beta/copilot/conversations');
  });

  it('validates supported M365 context URIs', () => {
    expect(isM365ContextUri('https://contoso.sharepoint.com/sites/a/Shared%20Documents/a.pdf')).toBe(true);
    expect(isM365ContextUri('https://contoso-my.sharepoint.com/personal/u/Documents/b.docx')).toBe(true);
    expect(isM365ContextUri('https://example.com/file.pdf')).toBe(false);
    expect(isM365ContextUri('not-a-url')).toBe(false);
  });
});
