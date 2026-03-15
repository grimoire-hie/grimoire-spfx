jest.mock('../../store/useGrimoireStore', () => ({
  useGrimoireStore: {
    getState: jest.fn()
  }
}));

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    getCurrentArtifacts: jest.fn(() => ({})),
    onBlockCreated: jest.fn(),
    onBlockUpdated: jest.fn(),
    onToolComplete: jest.fn()
  }
}));

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

import { CopilotChatService } from '../copilot/CopilotChatService';
import { ASSISTANT_SUMMARY_TARGET_TEXT } from '../../config/assistantLengthLimits';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { buildContentRuntimeHandlers } from './ToolRegistryRuntimeHandlersContent';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeDispatchOutcome, IToolRuntimeHandlerDeps, ToolRuntimeHandler } from './ToolRuntimeHandlerTypes';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type { ContentRuntimeToolName } from './ToolRuntimeHandlerPartitions';

interface IStoreStateForTests {
  blocks: unknown[];
  pushBlock?: jest.Mock;
  updateBlock?: jest.Mock;
}

function createStore(overrides: Partial<IFunctionCallStore> = {}): IFunctionCallStore {
  return {
    aadHttpClient: { post: jest.fn() } as never,
    proxyConfig: undefined,
    getToken: undefined,
    mcpEnvironmentId: undefined,
    userContext: undefined,
    copilotWebGroundingEnabled: false,
    mcpConnections: [],
    pushBlock: jest.fn(),
    updateBlock: jest.fn(),
    removeBlock: jest.fn(),
    clearBlocks: jest.fn(),
    setExpression: jest.fn(),
    setActivityStatus: jest.fn(),
    ...overrides
  };
}

function createDeps(
  store: IFunctionCallStore,
  overrides: Partial<IToolRuntimeHandlerDeps> = {}
): IToolRuntimeHandlerDeps {
  return {
    store,
    awaitAsync: true,
    aadClient: store.aadHttpClient,
    sitesService: {} as never,
    peopleService: undefined,
    ...overrides
  };
}

function extractOutput(result: IToolRuntimeDispatchOutcome): string {
  return result.output;
}

describe('ToolRegistryRuntimeHandlersContent (Copilot gateway)', () => {
  const getStateMock = useGrimoireStore.getState as jest.Mock;

  let state: IStoreStateForTests;

  beforeEach(() => {
    jest.restoreAllMocks();
    jest.clearAllMocks();
    state = {
      blocks: [],
      pushBlock: jest.fn(),
      updateBlock: jest.fn()
    };
    getStateMock.mockImplementation(() => state);
  });

  function createHandlers(): Pick<Record<RuntimeHandledToolName, ToolRuntimeHandler>, ContentRuntimeToolName> {
    return buildContentRuntimeHandlers({});
  }

  it('routes file summarize through Copilot Chat with file URI context', async () => {
    jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'Voici un **résumé** avec du `markdown`.',
      references: [{ title: 'SharePoint', url: 'https://contoso.sharepoint.com/sites/eng/a.pdf', attributionType: 'file' }],
      conversationId: 'conv-1',
      durationMs: 42
    });

    const handlers = createHandlers();
    const store = createStore({ copilotWebGroundingEnabled: true });
    const fileUrl = 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/Quarterly%20Report.pdf';

    const output = await handlers.read_file_content({
      file_url: fileUrl,
      mode: 'summarize'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect(parsed.mode).toBe('summarize');
    expect(parsed.content).toBe('Voici un résumé avec du markdown.');
    expect(parsed.content).not.toContain('**');
    expect(parsed.content).not.toContain('`');
    expect(parsed.references[0].url).toBe('https://contoso.sharepoint.com/sites/eng/a.pdf');
    expect(parsed.note).toContain('already generated the final summary text');
    expect((CopilotChatService.prototype.chat as jest.Mock).mock.calls[0][0].prompt).toContain(ASSISTANT_SUMMARY_TARGET_TEXT);
    expect(CopilotChatService.prototype.chat).toHaveBeenCalledWith(expect.objectContaining({
      fileUris: [fileUrl],
      enableWebGrounding: true
    }));
  });

  it('routes multi-file answer through Copilot Chat with file_urls context', async () => {
    jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'Differences found.',
      references: [],
      conversationId: 'conv-multi',
      durationMs: 27
    });

    const handlers = createHandlers();
    const store = createStore();
    const fileA = 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/A.docx';
    const fileB = 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/B.docx';

    const output = await handlers.read_file_content({
      file_urls: [fileA, fileB],
      mode: 'answer',
      question: 'What changed between these documents?'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect(parsed.mode).toBe('answer');
    expect(parsed.fileCount).toBe(2);
    expect(parsed.fileUrls).toEqual([fileA, fileB]);
    expect(CopilotChatService.prototype.chat).toHaveBeenCalledWith(expect.objectContaining({
      fileUris: [fileA, fileB]
    }));
  });

  it('uses the shared target for multi-file summarize prompts', async () => {
    jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'Summary across both files.',
      references: [],
      conversationId: 'conv-multi-summarize',
      durationMs: 29
    });

    const handlers = createHandlers();
    const store = createStore();
    const fileA = 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/A.docx';
    const fileB = 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/B.docx';

    const output = await handlers.read_file_content({
      file_urls: [fileA, fileB],
      mode: 'summarize'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect((CopilotChatService.prototype.chat as jest.Mock).mock.calls[0][0].prompt).toContain(ASSISTANT_SUMMARY_TARGET_TEXT);
  });

  it('rejects multi-file full mode deterministically', async () => {
    const chatSpy = jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'unused',
      durationMs: 1
    });
    const handlers = createHandlers();
    const store = createStore();

    const output = await handlers.read_file_content({
      file_urls: [
        'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/A.docx',
        'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/B.docx'
      ],
      mode: 'full'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain('single file only');
    expect(chatSpy).not.toHaveBeenCalled();
  });

  it('rejects non-M365 file URLs deterministically', async () => {
    const chatSpy = jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'unused',
      durationMs: 1
    });
    const handlers = createHandlers();
    const store = createStore();

    const output = await handlers.read_file_content({
      file_url: 'https://example.com/file.pdf',
      mode: 'summarize'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain('Only SharePoint/OneDrive URLs are supported');
    expect(chatSpy).not.toHaveBeenCalled();
  });

  it('routes email full retrieval via Copilot gateway', async () => {
    jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'Full email content',
      references: [],
      conversationId: 'conv-email',
      durationMs: 30
    });

    const handlers = createHandlers();
    const store = createStore();

    const output = await handlers.read_email_content({
      subject: 'Quarterly update',
      sender: 'Adele Vance',
      date_hint: '2026-03-01',
      mode: 'full'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect(parsed.mode).toBe('full');
    expect(parsed.content).toBe('Full email content');
    expect(CopilotChatService.prototype.chat).toHaveBeenCalledWith(expect.objectContaining({
      enableWebGrounding: false
    }));
  });

  it('routes teams answer mode via Copilot gateway and requires question', async () => {
    jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: true,
      text: 'Owner is Sam.',
      references: [],
      conversationId: 'conv-teams',
      durationMs: 25
    });

    const handlers = createHandlers();
    const store = createStore();

    const output = await handlers.read_teams_messages({
      chat_or_channel: 'General',
      mode: 'answer',
      question: 'Who owns the migration?'
    }, createDeps(store));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect(parsed.mode).toBe('answer');
    expect(parsed.content).toContain('Owner is Sam');

    const missingQuestion = await handlers.read_teams_messages({
      chat_or_channel: 'General',
      mode: 'answer'
    }, createDeps(store));
    const missingParsed = JSON.parse(extractOutput(missingQuestion));
    expect(missingParsed.success).toBe(false);
    expect(missingParsed.error).toContain('requires the "question" parameter');
  });

  it('does not execute legacy fallback when Copilot file processing fails', async () => {
    jest.spyOn(CopilotChatService.prototype, 'chat').mockResolvedValue({
      success: false,
      error: { code: 'ChatError', message: 'Copilot unavailable' },
      durationMs: 12
    });

    const legacyReadSmallTextFile = jest.fn();
    const handlers = createHandlers();
    const store = createStore();

    const output = await handlers.read_file_content({
      file_url: 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/a.pdf',
      mode: 'summarize'
    }, createDeps(store, { sitesService: { readSmallTextFile: legacyReadSmallTextFile } as never }));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(false);
    expect(parsed.error).toContain('Copilot unavailable');
    expect(CopilotChatService.prototype.chat).toHaveBeenCalledTimes(1);
    expect(legacyReadSmallTextFile).not.toHaveBeenCalled();
  });

  it('preserves document library and item ids when browsing a library', async () => {
    const handlers = createHandlers();
    const store = createStore();
    const browseDrive = jest.fn().mockResolvedValue({
      success: true,
      data: [
        {
          name: 'SPFx.pdf',
          type: 'file',
          url: 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/SPFx.pdf',
          documentLibraryId: 'drive-eng-docs',
          fileOrFolderId: 'drive-item-spfx',
          fileType: 'pdf'
        }
      ]
    });

    const output = await handlers.browse_document_library({
      site_url: 'https://contoso.sharepoint.com/sites/eng'
    }, createDeps(store, { sitesService: { browseDrive } as never }));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect(state.pushBlock).toHaveBeenCalledTimes(1);
    const block = state.pushBlock?.mock.calls[0][0];
    expect(block.type).toBe('document-library');
    expect(block.data.items[0]).toEqual(expect.objectContaining({
      name: 'SPFx.pdf',
      documentLibraryId: 'drive-eng-docs',
      fileOrFolderId: 'drive-item-spfx'
    }));
  });

  it('preserves hidden site/list/item identifiers when showing list items', async () => {
    const handlers = createHandlers();
    const store = createStore();
    const getListItems = jest.fn().mockResolvedValue({
      success: true,
      data: {
        columns: ['Title'],
        items: [
          {
            Title: 'Release checklist',
            itemId: '42',
            siteId: 'contoso.sharepoint.com,site-guid,web-guid',
            listId: 'list-guid-checklists',
            listName: 'Checklists',
            listUrl: 'https://contoso.sharepoint.com/sites/eng/Lists/Checklists'
          }
        ],
        totalCount: 1
      }
    });

    const output = await handlers.show_list_items({
      site_url: 'https://contoso.sharepoint.com/sites/eng',
      list_name: 'Checklists'
    }, createDeps(store, { sitesService: { getListItems } as never }));
    const parsed = JSON.parse(extractOutput(output));

    expect(parsed.success).toBe(true);
    expect(state.pushBlock).toHaveBeenCalledTimes(1);
    const block = state.pushBlock?.mock.calls[0][0];
    expect(block.type).toBe('list-items');
    expect(block.data.items[0]).toEqual(expect.objectContaining({
      Title: 'Release checklist',
      itemId: '42',
      siteId: 'contoso.sharepoint.com,site-guid,web-guid',
      listId: 'list-guid-checklists',
      listName: 'Checklists'
    }));
  });
});
