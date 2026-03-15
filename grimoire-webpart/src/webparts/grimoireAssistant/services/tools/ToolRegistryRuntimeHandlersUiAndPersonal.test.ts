jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    getCurrentTaskContext: jest.fn(),
    captureCurrentSourceContext: jest.fn(),
    getCurrentArtifacts: jest.fn(() => ({})),
    getBlockTracker: jest.fn(() => ({ get: jest.fn(() => undefined) })),
    getRecentEvents: jest.fn(() => []),
    onBlockCreated: jest.fn(),
    onBlockUpdated: jest.fn(),
    onToolComplete: jest.fn(),
    onBlockRemoved: jest.fn()
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

jest.mock('../graph/GraphService', () => ({
  GraphService: jest.fn()
}));

jest.mock('../context/PersistenceService', () => ({
  PersistenceService: jest.fn()
}));

jest.mock('../mcp/McpExecutionAdapter', () => {
  const actual = jest.requireActual('../mcp/McpExecutionAdapter');
  return {
    ...actual,
    executeCatalogMcpTool: jest.fn(actual.executeCatalogMcpTool)
  };
});

jest.mock('../mcp/McpResultMapper', () => {
  const actual = jest.requireActual('../mcp/McpResultMapper');
  return {
    ...actual,
    mapMcpResultToBlock: jest.fn(actual.mapMcpResultToBlock)
  };
});

import type { IErrorData, IInfoCardData, IMarkdownData } from '../../models/IBlock';
import type { IFormData, ISearchResultsData, IUserCardData } from '../../models/IBlock';
import { createBlock } from '../../models/IBlock';
import type { IMcpConnection } from '../../models/IMcpTypes';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { GraphService } from '../graph/GraphService';
import * as shareSubmissionService from '../sharing/ShareSubmissionService';
import * as pnpContext from '../pnp/pnpContext';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { logService } from '../logging/LogService';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';
import { buildUiAndPersonalRuntimeHandlers } from './ToolRegistryRuntimeHandlersUiAndPersonal';
import { shouldSuppressSelectionList } from './ToolSelectionListGuard';

function createStore(overrides: Partial<IFunctionCallStore> = {}): IFunctionCallStore {
  return {
    aadHttpClient: undefined,
    proxyConfig: undefined,
    getToken: undefined,
    mcpEnvironmentId: 'env-123',
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

function createDeps(store: IFunctionCallStore): IToolRuntimeHandlerDeps {
  return {
    store,
    awaitAsync: false,
    aadClient: undefined,
    sitesService: undefined,
    peopleService: undefined
  };
}

describe('ToolRegistryRuntimeHandlersUiAndPersonal selection-list guardrail', () => {
  it('suppresses generic chooser list when actionable block is already visible', () => {
    expect(
      shouldSuppressSelectionList(
        'I found these docs. Pick one to open, preview, or summarize.',
        'search-results',
        'animals information'
      )
    ).toBe(true);
  });

  it('allows list when user explicitly asks for options UI', () => {
    expect(
      shouldSuppressSelectionList(
        'Pick one',
        'search-results',
        'show options as radio buttons'
      )
    ).toBe(false);
  });

  it('does not suppress for non-actionable latest blocks', () => {
    expect(
      shouldSuppressSelectionList(
        'Pick one',
        'selection-list',
        'animals information'
      )
    ).toBe(false);
  });
});

describe('ToolRegistryRuntimeHandlersUiAndPersonal.set_expression', () => {
  it('suppresses avatar expression changes when the avatar is disabled', async () => {
    const store = createStore({ avatarEnabled: false });
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.set_expression({ expression: 'happy' }, deps);

    expect(result.phase).toBe('complete');
    expect(store.setExpression).not.toHaveBeenCalled();
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({
      success: true,
      suppressed: true
    }));
  });
});

describe('ToolRegistryRuntimeHandlersUiAndPersonal recent document insights filtering', () => {
  beforeEach(() => {
    (GraphService as unknown as jest.Mock).mockReset();
  });

  it('filters SharePoint pages and sites out of recent documents', async () => {
    const graphGet = jest.fn().mockResolvedValue({
      success: true,
      data: {
        value: [
          {
            resourceVisualization: { title: 'Animal Stories and Facts.docx', type: 'docx' },
            resourceReference: { webUrl: 'https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx' },
            lastUsed: { lastAccessedDateTime: '2026-03-10T08:00:00.000Z' }
          },
          {
            resourceVisualization: { title: 'Power Platform.pdf', type: 'pdf' },
            resourceReference: { webUrl: 'https://tenant.sharepoint.com/sites/dev/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Power%20Platform.pdf&action=default' },
            lastUsed: { lastAccessedDateTime: '2026-03-10T07:00:00.000Z' }
          },
          {
            resourceVisualization: { title: 'Grimoire.aspx', type: 'Web' },
            resourceReference: { webUrl: 'https://tenant.sharepoint.com/sites/dev/SitePages/Grimoire.aspx' },
            lastUsed: { lastAccessedDateTime: '2026-03-10T06:00:00.000Z' }
          },
          {
            resourceVisualization: { title: 'copilot-test', type: 'spsite' },
            resourceReference: { webUrl: 'https://tenant.sharepoint.com/sites/copilot-test' },
            lastUsed: { lastAccessedDateTime: '2026-03-10T05:00:00.000Z' }
          },
          {
            resourceVisualization: { title: 'Nici.aspx', type: 'Web Page' },
            resourceReference: { webUrl: 'https://tenant.sharepoint.com/sites/dev/SitePages/Nici.aspx', type: 'sitePage' },
            lastUsed: { lastAccessedDateTime: '2026-03-10T04:00:00.000Z' }
          },
          {
            resourceVisualization: { title: 'copilot-test-renewable', type: '' },
            resourceReference: { webUrl: 'https://tenant.sharepoint.com/sites/copilot-test-renewable', type: 'spsite' },
            lastUsed: { lastAccessedDateTime: '2026-03-10T03:00:00.000Z' }
          }
        ]
      }
    });
    (GraphService as unknown as jest.Mock).mockImplementation(() => ({
      get: graphGet
    }));

    const store = createStore();
    const deps: IToolRuntimeHandlerDeps = {
      ...createDeps(store),
      aadClient: {} as never,
      awaitAsync: true
    };
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.get_recent_documents({ max_results: '5' }, deps);

    expect(result.phase).toBe('complete');
    expect(graphGet).toHaveBeenCalledWith('/me/insights/used?$top=5&$orderby=lastUsed/lastAccessedDateTime desc');
    expect(store.pushBlock).toHaveBeenCalledTimes(1);
    expect(store.updateBlock).toHaveBeenCalledTimes(1);

    const updatePayload = (store.updateBlock as jest.Mock).mock.calls[0][1] as { data: ISearchResultsData };
    expect(updatePayload.data.totalCount).toBe(2);
    expect(updatePayload.data.results.map((item) => item.title)).toEqual([
      'Animal Stories and Facts.docx',
      'Power Platform.pdf'
    ]);
    expect(updatePayload.data.results.map((item) => item.fileType)).toEqual([
      'docx',
      'pdf'
    ]);

    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({
      success: true,
      count: 2
    }));
  });
});

describe('ToolRegistryRuntimeHandlersUiAndPersonal.list_m365_servers', () => {
  it('renders a plain-language capability overview', async () => {
    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.list_m365_servers({}, deps);

    expect(result.phase).toBe('complete');
    expect(store.pushBlock).toHaveBeenCalledTimes(1);

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IMarkdownData;

    expect(block.title).toBe('Grimoire Capabilities');
    expect(data.content).toContain('## What Grimoire Can Help With');
    expect(data.content).toContain('## Connected Microsoft 365 Services');
    expect(data.content).toContain('Behind the scenes:');
    expect(data.content).not.toContain('\n  Tools:');
    expect(data.content).toContain('Find information');
    expect(data.content).toContain('Create with guided forms');
    expect(data.content).toContain('catalog fallback');
    expect(data.content).toContain('Common required inputs:');
    expect(data.content).not.toContain('permissions, and recent activity');

    const payload = JSON.parse(result.output) as {
      success: boolean;
      totalBuiltInToolCount: number;
      userFacingToolCount: number;
      internalToolCount: number;
    };
    expect(payload.success).toBe(true);
    expect(payload.totalBuiltInToolCount).toBeGreaterThan(payload.userFacingToolCount);
    expect(payload.userFacingToolCount).toBeGreaterThan(0);
    expect(payload.internalToolCount).toBeGreaterThan(0);
  });

  it('renders a focused SharePoint capability drill-down with tool-level detail', async () => {
    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.list_m365_servers({ focus: 'sharepoint' }, deps);

    expect(result.phase).toBe('complete');
    expect(store.pushBlock).toHaveBeenCalledTimes(1);

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IMarkdownData;

    expect(block.title).toBe('SharePoint Capabilities');
    expect(data.content).toContain('## SharePoint Capabilities');
    expect(data.content).toContain('### SharePoint & OneDrive');
    expect(data.content).toContain('### SharePoint Lists');
    expect(data.content).toContain('Schema source: generated catalog fallback');
    expect(data.content).toContain('Common required inputs:');
    expect(data.content).toContain('Common schema fields:');
    expect(data.content).toContain('Generic runtime help applied by Grimoire');
    expect(data.content).toContain('- Aliases:');
    expect(data.content).toContain('- Target resolution:');
    expect(data.content).toContain('- Result shaping:');
    expect(data.content).toContain('`getFolderChildren`');
    expect(data.content).toContain('Required:');
    expect(data.content).toContain('Inputs: `documentLibraryId`');
    expect(data.content).toContain('Runtime help:');
    expect(data.content).toContain('Current Limitation');

    const payload = JSON.parse(result.output) as {
      success: boolean;
      focus: string;
      focusedToolCount: number;
    };
    expect(payload.success).toBe(true);
    expect(payload.focus).toBe('sharepoint');
    expect(payload.focusedToolCount).toBeGreaterThan(0);
  });
});

describe('ToolRegistryRuntimeHandlersUiAndPersonal.show_permissions', () => {
  it('fails honestly when the connected ODSP MCP server lacks a permission-read tool', async () => {
    const odspConnection: IMcpConnection = {
      sessionId: 'session-1',
      serverUrl: 'https://example.com/mcp_ODSPRemoteServer',
      serverName: 'SharePoint & OneDrive',
      state: 'connected',
      connectedAt: new Date(),
      tools: [
        {
          name: 'shareFileOrFolder',
          description: 'Grant read/write permissions on a file or folder.',
          inputSchema: { type: 'object', properties: {} }
        },
        {
          name: 'getFileOrFolderMetadataByUrl',
          description: 'Get metadata for a file or folder by URL.',
          inputSchema: { type: 'object', properties: { fileOrFolderUrl: { type: 'string' } }, required: ['fileOrFolderUrl'] }
        }
      ]
    };
    const store = createStore({
      proxyConfig: {
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test-key',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      mcpConnections: [odspConnection]
    });
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_permissions({
      target_name: 'SPFx',
      target_url: 'https://contoso.sharepoint.com/sites/copilot-test/Freigegebene Dokumente/SPFx.pdf'
    }, deps);

    expect(result.phase).toBe('error');
    expect(store.pushBlock).toHaveBeenCalledTimes(1);

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IErrorData;

    expect(block.type).toBe('error');
    expect(block.title).toBe('Permissions: SPFx');
    expect(data.message).toContain('"SPFx"');
    expect(data.message).toContain('does not expose a permission-inspection tool');
    expect(data.detail).toContain('sharing actions');
  });

});

describe('ToolRegistryRuntimeHandlersUiAndPersonal.show_compose_form', () => {
  beforeEach(() => {
    useGrimoireStore.setState({
      blocks: [],
      transcript: [],
      activeActionBlockId: undefined,
      selectedActionIndices: []
    });
    (hybridInteractionEngine.getCurrentTaskContext as jest.Mock).mockReset().mockReturnValue(undefined);
    (hybridInteractionEngine.captureCurrentSourceContext as jest.Mock).mockReset().mockReturnValue(undefined);
    (hybridInteractionEngine.getCurrentArtifacts as jest.Mock).mockReset().mockReturnValue({});
    jest.restoreAllMocks();
  });

  it('hydrates email compose forms with detailed visible content and attachments', async () => {
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'i am searching for doc about spfx',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx_de',
          summary: 'German SPFx guidance.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx_ja',
          summary: 'Japanese SPFx guidance.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
          author: 'SharePoint-App',
          fileType: 'pdf',
          sources: ['sharepoint-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 2 documents. They are in the panel.', timestamp: new Date('2026-03-08T10:00:00.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: [2]
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Send by Email',
      prefill_json: JSON.stringify({
        body: 'I found these SPFx-related documents in the site https://tenant.sharepoint.com/sites/dev:\n\n1) SPFx_de (pdf) — author: Test User\n2) SPFx_ja (pdf) — author: SharePoint-App\n\nThey are listed in the action panel — please attach the ones you want to include or review them before sending.'
      })
    }, deps);

    expect(result.phase).toBe('complete');
    expect(store.pushBlock).toHaveBeenCalledTimes(1);

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const subjectField = data.fields.find((field) => field.key === 'subject');
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(subjectField?.defaultValue).toBe('Grimoire share — Search: SPFx');
    expect(bodyField?.defaultValue).toContain('I found these SPFx-related documents in the site');
    expect(bodyField?.defaultValue).toContain('This recap covers all 2 visible results');
    expect(bodyField?.defaultValue).toContain('SPFx_ja');
    expect(bodyField?.defaultValue).toContain('SPFx_de');
    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf']
    });
  });

  it('prefers HIE source selection when compose context is no longer mirrored in the store', async () => {
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'i am searching for doc about spfx',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx_de',
          summary: 'German overview',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx_ja',
          summary: 'Japanese overview',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
          author: 'SharePoint-App',
          fileType: 'pdf',
          sources: ['sharepoint-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [],
      activeActionBlockId: undefined,
      selectedActionIndices: []
    });
    (hybridInteractionEngine.captureCurrentSourceContext as jest.Mock).mockReturnValue({
      sourceBlockId: searchBlock.id,
      selectedItems: [{
        index: 2,
        title: 'SPFx_ja',
        url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf'
      }]
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Send by Email'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(bodyField?.defaultValue).toContain('This recap covers all 2 visible results');
    expect(bodyField?.defaultValue).toContain('SPFx_ja');
    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf']
    });
  });

  it('infers the first visible document from the latest user transcript for voice-style email compose', async () => {
    const searchBlock = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'animals',
      totalCount: 4,
      source: 'copilot-search',
      results: [
        {
          title: 'Animal Stories and Facts',
          summary: 'Animal stories overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Histoires et Faits sur les Animaux',
          summary: 'French animal stories.',
          url: 'https://tenant.sharepoint.com/sites/dev/Histoires%20et%20Faits%20sur%20les%20Animaux.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Tiergeschichten und Fakten',
          summary: 'German animal stories.',
          url: 'https://tenant.sharepoint.com/sites/dev/Tiergeschichten%20und%20Fakten.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Animal Library Guide',
          summary: 'Guide page.',
          url: 'https://tenant.sharepoint.com/sites/dev/Animal-Library-Guide.aspx',
          author: 'System Account',
          fileType: 'aspx',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 4 documents. They are in the panel.', timestamp: new Date('2026-03-11T08:00:00.000Z') },
        { role: 'user', text: 'send first document by email', timestamp: new Date('2026-03-11T08:00:05.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Share by Email'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(bodyField?.defaultValue).toContain('This recap covers all 4 visible results');
    expect(bodyField?.defaultValue).toContain('Animal Stories and Facts');
    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx']
    });
  });

  it('does not widen explicit narrow voice share requests when no current selection exists', async () => {
    const searchBlock = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'animals',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'Animal Stories and Facts',
          summary: 'Animal stories overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Histoires et Faits sur les Animaux',
          summary: 'French animal stories.',
          url: 'https://tenant.sharepoint.com/sites/dev/Histoires%20et%20Faits%20sur%20les%20Animaux.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 2 documents. They are in the panel.', timestamp: new Date('2026-03-11T08:15:00.000Z') },
        { role: 'user', text: 'send selected document by email', timestamp: new Date('2026-03-11T08:15:05.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Share by Email'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(bodyField?.defaultValue).toContain('This recap covers all 2 visible results');
    expect(data.submissionTarget.staticArgs).toEqual({});
  });

  it('narrows compose hydration to a referenced search result from static args', async () => {
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'spfx',
      totalCount: 5,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx_ja',
          summary: 'Japanese SPFx guidance.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
          author: 'SharePoint-App',
          fileType: 'pdf',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx_de',
          summary: 'German SPFx guidance.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf',
          author: 'SharePoint-App',
          fileType: 'pdf',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx',
          summary: 'SPFx overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['copilot-search']
        },
        {
          title: 'TeamsFx',
          summary: 'TeamsFx overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/TeamsFx.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['copilot-search']
        },
        {
          title: 'Power Platform',
          summary: 'Power Platform overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['sharepoint-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 5 documents. They are in the panel.', timestamp: new Date('2026-03-09T13:00:00.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Share by Email',
      description: 'Pre-filled to send the selected item. Add recipients and send when ready.',
      static_args_json: JSON.stringify({
        attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf'],
        shareSelectionIndices: [5],
        shareBlockId: searchBlock.id,
        shareItemTitle: 'Power Platform',
        fileOrFolderUrl: 'https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf',
        fileOrFolderName: 'Power Platform'
      })
    }, deps);

    expect(result.phase).toBe('complete');
    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(data.description).toBeUndefined();
    expect(bodyField?.defaultValue).toContain('Power Platform');
    expect(bodyField?.defaultValue).toContain('This recap covers all 5 visible results');
    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf']
    });
    expect(data.submissionTarget.targetContext).toMatchObject({
      fileOrFolderUrl: 'https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf',
      fileOrFolderName: 'Power Platform'
    });
  });

  it('resolves attachment titles from voice-style compose prefill into real attachment URIs', async () => {
    const searchBlock = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'animals',
      totalCount: 3,
      source: 'copilot-search',
      results: [
        {
          title: 'Animal Stories and Facts',
          summary: 'Animal stories overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Histoires et Faits sur les Animaux',
          summary: 'French animal stories.',
          url: 'https://tenant.sharepoint.com/sites/dev/Histoires%20et%20Faits%20sur%20les%20Animaux.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Tiergeschichten und Fakten',
          summary: 'German animal stories.',
          url: 'https://tenant.sharepoint.com/sites/dev/Tiergeschichten%20und%20Fakten.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 3 documents. They are in the panel.', timestamp: new Date('2026-03-10T19:00:00.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Send Animal Documents by Email',
      prefill_json: JSON.stringify({
        subject: 'Animal Documents',
        body: 'Sharing these animal-related documents.',
        attachments: 'Animal Stories and Facts; Histoires et Faits sur les Animaux; Tiergeschichten und Fakten'
      })
    }, deps);

    expect(result.phase).toBe('complete');
    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;

    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: [
        'https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx',
        'https://tenant.sharepoint.com/sites/dev/Histoires%20et%20Faits%20sur%20les%20Animaux.docx',
        'https://tenant.sharepoint.com/sites/dev/Tiergeschichten%20und%20Fakten.docx'
      ]
    });
  });

  it('builds reply-all thread compose forms with prefilled recap content and preserved message id', async () => {
    const recapBlock = createBlock('markdown', 'Nova recap', {
      kind: 'markdown',
      content: 'Nova recap body'
    } as IMarkdownData);
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'Recap ready in the panel.', timestamp: new Date('2026-03-11T19:00:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-reply-all-thread',
      title: 'Reply to Mail Discussion',
      static_args_json: JSON.stringify({
        messageId: 'mail-123',
        includeOriginalNonInlineAttachments: false
      })
    }, deps);

    expect(result.phase).toBe('complete');
    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const introCommentField = data.fields.find((field) => field.key === 'introComment');

    expect(introCommentField?.defaultValue).toContain('Nova recap body');
    expect(data.submissionTarget.toolName).toBe('ReplyAllWithFullThread');
    expect(data.submissionTarget.serverId).toBe('mcp_MailTools');
    expect(data.submissionTarget.staticArgs).toEqual({
      messageId: 'mail-123',
      includeOriginalNonInlineAttachments: false
    });
    expect(data.submissionTarget.targetContext).toMatchObject({
      mailItemId: 'mail-123'
    });
  });

  it('skips session hydration for reply-all thread workflows when the recap intro is provided explicitly', async () => {
    useGrimoireStore.setState({
      blocks: [
        createBlock('search-results', 'Search: Nova launch recap', {
          kind: 'search-results',
          query: 'nova launch recap',
          totalCount: 1,
          source: 'copilot-search',
          results: [
            {
              title: 'Nova_Launch_Recap',
              summary: 'Launch recap document.',
              url: 'https://tenant.sharepoint.com/sites/dev/Nova_Launch_Recap.docx',
              author: 'Anna Muller',
              fileType: 'docx',
              sources: ['copilot-search']
            }
          ]
        } as ISearchResultsData)
      ],
      transcript: [
        { role: 'assistant', text: 'I found 1 document. It is in the panel.', timestamp: new Date('2026-03-11T19:05:00.000Z') }
      ],
      activeActionBlockId: undefined,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-reply-all-thread',
      title: 'Reply to Mail Discussion',
      prefill_json: JSON.stringify({
        introComment: 'Project Nova launch recap summary.'
      }),
      static_args_json: JSON.stringify({
        messageId: 'mail-456',
        includeOriginalNonInlineAttachments: false,
        skipSessionHydration: true
      })
    }, deps);

    expect(result.phase).toBe('complete');
    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const introCommentField = data.fields.find((field) => field.key === 'introComment');

    expect(introCommentField?.defaultValue).toBe('Project Nova launch recap summary.');
    expect(data.submissionTarget.staticArgs).toEqual({
      messageId: 'mail-456',
      includeOriginalNonInlineAttachments: false
    });
  });

  it('normalizes voice-style attachmentUrls and keeps compose hydration narrowed to the explicit result', async () => {
    const searchBlock = createBlock('search-results', 'Search: document about animals', {
      kind: 'search-results',
      query: 'document about animals',
      totalCount: 4,
      source: 'copilot-search',
      results: [
        {
          title: 'Animal Stories and Facts',
          summary: 'Animal stories overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Histoires et Faits sur les Animaux',
          summary: 'French animal stories.',
          url: 'https://tenant.sharepoint.com/sites/dev/Histoires%20et%20Faits%20sur%20les%20Animaux.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Farm_to_Table_Sustainable_Cooking_EN',
          summary: 'Cooking document.',
          url: 'https://tenant.sharepoint.com/sites/dev/Farm_to_Table_Sustainable_Cooking_EN.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'How To Use This Library',
          summary: 'Library help page.',
          url: 'https://tenant.sharepoint.com/sites/dev/How-To-Use-This-Library.aspx',
          author: 'System Account',
          fileType: 'aspx',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 4 documents. They are in the panel.', timestamp: new Date('2026-03-10T20:00:00.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: "Share 'Animal Stories and Facts'",
      description: 'Send this document as an email attachment.',
      prefill_json: JSON.stringify({
        subject: 'Sharing: Animal Stories and Facts',
        body: "Please find attached the document 'Animal Stories and Facts'."
      }),
      static_args_json: JSON.stringify({
        attachmentUrls: ['https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx']
      })
    }, deps);

    expect(result.phase).toBe('complete');
    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Animal%20Stories%20and%20Facts.docx']
    });
    expect(bodyField?.defaultValue).toContain('Animal Stories and Facts');
    expect(bodyField?.defaultValue).toContain('This recap covers all 4 visible results');
  });

  it('normalizes mailto-style encoded newlines in compose prefill text', async () => {
    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Encoded Email',
      prefill_json: JSON.stringify({
        body: 'Hi,%0D%0A%0D%0AHere are the results.%0D%0AThanks'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(bodyField?.defaultValue).toContain('Hi,\n\nHere are the results.\nThanks');
    expect(bodyField?.defaultValue).not.toContain('%0D%0A');
  });

  it('accepts voice-style object compose args and falls back from tool-name share ids to the latest document block', async () => {
    (logService.warning as jest.Mock).mockClear();

    const searchBlock = createBlock('search-results', 'Search: Nova Launch recap', {
      kind: 'search-results',
      query: 'nova launch recap',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'Nova_Launch_Recap',
          summary: 'Launch recap document.',
          url: 'https://tenant.sharepoint.com/sites/dev/Nova_Launch_Recap.docx',
          author: 'Alice Smith',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'Nova_Launch_Budget',
          summary: 'Budget spreadsheet.',
          url: 'https://tenant.sharepoint.com/sites/dev/Nova_Launch_Budget.xlsx',
          author: 'Bob Jones',
          fileType: 'xlsx',
          sources: ['sharepoint-search']
        }
      ]
    } as ISearchResultsData);
    const userCardBlock = createBlock('user-card', 'Alice Smith', {
      kind: 'user-card',
      displayName: 'Alice Smith',
      email: 'alice.smith@contoso.onmicrosoft.com',
      title: 'Head of Marketing'
    } as IUserCardData);
    useGrimoireStore.setState({
      blocks: [searchBlock, userCardBlock],
      transcript: [
        { role: 'assistant', text: 'Summary is ready in the panel.', timestamp: new Date('2026-03-11T20:57:00.000Z') }
      ],
      activeActionBlockId: userCardBlock.id,
      selectedActionIndices: [1]
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Email Nova_Launch_Recap to a colleague',
      description: 'Compose your message and I\'ll send the Nova_Launch_Recap document as an attachment.',
      prefill_json: {
        to: [
          'alice.smith@contoso.onmicrosoft.com',
          'bob.jones@contoso.onmicrosoft.com'
        ],
        subject: 'Nova Launch Recap',
        body: 'Hi team,\n\nSharing the recap of the Nova launch as requested.\n\nBest,\nTest'
      },
      static_args_json: {
        shareSelectionIndices: [1],
        shareBlockId: 'search_sharepoint',
        shareScopeMode: 'specific'
      }
    }, deps);

    expect(result.phase).toBe('complete');
    expect(logService.warning).not.toHaveBeenCalledWith('llm', 'Invalid prefill_json for show_compose_form');
    expect(logService.warning).not.toHaveBeenCalledWith('llm', 'Invalid static_args_json for show_compose_form');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const toField = data.fields.find((field) => field.key === 'to');
    const bodyField = data.fields.find((field) => field.key === 'body');

    expect(toField?.defaultValue).toBe('alice.smith@contoso.onmicrosoft.com; bob.jones@contoso.onmicrosoft.com');
    expect(bodyField?.defaultValue).toContain('Sharing the recap of the Nova launch as requested.');
    expect(bodyField?.defaultValue).toContain('Nova_Launch_Recap');
    expect(bodyField?.defaultValue).toContain('This recap covers all 2 visible results');
    expect(bodyField?.defaultValue).not.toContain('Head of Marketing');
    expect(bodyField?.defaultValue).not.toContain('alice.smith@contoso.onmicrosoft.com\n   Role');
    expect(data.submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Nova_Launch_Recap.docx']
    });
  });

  it('hydrates non-email share forms from the current pane content', async () => {
    const searchBlock = createBlock('search-results', 'Search: Calendar Follow-up', {
      kind: 'search-results',
      query: 'recent project updates',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'Project kickoff notes',
          summary: 'Meeting recap and next steps.',
          url: 'https://tenant.sharepoint.com/sites/dev/KickoffNotes.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: [1]
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'share-teams-chat',
      title: 'Share to Teams',
      prefill_json: JSON.stringify({
        content: 'Please review.'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const topicField = data.fields.find((field) => field.key === 'topic');
    const contentField = data.fields.find((field) => field.key === 'content');

    expect(topicField?.defaultValue).toBe('Grimoire share — Search: Calendar Follow-up');
    expect(contentField?.defaultValue).toContain('Please review.');
    expect(contentField?.defaultValue).toContain('This recap covers all 1 visible results');
    expect(contentField?.defaultValue).toContain('Project kickoff notes');
    expect(data.submissionTarget.staticArgs).toEqual({});
  });

  it('normalizes word-document create forms and pre-fills them from the active recap block', async () => {
    const recapBlock = createBlock('info-card', 'Recap: Search: nova launch', {
      kind: 'info-card',
      heading: 'Recap: Search: nova launch',
      body: [
        'Nova launch results center on detailing the Nova analytics workspace rollout, its timing, risks, budget, and positioning.',
        '',
        '- Strongest items: Nova_Launch_Recap shows a June 15, 2026 launch date with status Successful and adoption signals.',
        '- Top content focus: strategy and recap documents plus a budget sheet.'
      ].join('\n')
    } as IInfoCardData);

    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'Recap ready in the panel.', timestamp: new Date('2026-03-14T08:31:00.000Z') },
        { role: 'user', text: 'create a word document with the recap in the document library mydoclib', timestamp: new Date('2026-03-14T08:32:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'file-create',
      title: 'Create Word document: Recap - Nova launch',
      description: 'Create a .docx in the mydoclib document library containing the visible recap.',
      static_args_json: JSON.stringify({
        siteUrl: 'https://tenant.sharepoint.com/sites/ProjectNova',
        siteName: 'ProjectNova',
        documentLibraryName: 'mydoclib'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const filenameField = data.fields.find((field) => field.key === 'filename');
    const contentField = data.fields.find((field) => field.key === 'contentText');

    expect(data.preset).toBe('word-document-create');
    expect(filenameField?.label).toBe('Document name');
    expect(filenameField?.placeholder).toBe('Document.docx');
    expect(filenameField?.defaultValue).toBe('Recap - Nova launch.docx');
    expect(contentField?.defaultValue).toContain('Nova launch results center on detailing the Nova analytics workspace rollout');
    expect(contentField?.defaultValue).toContain('Strongest items: Nova_Launch_Recap');
    expect(contentField?.defaultValue).not.toContain('Shared from Grimoire.');
    expect(data.submissionTarget.toolName).toBe('createSmallBinaryFile');
    expect(data.submissionTarget.serverId).toBe('mcp_ODSPRemoteServer');
    expect(data.submissionTarget.targetContext).toMatchObject({
      siteUrl: 'https://tenant.sharepoint.com/sites/ProjectNova',
      siteName: 'ProjectNova',
      documentLibraryName: 'mydoclib'
    });
  });

  it('normalizes legacy create-file aliases and prefers the recap artifact body over the source search block', async () => {
    const searchBlock = createBlock('search-results', 'Search: launch recap', {
      kind: 'search-results',
      query: 'launch recap',
      totalCount: 2,
      source: 'sharepoint-search',
      results: [
        {
          title: 'Nova_Launch_Recap',
          summary: 'The visible results are focused on launch recap source material.',
          url: 'https://tenant.sharepoint.com/sites/ProjectNova/Shared%20Documents/Nova_Launch_Recap.docx',
          author: 'Anna Muller',
          fileType: 'docx',
          sources: ['sharepoint-search']
        },
        {
          title: 'Nova_Launch_Strategy',
          summary: 'Strategy notes for the Nova launch.',
          url: 'https://tenant.sharepoint.com/sites/ProjectNova/Shared%20Documents/Nova_Launch_Strategy.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['sharepoint-search']
        }
      ]
    } as ISearchResultsData);
    const recapBlock = createBlock('info-card', 'Recap: Search: launch recap', {
      kind: 'info-card',
      heading: 'Recap: Search: launch recap',
      body: [
        'Nova launch results center on detailing the Nova analytics workspace rollout, its timing, risks, budget, and positioning.',
        '',
        '- Strongest items: Nova_Launch_Recap and Nova_Launch_Strategy.',
        '- Top content focus: strategy and recap documents plus a budget sheet.'
      ].join('\n')
    } as IInfoCardData);

    useGrimoireStore.setState({
      blocks: [searchBlock, recapBlock],
      transcript: [
        { role: 'assistant', text: 'Recap ready in the panel.', timestamp: new Date('2026-03-14T09:13:00.000Z') },
        { role: 'user', text: 'create a document named testrecap with the recap and save it in the document library mydoclib', timestamp: new Date('2026-03-14T09:14:00.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: []
    });

    (hybridInteractionEngine.getCurrentTaskContext as jest.Mock).mockReturnValue({
      kind: 'recap',
      eventName: 'artifact.recap.ready',
      sourceBlockId: searchBlock.id,
      sourceBlockType: searchBlock.type,
      sourceBlockTitle: searchBlock.title,
      derivedBlockId: recapBlock.id,
      derivedBlockType: recapBlock.type,
      derivedBlockTitle: recapBlock.title,
      targetContext: undefined,
      updatedAt: Date.now()
    });
    (hybridInteractionEngine.captureCurrentSourceContext as jest.Mock).mockReturnValue({
      sourceBlockId: searchBlock.id,
      sourceBlockType: searchBlock.type,
      sourceBlockTitle: searchBlock.title,
      sourceArtifactId: recapBlock.id,
      sourceTaskKind: 'recap',
      sourceEventName: 'artifact.recap.ready'
    });
    (hybridInteractionEngine.getCurrentArtifacts as jest.Mock).mockReturnValue({
      [recapBlock.id]: {
        artifactId: recapBlock.id,
        artifactKind: 'recap',
        sourceBlockId: searchBlock.id,
        blockId: recapBlock.id,
        blockType: recapBlock.type,
        title: recapBlock.title,
        status: 'ready',
        updatedAt: Date.now()
      }
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'file-create',
      title: "Create file 'testrecap' in mydoclib",
      description: 'Saving the latest recap artifact into the specified document library. Review and submit to create the file.',
      prefill_json: JSON.stringify({
        site_url: 'https://tenant.sharepoint.com/sites/ProjectNova',
        library_name: 'mydoclib',
        file_name: 'testrecap.docx',
        file_content: 'Recap: Search: launch recap'
      }),
      static_args_json: JSON.stringify({
        targetSiteUrl: 'https://tenant.sharepoint.com/sites/ProjectNova',
        targetLibrary: 'mydoclib'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const filenameField = data.fields.find((field) => field.key === 'filename');
    const contentField = data.fields.find((field) => field.key === 'contentText');

    expect(data.preset).toBe('word-document-create');
    expect(filenameField?.defaultValue).toBe('testrecap.docx');
    expect(contentField?.defaultValue).toContain('Nova launch results center on detailing the Nova analytics workspace rollout');
    expect(contentField?.defaultValue).toContain('Strongest items: Nova_Launch_Recap and Nova_Launch_Strategy');
    expect(contentField?.defaultValue).not.toContain('The visible results are focused on launch recap source material');
    expect(contentField?.defaultValue).not.toContain('Recap: Search: launch recap');
    expect(data.submissionTarget.targetContext).toMatchObject({
      siteUrl: 'https://tenant.sharepoint.com/sites/ProjectNova',
      documentLibraryName: 'mydoclib'
    });
  });

  it('sanitizes invalid SharePoint characters in prefilled word-document filenames', async () => {
    useGrimoireStore.setState({
      blocks: [],
      transcript: [],
      activeActionBlockId: undefined,
      selectedActionIndices: []
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'file-create',
      title: 'Create Word document in mydoclib',
      description: 'Creates a .docx in site Project Nova › mydoclib with the latest recap.',
      prefill_json: JSON.stringify({
        fileName: 'Recap - Search: launch recap.docx',
        fileContent: 'Recap: Search: launch recap'
      }),
      static_args_json: JSON.stringify({
        siteUrl: 'https://tenant.sharepoint.com/sites/ProjectNova',
        libraryName: 'mydoclib',
        fileFormat: 'docx'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const filenameField = data.fields.find((field) => field.key === 'filename');

    expect(data.preset).toBe('word-document-create');
    expect(filenameField?.defaultValue).toBe('Recap - Search launch recap.docx');
  });

  it('does not inject visible pane context into event creation forms by default', async () => {
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'spfx',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx',
          summary: 'SharePoint Framework overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [
        { role: 'assistant', text: 'I found 1 document. It is in the panel.', timestamp: new Date('2026-03-08T11:00:00.000Z') }
      ],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: [1]
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'event-create',
      title: 'Create appointment',
      prefill_json: JSON.stringify({
        subject: 'Appointment tomorrow',
        bodyContent: 'Please prepare the agenda.'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const subjectField = data.fields.find((field) => field.key === 'subject');
    const bodyField = data.fields.find((field) => field.key === 'bodyContent');

    expect(subjectField?.defaultValue).toBe('Appointment tomorrow');
    expect(bodyField?.defaultValue).toBe('Please prepare the agenda.');
    expect(bodyField?.defaultValue).not.toContain('Current visible content');
    expect(bodyField?.defaultValue).not.toContain('SPFx');
  });

  it('persists explicit HIE target context onto compose form submission targets', async () => {
    (hybridInteractionEngine.getCurrentTaskContext as jest.Mock).mockReturnValue({
      kind: 'select',
      eventName: 'block.interaction.select',
      selectedItems: [],
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      },
      updatedAt: Date.now()
    });

    const store = createStore({
      userContext: {
        displayName: 'Test User',
        email: 'test.user@example.com',
        loginName: 'test.user@example.com',
        resolvedLanguage: 'en',
        currentWebTitle: 'copilot-test',
        currentWebUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
        currentSiteTitle: 'copilot-test',
        currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
      }
    });
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'folder-create',
      title: 'Create folder in Documents'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;

    expect(data.submissionTarget.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      siteName: 'copilot-test-cooking'
    }));
  });

  it('derives folder-create target context from the latest site artifact when no task target exists', async () => {
    (hybridInteractionEngine.getCurrentArtifacts as jest.Mock).mockReturnValue({
      'artifact-site': {
        artifactId: 'artifact-site',
        artifactKind: 'block',
        blockId: 'block-site',
        blockType: 'site-info',
        title: 'copilot-test-cooking',
        status: 'ready',
        updatedAt: Date.now(),
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
          siteName: 'copilot-test-cooking'
        }
      }
    });

    const store = createStore({
      userContext: {
        displayName: 'Test User',
        email: 'test.user@example.com',
        loginName: 'test.user@example.com',
        resolvedLanguage: 'en',
        currentWebTitle: 'copilot-test',
        currentWebUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
        currentSiteTitle: 'copilot-test',
        currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
      }
    });
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'folder-create',
      title: 'Create folder in Documents'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;

    expect(data.submissionTarget.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      siteName: 'copilot-test-cooking'
    }));
  });

  it('persists selected source-context item targets onto share compose forms', async () => {
    const store = createStore({
      userContext: {
        displayName: 'Test User',
        email: 'test.user@example.com',
        loginName: 'test.user@example.com',
        resolvedLanguage: 'en',
        currentWebTitle: 'copilot-test',
        currentWebUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
        currentSiteTitle: 'copilot-test',
        currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
      }
    });
    const deps: IToolRuntimeHandlerDeps = {
      ...createDeps(store),
      sourceContext: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx',
        sourceTaskKind: 'click-result',
        selectedItems: [],
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
          siteName: 'copilot-test-cooking',
          fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/SPFx.pdf',
          fileOrFolderName: 'SPFx.pdf'
        }
      }
    };
    const handlers = buildUiAndPersonalRuntimeHandlers();

    const result = await handlers.show_compose_form({
      preset: 'email-compose',
      title: 'Share result via email'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;

    expect(data.submissionTarget.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/SPFx.pdf',
      fileOrFolderName: 'SPFx.pdf'
    }));
  });

  it('normalizes unresolved Teams channel posts to the share preset with destination fields', async () => {
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'spfx',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx',
          summary: 'SharePoint Framework overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
          author: 'Test User',
          fileType: 'pdf',
          sources: ['copilot-search']
        }
      ]
    } as ISearchResultsData);
    useGrimoireStore.setState({
      blocks: [searchBlock],
      transcript: [],
      activeActionBlockId: searchBlock.id,
      selectedActionIndices: [1]
    });

    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();
    jest.spyOn(pnpContext, 'isInitialized').mockReturnValue(false);
    jest.spyOn(shareSubmissionService, 'loadTeamsChannelDestinationOptions').mockResolvedValue([
      {
        key: JSON.stringify({
          teamId: 'team-1',
          channelId: 'channel-1',
          teamName: 'Engineering',
          channelName: 'General'
        }),
        text: 'Engineering / General'
      }
    ]);

    const result = await handlers.show_compose_form({
      preset: 'teams-channel-message',
      title: 'Post to Channel',
      prefill_json: JSON.stringify({
        message: 'Hi team,\n\nPlease review these SPFx documents.'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const destinationField = data.fields.find((field) => field.key === 'destination');
    const contentField = data.fields.find((field) => field.key === 'content');

    expect(data.preset).toBe('share-teams-channel');
    expect(data.submissionTarget.toolName).toBe('share_teams_channel');
    expect(destinationField?.type).toBe('dropdown');
    expect(destinationField?.options).toEqual([
      {
        key: JSON.stringify({
          teamId: 'team-1',
          channelId: 'channel-1',
          teamName: 'Engineering',
          channelName: 'General'
        }),
        text: 'Engineering / General'
      }
    ]);
    expect(contentField?.defaultValue).toContain('Hi team,');
    expect(contentField?.defaultValue).toContain('This recap covers all 1 visible results');
    expect(contentField?.defaultValue).toContain('SPFx');
  });

  it('keeps Teams picker fields when the SPFx context is initialized', async () => {
    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();
    jest.spyOn(pnpContext, 'isInitialized').mockReturnValue(true);

    const result = await handlers.show_compose_form({
      preset: 'share-teams-channel',
      title: 'Post to Channel',
      description: 'Choose the destination channel, review the message, then send.'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;

    expect(data.fields.find((field) => field.key === 'teamId')?.type).toBe('team-picker');
    expect(data.fields.find((field) => field.key === 'channelId')?.type).toBe('channel-picker');
    expect(data.fields.find((field) => field.key === 'teamName')?.type).toBe('hidden');
    expect(data.fields.find((field) => field.key === 'channelName')?.type).toBe('hidden');
    expect(data.description).toBe('Choose the destination channel, review the message, then send.');
  });

  it('shows a visible manual-entry notice when Teams channel options are unavailable', async () => {
    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();
    jest.spyOn(pnpContext, 'isInitialized').mockReturnValue(false);
    jest.spyOn(shareSubmissionService, 'loadTeamsChannelDestinationOptions').mockResolvedValue([]);

    const result = await handlers.show_compose_form({
      preset: 'share-teams-channel',
      title: 'Post to Channel',
      description: 'Choose the destination channel, review the message, then send.',
      prefill_json: JSON.stringify({
        content: 'Hi team,\n\nPlease review these documents.'
      })
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;
    const destinationField = data.fields.find((field) => field.key === 'destination');
    const teamField = data.fields.find((field) => field.key === 'teamName');
    const channelField = data.fields.find((field) => field.key === 'channelName');

    expect(destinationField).toBeUndefined();
    expect(teamField?.type).toBe('text');
    expect(channelField?.type).toBe('text');
    expect(data.description).toContain('Choose the destination channel, review the message, then send.');
    expect(data.description).toContain('The Teams channel picker is unavailable right now');
    expect(data.description).toContain('Names are resolved when you submit.');
  });

  it('shows the same manual-entry notice when Teams channel preload fails', async () => {
    const store = createStore();
    const deps = createDeps(store);
    const handlers = buildUiAndPersonalRuntimeHandlers();
    jest.spyOn(pnpContext, 'isInitialized').mockReturnValue(false);
    jest.spyOn(shareSubmissionService, 'loadTeamsChannelDestinationOptions').mockRejectedValue(new Error('Teams server unavailable'));

    const result = await handlers.show_compose_form({
      preset: 'share-teams-channel',
      title: 'Post to Channel',
      description: 'Choose the destination channel, review the message, then send.'
    }, deps);

    expect(result.phase).toBe('complete');

    const block = (store.pushBlock as jest.Mock).mock.calls[0][0];
    const data = block.data as IFormData;

    expect(data.fields.find((field) => field.key === 'destination')).toBeUndefined();
    expect(data.fields.find((field) => field.key === 'teamName')?.type).toBe('text');
    expect(data.fields.find((field) => field.key === 'channelName')?.type).toBe('text');
    expect(data.description).toContain('The Teams channel picker is unavailable right now');
    expect(data.description).toContain('Names are resolved when you submit.');
  });
});
