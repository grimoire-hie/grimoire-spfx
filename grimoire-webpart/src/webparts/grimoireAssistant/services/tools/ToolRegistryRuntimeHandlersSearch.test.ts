jest.mock('../../store/useGrimoireStore', () => ({
  useGrimoireStore: {
    getState: jest.fn()
  }
}));

jest.mock('@microsoft/sp-http', () => ({
  AadHttpClient: {
    configurations: {
      v1: {}
    }
  }
}));

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    getCurrentTaskContext: jest.fn(() => undefined),
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

jest.mock('../pnp/pnpConfig', () => ({
  getSP: jest.fn()
}));

import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';
import { buildSearchRuntimeHandlers } from './ToolRegistryRuntimeHandlersSearch';
import { CopilotSearchService } from '../search/CopilotSearchService';
import { CopilotRetrievalService } from '../search/CopilotRetrievalService';
import { SharePointSearchService } from '../search/SharePointSearchService';
import { SearchIntentPlanner } from '../search/SearchIntentPlanner';
import type { ICopilotSearchResult } from '../../models/ISearchTypes';
import { PublicWebSearchService } from '../web/PublicWebSearchService';
import * as McpExecutionAdapter from '../mcp/McpExecutionAdapter';

interface IStoreStateForTests {
  updateBlock: jest.Mock;
  pushBlock: jest.Mock;
  setPublicWebSearchCapability: jest.Mock;
  proxyConfig?: IFunctionCallStore['proxyConfig'];
}

function createStore(overrides: Partial<IFunctionCallStore> = {}): IFunctionCallStore {
  return {
    aadHttpClient: {} as never,
    proxyConfig: undefined,
    getToken: undefined,
    mcpEnvironmentId: undefined,
    userContext: {
      displayName: 'Test User',
      email: 'test.user@example.com',
      loginName: 'test.user@example.com',
      resolvedLanguage: 'fr',
      currentWebTitle: 'Finance',
      currentWebUrl: 'https://tenant.sharepoint.com/sites/finance',
      currentSiteTitle: 'Finance',
      currentSiteUrl: 'https://tenant.sharepoint.com/sites/finance'
    },
    publicWebSearchEnabled: true,
    publicWebSearchCapability: 'unknown',
    copilotWebGroundingEnabled: false,
    setPublicWebSearchCapability: jest.fn(),
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
    awaitAsync: true,
    aadClient: store.aadHttpClient,
    sitesService: undefined,
    peopleService: undefined
  };
}

function makeResult(id: string, source: ICopilotSearchResult['source']): ICopilotSearchResult {
  return {
    title: `Doc ${id}`,
    summary: `Summary ${id}`,
    url: `https://tenant.sharepoint.com/sites/finance/Shared%20Documents/doc-${id}.docx`,
    fileType: 'docx',
    author: 'Test User',
    siteName: 'Finance',
    source
  };
}

describe('ToolRegistryRuntimeHandlersSearch.search_sharepoint', () => {
  const getStateMock = useGrimoireStore.getState as jest.Mock;

  let state: IStoreStateForTests;
  let searchSpy: jest.SpyInstance;
  let retrievalSpy: jest.SpyInstance;
  let sharePointSpy: jest.SpyInstance;
  let plannerSpy: jest.SpyInstance;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.restoreAllMocks();

    state = {
      updateBlock: jest.fn(),
      pushBlock: jest.fn(),
      setPublicWebSearchCapability: jest.fn()
    };
    getStateMock.mockImplementation(() => state);

    searchSpy = jest.spyOn(CopilotSearchService.prototype, 'search').mockResolvedValue({
      success: true,
      data: []
    });
    retrievalSpy = jest.spyOn(CopilotRetrievalService.prototype, 'search').mockResolvedValue({
      success: true,
      data: []
    });
    sharePointSpy = jest.spyOn(SharePointSearchService.prototype, 'search').mockResolvedValue({
      success: true,
      data: []
    });
    plannerSpy = jest.spyOn(SearchIntentPlanner.prototype, 'plan').mockResolvedValue({
      rawQuery: 'budget report',
      queryLanguage: 'en',
      usedCorrection: false,
      usedTranslation: false
    });
  });

  afterEach(() => {
    searchSpy.mockRestore();
    retrievalSpy.mockRestore();
    sharePointSpy.mockRestore();
    plannerSpy.mockRestore();
  });

  it('preserves the raw query across primary semantic branches', async () => {
    plannerSpy.mockResolvedValue({
      rawQuery: 'find the budget report',
      sharePointLexicalQuery: 'budget report',
      sharePointLexicalConfidence: 0.92,
      queryLanguage: 'en',
      usedCorrection: false,
      usedTranslation: false
    });
    searchSpy.mockResolvedValue({
      success: true,
      data: [makeResult('1', 'copilot-search'), makeResult('2', 'copilot-search')]
    });
    retrievalSpy.mockResolvedValue({
      success: true,
      data: [makeResult('3', 'copilot-retrieval'), makeResult('4', 'copilot-retrieval')]
    });
    sharePointSpy.mockResolvedValue({
      success: true,
      data: [makeResult('5', 'sharepoint-search'), makeResult('6', 'sharepoint-search')]
    });

    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();

    await handlers.search_sharepoint({ query: 'find the budget report' }, createDeps(store));

    expect(searchSpy).toHaveBeenCalledTimes(1);
    expect(searchSpy).toHaveBeenCalledWith('find the budget report', 20);
    expect(retrievalSpy).toHaveBeenCalledTimes(1);
    expect(retrievalSpy).toHaveBeenCalledWith('find the budget report', 'sharePoint', 10);
    expect(sharePointSpy).toHaveBeenCalledTimes(1);
    expect(sharePointSpy.mock.calls[0][0]).toBe('find the budget report');
    expect(sharePointSpy.mock.calls[0][1]).toEqual(expect.objectContaining({
      requestQueryText: 'budget report'
    }));
    expect(plannerSpy).toHaveBeenCalledWith('find the budget report', expect.objectContaining({ userLanguage: 'fr' }));

    const updateArgs = state.updateBlock.mock.calls[0][1] as {
      data: { queryVariants?: unknown };
    };
    expect(updateArgs.data.queryVariants).toBeUndefined();
  });

  it('runs corrected, translation, and keyword fallbacks adaptively when primary recall is weak', async () => {
    plannerSpy.mockResolvedValue({
      rawQuery: 'budegt report',
      sharePointLexicalQuery: 'budget report',
      sharePointLexicalConfidence: 0.94,
      correctedQuery: 'budget report',
      queryLanguage: 'en',
      translationFallbackQuery: 'rapport budget',
      translationFallbackLanguage: 'fr',
      keywordFallbackQuery: 'budget report finance',
      usedCorrection: true,
      usedTranslation: true
    });

    searchSpy
      .mockResolvedValueOnce({ success: true, data: [] })
      .mockResolvedValueOnce({ success: true, data: [] })
      .mockResolvedValueOnce({ success: true, data: [] });
    retrievalSpy
      .mockResolvedValueOnce({ success: true, data: [] })
      .mockResolvedValueOnce({ success: true, data: [] })
      .mockResolvedValueOnce({ success: true, data: [] });
    sharePointSpy
      .mockResolvedValueOnce({ success: true, data: [] })
      .mockResolvedValueOnce({ success: true, data: [] });

    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();

    await handlers.search_sharepoint({ query: 'budegt report' }, createDeps(store));

    expect(searchSpy.mock.calls.map((call) => call[0])).toEqual([
      'budegt report',
      'budget report',
      'rapport budget'
    ]);
    expect(retrievalSpy.mock.calls.map((call) => call[0])).toEqual([
      'budegt report',
      'budget report',
      'rapport budget'
    ]);
    expect(sharePointSpy.mock.calls.map((call) => call[0])).toEqual([
      'budegt report',
      'budget report finance'
    ]);
    expect(sharePointSpy.mock.calls[0][1]).toEqual(expect.objectContaining({
      requestQueryText: 'budget report'
    }));

    const updateArgs = state.updateBlock.mock.calls[0][1] as {
      data: { queryVariants?: Array<{ kind: string; query: string }> };
    };
    expect(updateArgs.data.queryVariants).toEqual([
      { kind: 'corrected', query: 'budget report', language: 'en' },
      { kind: 'translation', query: 'rapport budget', language: 'fr' },
      { kind: 'keyword-fallback', query: 'budget report finance', language: 'en' }
    ]);
  });

  it('runs a high-confidence semantic rewrite against Copilot branches before sparse-recall fallbacks', async () => {
    plannerSpy.mockResolvedValue({
      rawQuery: 'docs about animals',
      semanticRewriteQuery: 'animal-related documents',
      semanticRewriteConfidence: 0.92,
      sharePointLexicalQuery: 'animals',
      sharePointLexicalConfidence: 0.96,
      queryLanguage: 'en',
      keywordFallbackQuery: 'animals',
      usedCorrection: false,
      usedTranslation: false
    });

    searchSpy
      .mockResolvedValueOnce({
        success: true,
        data: [makeResult('1', 'copilot-search'), makeResult('2', 'copilot-search')]
      })
      .mockResolvedValueOnce({
        success: true,
        data: [makeResult('3', 'copilot-search')]
      });
    retrievalSpy
      .mockResolvedValueOnce({
        success: true,
        data: [makeResult('4', 'copilot-retrieval'), makeResult('5', 'copilot-retrieval')]
      })
      .mockResolvedValueOnce({
        success: true,
        data: [makeResult('6', 'copilot-retrieval')]
      });
    sharePointSpy.mockResolvedValue({
      success: true,
      data: [makeResult('7', 'sharepoint-search')]
    });

    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();

    await handlers.search_sharepoint({ query: 'docs about animals' }, createDeps(store));

    expect(searchSpy.mock.calls.map((call) => call[0])).toEqual([
      'docs about animals',
      'animal-related documents'
    ]);
    expect(retrievalSpy.mock.calls.map((call) => call[0])).toEqual([
      'docs about animals',
      'animal-related documents'
    ]);
    expect(sharePointSpy.mock.calls.map((call) => call[0])).toEqual([
      'docs about animals'
    ]);
    expect(sharePointSpy.mock.calls[0][1]).toEqual(expect.objectContaining({
      requestQueryText: 'animals'
    }));

    const updateArgs = state.updateBlock.mock.calls[0][1] as {
      data: { queryVariants?: Array<{ kind: string; query: string; language?: string }> };
    };
    expect(updateArgs.data.queryVariants).toEqual([
      { kind: 'semantic-rewrite', query: 'animal-related documents', language: 'en' }
    ]);
  });

  it('skips the SharePoint primary branch when the planner has no safe lexical query', async () => {
    plannerSpy.mockResolvedValue({
      rawQuery: 'i am searching for documents about spfx',
      queryLanguage: 'en',
      usedCorrection: false,
      usedTranslation: false
    });

    searchSpy.mockResolvedValue({
      success: true,
      data: [makeResult('1', 'copilot-search')]
    });
    retrievalSpy.mockResolvedValue({
      success: true,
      data: [makeResult('2', 'copilot-retrieval')]
    });

    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();

    await handlers.search_sharepoint({ query: 'i am searching for documents about spfx' }, createDeps(store));

    expect(sharePointSpy).not.toHaveBeenCalled();
  });
});

describe('ToolRegistryRuntimeHandlersSearch.search_emails', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('routes email search through the shared adapter', async () => {
    const findExistingSession = jest.fn();
    const connectToM365Server = jest.fn();
    const handlers = buildSearchRuntimeHandlers({
      findExistingSession,
      connectToM365Server
    });
    const store = createStore({
      proxyConfig: {
        proxyUrl: 'https://proxy.example.com/api',
        proxyApiKey: 'test-key',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      mcpEnvironmentId: 'env-123'
    });

    await handlers.search_emails({ query: 'animals' }, createDeps(store));

    expect(findExistingSession).toHaveBeenCalledTimes(1);
    expect(connectToM365Server).toHaveBeenCalledTimes(1);
  });
});

describe('ToolRegistryRuntimeHandlersSearch.search_people', () => {
  const getStateMock = useGrimoireStore.getState as jest.Mock;

  let state: IStoreStateForTests;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.restoreAllMocks();
    state = {
      updateBlock: jest.fn(),
      pushBlock: jest.fn(),
      setPublicWebSearchCapability: jest.fn()
    };
    getStateMock.mockImplementation(() => state);
  });

  it('enriches specific people results with profile details from MCP', async () => {
    const executeCatalogMcpToolSpy = jest
      .spyOn(McpExecutionAdapter, 'executeCatalogMcpTool')
      .mockImplementation(async (options) => ({
        success: true,
        serverId: options.serverId,
        serverName: options.serverName,
        serverUrl: options.serverUrl,
        sessionId: 'session-profile',
        realToolName: options.toolName,
        requiredFields: ['userIdentifier'],
        schemaProps: {},
        normalizedArgs: options.rawArgs,
        resolvedArgs: options.rawArgs,
        targetSource: 'explicit-user',
        recoverySteps: [],
        mcpResult: {
          success: true,
          content: [
            {
              type: 'text',
              text: `User profile retrieved successfully.\n${JSON.stringify({
                displayName: options.rawArgs.userIdentifier === 'user@contoso.onmicrosoft.com' ? 'Test User' : 'Test User',
                mail: options.rawArgs.userIdentifier,
                jobTitle: 'Solution Architect',
                department: 'Digital Workplace',
                officeLocation: 'Zurich',
                businessPhones: ['+1 555 000 0000']
              })}; CorrelationId: abc, TimeStamp: 2026-03-10_19:32:15`
            }
          ]
        },
        trace: {
          toolName: options.toolName,
          rawArgs: options.rawArgs,
          targetSource: 'explicit-user',
          recoverySteps: []
        }
      } as never));

    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore({
      proxyConfig: {
        proxyUrl: 'https://proxy.example.com/api',
        proxyApiKey: 'test-key',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      mcpEnvironmentId: 'env-123'
    });
    getStateMock.mockImplementation(() => store);
    const peopleService = {
      searchPeople: jest.fn().mockResolvedValue({
        success: true,
        data: [
          {
            displayName: 'Test User',
            email: 'user@contoso.com'
          },
          {
            displayName: 'Test User',
            email: 'user@contoso.onmicrosoft.com'
          }
        ]
      })
    };

    const result = await handlers.search_people(
      { query: 'Test User' },
      { ...createDeps(store), peopleService: peopleService as never }
    );

    expect(result.phase).toBe('complete');
    expect(peopleService.searchPeople).toHaveBeenCalledWith('Test User', 5);
    expect(executeCatalogMcpToolSpy).toHaveBeenCalledTimes(2);
    expect(store.pushBlock).toHaveBeenCalledTimes(2);

    const firstCard = (store.pushBlock as jest.Mock).mock.calls[0][0];
    expect(firstCard.type).toBe('user-card');
    expect(firstCard.data).toEqual(expect.objectContaining({
      displayName: 'Test User',
      email: 'user@contoso.com',
      jobTitle: 'Solution Architect',
      department: 'Digital Workplace',
      officeLocation: 'Zurich',
      phone: '+1 555 000 0000'
    }));
  });

  it('asks for a name or email instead of searching with a wildcard people query', async () => {
    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();
    const peopleService = {
      searchPeople: jest.fn()
    };

    const result = await handlers.search_people(
      { query: '*' },
      { ...createDeps(store), peopleService: peopleService as never }
    );

    expect(result.phase).toBe('complete');
    expect(peopleService.searchPeople).not.toHaveBeenCalled();
    expect(store.pushBlock).toHaveBeenCalledTimes(1);

    const promptBlock = (store.pushBlock as jest.Mock).mock.calls[0][0];
    expect(promptBlock.type).toBe('info-card');
    expect(promptBlock.data).toEqual(expect.objectContaining({
      heading: 'People Search',
      body: expect.stringContaining("person's name or email")
    }));
  });

  it('normalizes natural-language people-card prompts before searching', async () => {
    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();
    const peopleService = {
      searchPeople: jest.fn().mockResolvedValue({
        success: true,
        data: []
      })
    };

    const result = await handlers.search_people(
      { query: 'show me Test User people card' },
      { ...createDeps(store), peopleService: peopleService as never }
    );

    expect(result.phase).toBe('complete');
    expect(peopleService.searchPeople).toHaveBeenCalledWith('Test User', 5);

    const payload = JSON.parse(result.output);
    expect(payload.query).toBe('Test User');

    const noResultBlock = state.pushBlock.mock.calls[0][0];
    expect(noResultBlock.type).toBe('info-card');
    expect(noResultBlock.title).toBe('People: Test User');
  });
});

describe('ToolRegistryRuntimeHandlersSearch.research_public_web', () => {
  const getStateMock = useGrimoireStore.getState as jest.Mock;

  let state: IStoreStateForTests;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.restoreAllMocks();
    state = {
      updateBlock: jest.fn(),
      pushBlock: jest.fn(),
      setPublicWebSearchCapability: jest.fn()
    };
    getStateMock.mockImplementation(() => state);
  });

  it('returns a deterministic disabled error when public web search is turned off', async () => {
    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore({ publicWebSearchEnabled: false });

    const result = await handlers.research_public_web({
      query: 'check this url',
      target_url: 'https://en.wikipedia.org/wiki/Microsoft'
    }, createDeps(store));

    expect(result.phase).toBe('error');
    expect(JSON.parse(result.output)).toMatchObject({
      success: false,
      error: 'Public web search is disabled in settings.'
    });
    expect(state.pushBlock).toHaveBeenCalledTimes(1);
  });

  it('renders a markdown block with citations when public web search succeeds', async () => {
    jest.spyOn(PublicWebSearchService.prototype, 'research').mockResolvedValue({
      success: true,
      text: 'Microsoft develops software and cloud services.',
      references: [
        { title: 'Microsoft - Wikipedia', url: 'https://en.wikipedia.org/wiki/Microsoft' },
        { title: 'Microsoft', url: 'https://www.microsoft.com/' }
      ],
      capabilityStatus: 'available',
      durationMs: 42
    });

    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore({
      proxyConfig: {
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test-key',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2025-01-01-preview'
      }
    });

    const result = await handlers.research_public_web({
      query: 'check this url for me',
      target_url: 'https://en.wikipedia.org/wiki/Microsoft'
    }, createDeps(store));

    expect(result.phase).toBe('complete');
    expect(state.pushBlock).toHaveBeenCalledTimes(1);

    const block = state.pushBlock.mock.calls[0][0];
    expect(block.type).toBe('markdown');
    expect(block.originTool).toBe('research_public_web');
    expect(block.data.content).toContain('Microsoft develops software');
    expect(block.data.content).toContain('[Microsoft - Wikipedia](https://en.wikipedia.org/wiki/Microsoft)');
    expect(state.setPublicWebSearchCapability).toHaveBeenCalledWith('available');
    expect(JSON.parse(result.output)).toMatchObject({
      success: true,
      referenceCount: 2
    });
  });
});

describe('ToolRegistryRuntimeHandlersSearch.search_sites', () => {
  const getStateMock = useGrimoireStore.getState as jest.Mock;

  let state: IStoreStateForTests;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.restoreAllMocks();
    state = {
      updateBlock: jest.fn(),
      pushBlock: jest.fn(),
      setPublicWebSearchCapability: jest.fn()
    };
    getStateMock.mockImplementation(() => state);
  });

  it('keeps reported site counts aligned with the rendered selection list', async () => {
    const handlers = buildSearchRuntimeHandlers({
      findExistingSession: jest.fn(),
      connectToM365Server: jest.fn()
    });
    const store = createStore();
    const sitesService = {
      searchSites: jest.fn().mockResolvedValue({
        success: true,
        data: Array.from({ length: 11 }, (_, index) => ({
          siteName: `Site ${index < 10 ? index + 1 : 1}`,
          siteUrl: `https://tenant.sharepoint.com/sites/${index < 10 ? `site-${index + 1}` : 'site-1'}`,
          description: `Description ${index + 1}`,
          libraries: [],
          lists: []
        }))
      })
    };

    const result = await handlers.search_sites(
      { query: 'copilot-test', max_results: '10' },
      { ...createDeps(store), sitesService: sitesService as never }
    );

    expect(result.phase).toBe('complete');
    expect(sitesService.searchSites).toHaveBeenCalledWith('copilot-test');

    const payload = JSON.parse(result.output);
    expect(payload.count).toBe(10);
    expect(payload.sites).toHaveLength(10);

    const selectionBlock = state.pushBlock.mock.calls[0][0];
    expect(selectionBlock.type).toBe('selection-list');
    expect(selectionBlock.data.items).toHaveLength(10);
  });
});
