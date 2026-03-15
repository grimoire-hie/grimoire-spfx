const executeMock = jest.fn();
const withMcpRetryMock = jest.fn();
const mapMcpResultToBlockMock = jest.fn();
const getStateMock = jest.fn();

jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('../../store/useGrimoireStore', () => ({
  useGrimoireStore: {
    getState: (...args: unknown[]) => getStateMock(...args)
  }
}));

jest.mock('../mcp/McpClientService', () => ({
  McpClientService: jest.fn().mockImplementation(() => ({
    execute: executeMock
  }))
}));

jest.mock('../mcp/McpResultMapper', () => ({
  mapMcpResultToBlock: (...args: unknown[]) => mapMcpResultToBlockMock(...args)
}));

jest.mock('../mcp/mcpUtils', () => {
  const actual = jest.requireActual('../mcp/mcpUtils');
  return {
    ...actual,
    withMcpRetry: (...args: unknown[]) => withMcpRetryMock(...args)
  };
});

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    getCurrentTaskContext: jest.fn(() => undefined),
    getCurrentArtifacts: jest.fn(() => ({})),
    onBlockCreated: jest.fn(),
    onToolComplete: jest.fn(),
    onBlockRemoved: jest.fn(),
    sendToolError: jest.fn()
  }
}));

import type { IMcpConnection } from '../../models/IMcpTypes';
import type { IInfoCardData } from '../../models/IBlock';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { logService } from '../logging/LogService';
import { applySchemaArgAliases, buildMcpRuntimeHandlers } from './ToolRegistryRuntimeHandlersMcp';

function createStore(overrides: Partial<IFunctionCallStore> = {}): IFunctionCallStore {
  return {
    aadHttpClient: undefined,
    proxyConfig: {
      proxyUrl: 'https://proxy.example.com/api',
      proxyApiKey: 'test-key',
      backend: 'reasoning',
      deployment: 'grimoire-reasoning',
      apiVersion: '2024-10-21'
    },
    getToken: jest.fn(),
    mcpEnvironmentId: 'env-123',
    userContext: {
      displayName: 'Test User',
      email: 'test.user@example.com',
      loginName: 'test.user@example.com',
      resolvedLanguage: 'en',
      currentWebTitle: 'copilot-test',
      currentWebUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
      currentSiteTitle: 'copilot-test',
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    },
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

function createDeps(store: IFunctionCallStore, sourceContext?: IToolRuntimeHandlerDeps['sourceContext']): IToolRuntimeHandlerDeps {
  return {
    store,
    awaitAsync: true,
    aadClient: undefined,
    sitesService: undefined,
    peopleService: undefined,
    sourceContext
  };
}

function createConnection(serverId: string, serverName: string, tools: IMcpConnection['tools']): IMcpConnection {
  return {
    sessionId: 'session-1',
    serverUrl: `https://agent365.svc.cloud.microsoft/mcp/environments/env-123/servers/${serverId}`,
    serverName,
    tools,
    state: 'connected',
    connectedAt: new Date('2026-03-09T10:00:00.000Z')
  };
}

describe('ToolRegistryRuntimeHandlersMcp argument alias normalization', () => {
  beforeEach(() => {
    executeMock.mockReset();
    withMcpRetryMock.mockReset();
    mapMcpResultToBlockMock.mockReset().mockReturnValue('mapped');
    getStateMock.mockReset();
  });

  it('maps userEmail to userIdentifier for profile tools and defaults a rich select projection', () => {
    const normalized = applySchemaArgAliases(
      'GetUserDetails',
      {},
      { userEmail: 'user@contoso.com' },
      {
        userIdentifier: { type: 'string' },
        select: { type: 'string' }
      }
    );

    expect(normalized.userIdentifier).toBe('user@contoso.com');
    expect(normalized.select).toBe('displayName,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone,givenName,surname,id');
  });

  it('maps sitePath plus current site URL into getSiteByPath schema arguments', () => {
    const normalized = applySchemaArgAliases(
      'getSiteByPath',
      {},
      { sitePath: '/sites/copilot-test' },
      {
        hostname: { type: 'string' },
        serverRelativePath: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test'
    );

    expect(normalized.hostname).toBe('contoso.sharepoint.com');
    expect(normalized.serverRelativePath).toBe('sites/copilot-test');
  });

  it('maps snake_case site_path into getSiteByPath schema arguments', () => {
    const normalized = applySchemaArgAliases(
      'getSiteByPath',
      {},
      { site_path: '/sites/copilot-test' },
      {
        hostname: { type: 'string' },
        serverRelativePath: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test'
    );

    expect(normalized.hostname).toBe('contoso.sharepoint.com');
    expect(normalized.serverRelativePath).toBe('sites/copilot-test');
  });

  it('maps HIE-selected person context into userIdentifier', () => {
    const normalized = applySchemaArgAliases(
      'GetUserDetails',
      {},
      {},
      {
        userIdentifier: { type: 'string' }
      },
      undefined,
      {
        personEmail: 'user@contoso.com',
        personDisplayName: 'Test User'
      }
    );

    expect(normalized.userIdentifier).toBe('user@contoso.com');
  });

  it('uses current site context for getSiteByPath when no explicit site args are provided', () => {
    const normalized = applySchemaArgAliases(
      'getSiteByPath',
      {},
      {},
      {
        hostname: { type: 'string' },
        serverRelativePath: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test'
    );

    expect(normalized.hostname).toBe('contoso.sharepoint.com');
    expect(normalized.serverRelativePath).toBe('sites/copilot-test');
  });

  it('prefers explicit target site context over current page fallback', () => {
    const normalized = applySchemaArgAliases(
      'getSiteByPath',
      {},
      {},
      {
        hostname: { type: 'string' },
        serverRelativePath: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test',
      {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      }
    );

    expect(normalized.hostname).toBe('contoso.sharepoint.com');
    expect(normalized.serverRelativePath).toBe('sites/copilot-test-cooking');
  });
});

describe('ToolRegistryRuntimeHandlersMcp use_m365_capability conformance', () => {
  beforeEach(() => {
    executeMock.mockReset();
    withMcpRetryMock.mockReset();
    mapMcpResultToBlockMock.mockReset().mockReturnValue('mapped');
    getStateMock.mockReset();
    (hybridInteractionEngine.getCurrentTaskContext as jest.Mock).mockReset().mockReturnValue(undefined);
    (hybridInteractionEngine.getCurrentArtifacts as jest.Mock).mockReset().mockReturnValue({});
  });

  it('resolves selected HIE site context before executing listLists', async () => {
    const connection = createConnection('mcp_SharePointListsTools', 'SharePoint Lists', [
      {
        name: 'getSiteByPath',
        description: 'Resolve a site by hostname and server-relative path.',
        inputSchema: {
          type: 'object',
          properties: {
            hostname: { type: 'string' },
            serverRelativePath: { type: 'string' }
          },
          required: ['hostname', 'serverRelativePath']
        }
      },
      {
        name: 'listLists',
        description: 'List lists for a site.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' }
          },
          required: ['siteId']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [connection] });
    getStateMock.mockReturnValue(store);
    executeMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ id: 'contoso.sharepoint.com,site-cooking,web-cooking' }) }]
    });
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ value: [{ id: 'documents', displayName: 'Documents' }] }) }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'listLists',
      arguments_json: '{}',
      server_hint: 'mcp_SharePointListsTools'
    }, createDeps(store, {
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      }
    }));

    expect(executeMock).toHaveBeenCalledWith('session-1', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/copilot-test-cooking'
    });
    expect(logService.debug).toHaveBeenCalledWith(
      'mcp',
      'MCP execution trace',
      expect.stringContaining('"siteId":"contoso.sharepoint.com,site-cooking,web-cooking"')
    );
    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
  });

  it('resolves explicit site-name args before executing listLists', async () => {
    const connection = createConnection('mcp_SharePointListsTools', 'SharePoint Lists', [
      {
        name: 'searchSitesByName',
        description: 'Search sites by display name.',
        inputSchema: {
          type: 'object',
          properties: {
            search: { type: 'string' }
          },
          required: ['search']
        }
      },
      {
        name: 'listLists',
        description: 'List lists for a site.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' }
          },
          required: ['siteId']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [connection] });
    getStateMock.mockReturnValue(store);
    executeMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ value: [{ id: 'contoso.sharepoint.com,site-cooking,web-cooking' }] }) }]
    });
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ value: [{ id: 'documents', displayName: 'Documents' }] }) }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'listLists',
      arguments_json: JSON.stringify({ siteName: 'copilot-test-cooking' }),
      server_hint: 'mcp_SharePointListsTools'
    }, createDeps(store));

    expect(executeMock).toHaveBeenCalledWith('session-1', 'searchSitesByName', {
      search: 'copilot-test-cooking'
    });
    expect(logService.debug).toHaveBeenCalledWith(
      'mcp',
      'MCP execution trace',
      expect.stringContaining('"siteId":"contoso.sharepoint.com,site-cooking,web-cooking"')
    );
    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
  });

  it('resolves the latest site artifact before executing listLists when no explicit args are provided', async () => {
    const connection = createConnection('mcp_SharePointListsTools', 'SharePoint Lists', [
      {
        name: 'getSiteByPath',
        description: 'Resolve a site by hostname and server-relative path.',
        inputSchema: {
          type: 'object',
          properties: {
            hostname: { type: 'string' },
            serverRelativePath: { type: 'string' }
          },
          required: ['hostname', 'serverRelativePath']
        }
      },
      {
        name: 'listLists',
        description: 'List lists for a site.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' }
          },
          required: ['siteId']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [connection] });
    getStateMock.mockReturnValue(store);
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
    executeMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ id: 'contoso.sharepoint.com,site-cooking,web-cooking' }) }]
    });
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ value: [{ id: 'documents', displayName: 'Documents' }] }) }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'listLists',
      arguments_json: '{}',
      server_hint: 'mcp_SharePointListsTools'
    }, createDeps(store));

    expect(executeMock).toHaveBeenCalledWith('session-1', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/copilot-test-cooking'
    });
    expect(logService.debug).toHaveBeenCalledWith(
      'mcp',
      'MCP execution trace',
      expect.stringContaining('"siteId":"contoso.sharepoint.com,site-cooking,web-cooking"')
    );
    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
  });

  it('uses the selected person identity for GetUserDetails instead of defaulting to me', async () => {
    const connection = createConnection('mcp_MeServer', 'User Profile', [
      {
        name: 'GetUserDetails',
        description: 'Get user details.',
        inputSchema: {
          type: 'object',
          properties: {
            userIdentifier: { type: 'string' },
            select: { type: 'string' }
          },
          required: ['userIdentifier']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [connection] });
    getStateMock.mockReturnValue(store);
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{ type: 'text', text: JSON.stringify({ displayName: 'Test User', mail: 'user@contoso.com' }) }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'GetUserDetails',
      arguments_json: JSON.stringify({ userEmail: 'user@contoso.com' }),
      server_hint: 'mcp_MeServer'
    }, createDeps(store));

    expect(logService.debug).toHaveBeenCalledWith(
      'mcp',
      'MCP execution trace',
      expect.stringContaining('"userIdentifier":"user@contoso.com"')
    );
    expect(logService.debug).not.toHaveBeenCalledWith(
      'mcp',
      'MCP execution trace',
      expect.stringContaining('"userIdentifier":"me"')
    );
    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
  });

  it('treats duplicate createList conflicts as an acknowledgment instead of an error block', async () => {
    const connection = createConnection('mcp_SharePointListsTools', 'SharePoint Lists', [
      {
        name: 'createList',
        description: 'Create a new SharePoint list on a site.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' },
            displayName: { type: 'string' },
            list: { type: 'object' }
          },
          required: ['siteId', 'displayName', 'list']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [connection] });
    getStateMock.mockReturnValue(store);
    withMcpRetryMock.mockResolvedValue({
      success: false,
      error: 'Error executing tool: Graph API call failed for tool \'createList\': Failed to call graph API at https://graph.microsoft.com/v1.0/sites/tenant/lists. Status=409 (Conflict) Error=Name already exists.'
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'createList',
      arguments_json: JSON.stringify({
        siteId: 'contoso.sharepoint.com,site-project,web-project',
        displayName: 'Project Tracking',
        list: { template: 'genericList' }
      }),
      server_hint: 'mcp_SharePointListsTools'
    }, createDeps(store, {
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        siteName: 'Project Nova'
      }
    }));

    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({
      success: true,
      alreadyExists: true
    }));

    const pushedBlocks = (store.pushBlock as jest.Mock).mock.calls.map((call) => call[0]);
    const infoBlock = pushedBlocks[pushedBlocks.length - 1];
    expect(infoBlock.type).toBe('info-card');
    expect(infoBlock.title).toBe('Project Tracking');
    expect((infoBlock.data as IInfoCardData).body).toContain('already exists in Project Nova');
    expect(hybridInteractionEngine.sendToolError).not.toHaveBeenCalled();
  });

  it('bootstraps personal OneDrive root browsing through ODSP MCP without Graph fallback', async () => {
    const sharePointConnection = createConnection('mcp_SharePointListsTools', 'SharePoint Lists', [
      {
        name: 'getSiteByPath',
        description: 'Resolve a site by hostname and server-relative path.',
        inputSchema: {
          type: 'object',
          properties: {
            hostname: { type: 'string' },
            serverRelativePath: { type: 'string' }
          },
          required: ['hostname', 'serverRelativePath']
        }
      }
    ]);
    const odspConnection = createConnection('mcp_ODSPRemoteServer', 'ODSP', [
      {
        name: 'getDefaultDocumentLibraryInSite',
        description: 'Get the default document library for a site.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' }
          },
          required: ['siteId']
        }
      },
      {
        name: 'getFolderChildren',
        description: 'List folder children.',
        inputSchema: {
          type: 'object',
          properties: {
            documentLibraryId: { type: 'string' },
            fileOrFolderId: { type: 'string' }
          },
          required: ['documentLibraryId']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [sharePointConnection, odspConnection] });
    getStateMock.mockReturnValue(store);
    withMcpRetryMock
      .mockResolvedValueOnce({
        success: true,
        content: [{
          type: 'text',
          text: JSON.stringify({ id: 'contoso-my.sharepoint.com,personal-testuser,web-testuser' })
        }]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [{
          type: 'text',
          text: JSON.stringify({ id: 'drive-personal-root' })
        }]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [{
          type: 'text',
          text: JSON.stringify({
            value: [{
              id: 'file-1',
              name: 'Nova Plan.docx',
              webUrl: 'https://contoso-my.sharepoint.com/personal/test_user_example_com/Documents/Nova%20Plan.docx',
              parentReference: {
                driveId: 'drive-personal-root'
              }
            }]
          })
        }]
      });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'getFolderChildren',
      arguments_json: JSON.stringify({ personalOneDrive: true }),
      server_hint: 'mcp_ODSPRemoteServer'
    }, createDeps(store));

    expect(withMcpRetryMock).toHaveBeenCalledTimes(3);
    expect(withMcpRetryMock).toHaveBeenNthCalledWith(
      1,
      expect.any(Object),
      'session-1',
      'getSiteByPath',
      {
        hostname: 'contoso-my.sharepoint.com',
        serverRelativePath: 'personal/test_user_example_com'
      },
      expect.stringContaining('/mcp/environments/env-123/servers/mcp_SharePointListsTools'),
      'SharePoint Lists'
    );
    expect(withMcpRetryMock).toHaveBeenNthCalledWith(
      2,
      expect.any(Object),
      'session-1',
      'getDefaultDocumentLibraryInSite',
      {
        siteId: 'contoso-my.sharepoint.com,personal-testuser,web-testuser'
      },
      expect.stringContaining('/mcp/environments/env-123/servers/mcp_ODSPRemoteServer'),
      'SharePoint & OneDrive'
    );
    expect(withMcpRetryMock).toHaveBeenNthCalledWith(
      3,
      expect.any(Object),
      'session-1',
      'getFolderChildren',
      {
        documentLibraryId: 'drive-personal-root'
      },
      expect.stringContaining('/mcp/environments/env-123/servers/mcp_ODSPRemoteServer'),
      'SharePoint & OneDrive'
    );
    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
  });

  it('reuses block-carried ODSP ids for personal OneDrive follow-up folder browsing', async () => {
    const odspConnection = createConnection('mcp_ODSPRemoteServer', 'ODSP', [
      {
        name: 'getFolderChildren',
        description: 'List folder children.',
        inputSchema: {
          type: 'object',
          properties: {
            documentLibraryId: { type: 'string' },
            fileOrFolderId: { type: 'string' }
          },
          required: ['documentLibraryId']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [odspConnection] });
    getStateMock.mockReturnValue(store);
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{
        type: 'text',
        text: JSON.stringify({
          value: [{
            id: 'child-1',
            name: 'Nested Folder',
            webUrl: 'https://contoso-my.sharepoint.com/personal/test_user_example_com/Documents/Nested%20Folder',
            parentReference: {
              driveId: 'drive-existing'
            }
          }]
        })
      }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'getFolderChildren',
      arguments_json: JSON.stringify({
        personalOneDrive: true,
        documentLibraryId: 'drive-existing',
        fileOrFolderId: 'folder-123'
      }),
      server_hint: 'mcp_ODSPRemoteServer'
    }, createDeps(store));

    expect(withMcpRetryMock).toHaveBeenCalledTimes(1);
    expect(withMcpRetryMock).toHaveBeenCalledWith(
      expect.any(Object),
      'session-1',
      'getFolderChildren',
      {
        documentLibraryId: 'drive-existing',
        fileOrFolderId: 'folder-123'
      },
      expect.stringContaining('/mcp/environments/env-123/servers/mcp_ODSPRemoteServer'),
      'SharePoint & OneDrive'
    );
    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
  });

  it('fails honestly when personal OneDrive bootstrap cannot derive an email-like identity', async () => {
    const sharePointConnection = createConnection('mcp_SharePointListsTools', 'SharePoint Lists', [
      {
        name: 'getSiteByPath',
        description: 'Resolve a site by hostname and server-relative path.',
        inputSchema: {
          type: 'object',
          properties: {
            hostname: { type: 'string' },
            serverRelativePath: { type: 'string' }
          },
          required: ['hostname', 'serverRelativePath']
        }
      }
    ]);
    const odspConnection = createConnection('mcp_ODSPRemoteServer', 'ODSP', [
      {
        name: 'getDefaultDocumentLibraryInSite',
        description: 'Get the default document library for a site.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' }
          },
          required: ['siteId']
        }
      },
      {
        name: 'getFolderChildren',
        description: 'List folder children.',
        inputSchema: {
          type: 'object',
          properties: {
            documentLibraryId: { type: 'string' }
          },
          required: ['documentLibraryId']
        }
      }
    ]);
    const baseUserContext = createStore().userContext!;
    const store = createStore({
      mcpConnections: [sharePointConnection, odspConnection],
      userContext: {
        ...baseUserContext,
        email: '',
        loginName: 'not-an-email'
      }
    });
    getStateMock.mockReturnValue(store);

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'getFolderChildren',
      arguments_json: JSON.stringify({ personalOneDrive: true }),
      server_hint: 'mcp_ODSPRemoteServer'
    }, createDeps(store));

    expect(withMcpRetryMock).not.toHaveBeenCalled();
    expect(result.phase).toBe('error');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({
      success: false,
      error: 'I could not derive your personal OneDrive site URL for MCP browsing.'
    }));
  });

  it('filters personal OneDrive ODSP search results to the derived personal site only', async () => {
    const odspConnection = createConnection('mcp_ODSPRemoteServer', 'ODSP', [
      {
        name: 'findFileOrFolder',
        description: 'Find files or folders.',
        inputSchema: {
          type: 'object',
          properties: {
            searchQuery: { type: 'string' }
          },
          required: ['searchQuery']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [odspConnection] });
    getStateMock.mockReturnValue(store);
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{
        type: 'text',
        text: JSON.stringify({
          value: [
            {
              id: 'personal-1',
              name: 'Nova Plan.docx',
              webUrl: 'https://contoso-my.sharepoint.com/personal/test_user_example_com/Documents/Nova%20Plan.docx'
            },
            {
              id: 'site-1',
              name: 'Site Plan.docx',
              webUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Shared%20Documents/Site%20Plan.docx'
            }
          ]
        })
      }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'findFileOrFolder',
      arguments_json: JSON.stringify({
        searchQuery: 'nova',
        personalOneDrive: true
      }),
      server_hint: 'mcp_ODSPRemoteServer'
    }, createDeps(store));

    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({ success: true, summary: 'mapped' }));
    expect(mapMcpResultToBlockMock).toHaveBeenCalledTimes(1);
    const scopedContent = mapMcpResultToBlockMock.mock.calls[0][2] as Array<{ type: string; text?: string }>;
    const scopedPayload = JSON.parse(String(scopedContent[0].text || '{}')) as { value: Array<{ id: string; name: string; webUrl: string }> };
    expect(scopedPayload.value).toEqual([
      {
        id: 'personal-1',
        name: 'Nova Plan.docx',
        webUrl: 'https://contoso-my.sharepoint.com/personal/test_user_example_com/Documents/Nova%20Plan.docx'
      }
    ]);
  });

  it('shows an honest limitation when personal OneDrive ODSP search results cannot be safely scoped', async () => {
    const odspConnection = createConnection('mcp_ODSPRemoteServer', 'ODSP', [
      {
        name: 'findFileOrFolder',
        description: 'Find files or folders.',
        inputSchema: {
          type: 'object',
          properties: {
            searchQuery: { type: 'string' }
          },
          required: ['searchQuery']
        }
      }
    ]);
    const store = createStore({ mcpConnections: [odspConnection] });
    getStateMock.mockReturnValue(store);
    withMcpRetryMock.mockResolvedValue({
      success: true,
      content: [{
        type: 'text',
        text: JSON.stringify({
          value: [{
            id: 'personal-1',
            name: 'Nova Plan.docx'
          }]
        })
      }]
    });

    const handlers = buildMcpRuntimeHandlers({
      findExistingSession: jest.fn(() => 'session-1'),
      connectToM365Server: jest.fn(async () => 'session-1')
    });

    const result = await handlers.use_m365_capability({
      tool_name: 'findFileOrFolder',
      arguments_json: JSON.stringify({
        searchQuery: 'nova',
        personalOneDrive: true
      }),
      server_hint: 'mcp_ODSPRemoteServer'
    }, createDeps(store));

    expect(result.phase).toBe('complete');
    expect(JSON.parse(result.output)).toEqual(expect.objectContaining({
      success: true,
      unsupported: true
    }));
    expect(mapMcpResultToBlockMock).not.toHaveBeenCalled();
    const pushedBlocks = (store.pushBlock as jest.Mock).mock.calls.map((call) => call[0]);
    const infoBlock = pushedBlocks[pushedBlocks.length - 1];
    expect(infoBlock.type).toBe('info-card');
    expect((infoBlock.data as IInfoCardData).body).toContain('verifiable URL');
  });
});
