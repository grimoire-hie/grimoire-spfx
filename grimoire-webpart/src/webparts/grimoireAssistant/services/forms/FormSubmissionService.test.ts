const executeMock = jest.fn();
const getStateMock = jest.fn();

jest.mock('../mcp/McpClientService', () => ({
  McpClientService: jest.fn().mockImplementation(() => ({
    execute: executeMock
  }))
}));

jest.mock('../tools/ToolRuntimeSharedHelpers', () => ({
  findExistingSession: jest.fn(() => 'session-odsp'),
  connectToM365Server: jest.fn(async () => 'session-odsp')
}));

jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    emitEvent: jest.fn(),
    onBlockCreated: jest.fn(),
    getCurrentTaskContext: jest.fn(() => undefined),
    getCurrentArtifacts: jest.fn(() => ({}))
  }
}));

jest.mock('../../store/useGrimoireStore', () => ({
  useGrimoireStore: {
    getState: (...args: unknown[]) => getStateMock(...args)
  }
}));

import type { IFormData } from '../../models/IBlock';
import type { IMcpConnection } from '../../models/IMcpTypes';
import type { IFunctionCallStore } from '../tools/ToolRuntimeContracts';
import { executeFormSubmission } from './FormSubmissionService';

function createStore(): IFunctionCallStore {
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
    setActivityStatus: jest.fn()
  };
}

function createConnection(): IMcpConnection {
  return {
    sessionId: 'session-odsp',
    serverUrl: 'https://agent365.svc.cloud.microsoft/mcp/environments/env-123/servers/mcp_ODSPRemoteServer',
    serverName: 'OneDrive and SharePoint Files',
    state: 'connected',
    connectedAt: new Date('2026-03-09T10:00:00.000Z'),
    tools: [
      {
        name: 'createFolder',
        description: 'Create a folder.',
        inputSchema: {
          type: 'object',
          properties: {
            folderName: { type: 'string' },
            documentLibraryId: { type: 'string' }
          },
          required: ['folderName', 'documentLibraryId']
        }
      }
    ]
  };
}

function createMailConnection(): IMcpConnection {
  return {
    sessionId: 'session-odsp',
    serverUrl: 'https://agent365.svc.cloud.microsoft/mcp/environments/env-123/servers/mcp_MailTools',
    serverName: 'Mail',
    state: 'connected',
    connectedAt: new Date('2026-03-09T10:00:00.000Z'),
    tools: [
      {
        name: 'SendEmailWithAttachments',
        description: 'Send an email with attachments.',
        inputSchema: {
          type: 'object',
          properties: {
            to: { type: 'array' },
            cc: { type: 'array' },
            bcc: { type: 'array' },
            subject: { type: 'string' },
            body: { type: 'string' },
            attachmentUris: { type: 'array' }
          },
          required: ['to', 'subject', 'body']
        }
      }
    ]
  };
}

function createTeamsConnection(): IMcpConnection {
  return {
    sessionId: 'session-odsp',
    serverUrl: 'https://agent365.svc.cloud.microsoft/mcp/environments/env-123/servers/mcp_TeamsServer',
    serverName: 'Teams',
    state: 'connected',
    connectedAt: new Date('2026-03-09T10:00:00.000Z'),
    tools: [
      {
        name: 'CreateChat',
        description: 'Create a new Teams chat.',
        inputSchema: {
          type: 'object',
          properties: {
            recipients: { type: 'array' },
            topic: { type: 'string' }
          },
          required: ['recipients']
        }
      }
    ]
  };
}

function createListsConnection(): IMcpConnection {
  return {
    sessionId: 'session-odsp',
    serverUrl: 'https://agent365.svc.cloud.microsoft/mcp/environments/env-123/servers/mcp_SharePointListsTools',
    serverName: 'SharePoint Lists',
    state: 'connected',
    connectedAt: new Date('2026-03-09T10:00:00.000Z'),
    tools: [
      {
        name: 'getSiteByPath',
        description: 'Resolve a site by path.',
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
        name: 'createList',
        description: 'Create a SharePoint list.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' },
            displayName: { type: 'string' },
            list: { type: 'object' }
          },
          required: ['siteId', 'displayName', 'list']
        }
      },
      {
        name: 'listLists',
        description: 'List SharePoint lists.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' }
          },
          required: ['siteId']
        }
      },
      {
        name: 'createListColumn',
        description: 'Create a SharePoint list column.',
        inputSchema: {
          type: 'object',
          properties: {
            siteId: { type: 'string' },
            listId: { type: 'string' },
            name: { type: 'string' },
            displayName: { type: 'string' },
            text: { type: 'object' },
            choice: { type: 'object' }
          },
          required: ['siteId', 'listId', 'name']
        }
      }
    ]
  };
}

describe('FormSubmissionService target context resolution', () => {
  beforeEach(() => {
    executeMock.mockReset();
    getStateMock.mockReset().mockReturnValue({
      mcpConnections: [createConnection()],
      removeMcpConnection: jest.fn(),
      getToken: jest.fn()
    });
  });

  it('creates folders in the explicitly targeted site instead of the current site', async () => {
    executeMock
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'contoso.sharepoint.com,site-cooking,web-cooking',
              displayName: 'copilot-test-cooking'
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'drive-cooking',
              name: 'Freigegebene Dokumente'
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              name: 'test whatsoever',
              webUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/test%20whatsoever'
            })
          }
        ]
      });

    const formData: IFormData = {
      kind: 'form',
      preset: 'folder-create',
      description: 'Create folder in cooking site',
      fields: [
        { key: 'folderName', label: 'Folder name', type: 'text', required: true }
      ],
      submissionTarget: {
        toolName: 'createFolder',
        serverId: 'mcp_ODSPRemoteServer',
        staticArgs: {},
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
          siteName: 'copilot-test-cooking'
        }
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      { folderName: 'test whatsoever' },
      {},
      createStore()
    );

    expect(result.success).toBe(true);
    expect(executeMock).toHaveBeenNthCalledWith(1, 'session-odsp', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/copilot-test-cooking'
    });
    expect(executeMock).toHaveBeenNthCalledWith(2, 'session-odsp', 'getDefaultDocumentLibraryInSite', {
      siteId: 'contoso.sharepoint.com,site-cooking,web-cooking'
    });
    expect(executeMock).toHaveBeenNthCalledWith(3, 'session-odsp', 'createFolder', {
      folderName: 'test whatsoever',
      documentLibraryId: 'drive-cooking'
    });
    expect(result.message).toContain('copilot-test-cooking');
  });

  it('creates real docx files in a named document library for word document forms', async () => {
    executeMock
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'contoso.sharepoint.com,site-nova,web-nova',
              displayName: 'ProjectNova'
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              value: [
                {
                  id: 'drive-shared-docs',
                  name: 'Shared Documents'
                },
                {
                  id: 'drive-mydoclib',
                  name: 'mydoclib'
                }
              ]
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              name: 'Recap - Nova launch.docx',
              webUrl: 'https://contoso.sharepoint.com/sites/ProjectNova/mydoclib/Recap%20-%20Nova%20launch.docx'
            })
          }
        ]
      });

    const formData: IFormData = {
      kind: 'form',
      preset: 'word-document-create',
      description: 'Create a Word document from the visible recap',
      fields: [
        { key: 'filename', label: 'Document name', type: 'text', required: true },
        { key: 'contentText', label: 'Content', type: 'textarea', required: true }
      ],
      submissionTarget: {
        toolName: 'createSmallBinaryFile',
        serverId: 'mcp_ODSPRemoteServer',
        staticArgs: {},
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
          siteName: 'ProjectNova',
          documentLibraryName: 'mydoclib'
        }
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      {
        filename: 'Recap - Nova launch.docx',
        contentText: 'Nova launch recap summary.'
      },
      {},
      createStore()
    );

    expect(result.success).toBe(true);
    expect(executeMock).toHaveBeenNthCalledWith(1, 'session-odsp', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/ProjectNova'
    });
    expect(executeMock).toHaveBeenNthCalledWith(2, 'session-odsp', 'listDocumentLibrariesInSite', {
      siteId: 'contoso.sharepoint.com,site-nova,web-nova'
    });

    const createCallArgs = executeMock.mock.calls[2][2] as {
      filename: string;
      documentLibraryId: string;
      base64Content: string;
    };

    expect(createCallArgs.filename).toBe('Recap - Nova launch.docx');
    expect(createCallArgs.documentLibraryId).toBe('drive-mydoclib');
    expect(createCallArgs.base64Content.startsWith('UEsDB')).toBe(true);

    const docxPayload = Buffer.from(createCallArgs.base64Content, 'base64').toString('utf8');
    expect(docxPayload).toContain('[Content_Types].xml');
    expect(docxPayload).toContain('Nova launch recap summary.');
  });

  it('sanitizes invalid SharePoint filename characters before creating a word document', async () => {
    getStateMock.mockReturnValue({
      mcpConnections: [createConnection()],
      removeMcpConnection: jest.fn(),
      getToken: jest.fn()
    });

    executeMock
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'contoso.sharepoint.com,site-nova,web-nova'
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              value: [
                {
                  driveId: 'drive-mydoclib',
                  documentLibraryName: 'mydoclib'
                }
              ]
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              name: 'Recap - Search launch recap.docx'
            })
          }
        ]
      });

    const formData: IFormData = {
      kind: 'form',
      preset: 'word-document-create',
      description: 'Create a Word document from the visible recap',
      fields: [
        { key: 'filename', label: 'Document name', type: 'text', required: true },
        { key: 'contentText', label: 'Content', type: 'textarea', required: true }
      ],
      submissionTarget: {
        toolName: 'createSmallBinaryFile',
        serverId: 'mcp_ODSPRemoteServer',
        staticArgs: {},
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
          siteName: 'ProjectNova',
          documentLibraryName: 'mydoclib'
        }
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      {
        filename: 'Recap - Search: launch recap.docx',
        contentText: 'Nova launch recap summary.'
      },
      {},
      createStore()
    );

    expect(result.success).toBe(true);

    const createCallArgs = executeMock.mock.calls[2][2] as {
      filename: string;
      documentLibraryId: string;
      base64Content: string;
    };

    expect(createCallArgs.filename).toBe('Recap - Search launch recap.docx');
    expect(createCallArgs.documentLibraryId).toBe('drive-mydoclib');
    expect(createCallArgs.base64Content.startsWith('UEsDB')).toBe(true);
  });

  it('sends only the explicitly selected attachment for narrowed email shares', async () => {
    getStateMock.mockReturnValue({
      mcpConnections: [createMailConnection()],
      removeMcpConnection: jest.fn(),
      getToken: jest.fn()
    });

    executeMock.mockResolvedValueOnce({
      success: true,
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            message: 'Email sent successfully.'
          })
        }
      ]
    });

    const formData: IFormData = {
      kind: 'form',
      preset: 'email-compose',
      description: 'Send the selected file by email',
      fields: [
        { key: 'to', label: 'To', type: 'email-list', required: true },
        { key: 'subject', label: 'Subject', type: 'text', required: true },
        { key: 'body', label: 'Body', type: 'textarea', required: true }
      ],
      submissionTarget: {
        toolName: 'SendEmailWithAttachments',
        serverId: 'mcp_MailTools',
        staticArgs: {
          attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf']
        },
        targetContext: {
          siteUrl: 'https://tenant.sharepoint.com/sites/dev',
          fileOrFolderUrl: 'https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf',
          fileOrFolderName: 'Power Platform.pdf',
          source: 'explicit-user'
        }
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      {
        to: '',
        subject: 'Power Platform',
        body: 'Please review the attached document.'
      },
      {
        to: ['recipient@example.com']
      },
      createStore()
    );

    expect(result.success).toBe(true);
    expect(executeMock).toHaveBeenCalledWith('session-odsp', 'SendEmailWithAttachments', {
      to: ['recipient@example.com'],
      subject: 'Power Platform',
      body: 'Please review the attached document.',
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/Power%20Platform.pdf']
    });
  });

  it('maps people-picker recipients to array args for submission tools', async () => {
    getStateMock.mockReturnValue({
      mcpConnections: [createTeamsConnection()],
      removeMcpConnection: jest.fn(),
      getToken: jest.fn()
    });

    executeMock.mockResolvedValueOnce({
      success: true,
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            message: 'Email sent successfully.'
          })
        }
      ]
    });

    const formData: IFormData = {
      kind: 'form',
      preset: 'share-teams-chat',
      description: 'Share to Teams chat',
      fields: [
        { key: 'recipients', label: 'People', type: 'people-picker', required: true },
        { key: 'topic', label: 'Chat topic', type: 'text', required: false }
      ],
      submissionTarget: {
        toolName: 'CreateChat',
        serverId: 'mcp_TeamsServer',
        staticArgs: {}
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      {
        recipients: '',
        topic: 'Teams recap'
      },
      {
        recipients: ['test.user@example.com', 'teammate@example.com']
      },
      createStore()
    );

    expect(result.success).toBe(true);
    expect(executeMock).toHaveBeenCalledWith('session-odsp', 'CreateChat', {
      recipients: ['test.user@example.com', 'teammate@example.com'],
      topic: 'Teams recap'
    });
  });

  it('refreshes SharePoint list siteId from target context and preserves nested field mappings', async () => {
    getStateMock.mockReturnValue({
      mcpConnections: [createListsConnection()],
      removeMcpConnection: jest.fn(),
      getToken: jest.fn()
    });

    executeMock
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'contoso.sharepoint.com,site-project,web-project',
              displayName: 'Project Nova'
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'list-project-tracking',
              displayName: 'Project Tracking',
              webUrl: 'https://contoso.sharepoint.com/sites/ProjectNova/Lists/Project%20Tracking'
            })
          }
        ]
      });

    const formData: IFormData = {
      kind: 'form',
      preset: 'generic',
      description: 'Create a new SharePoint list',
      fields: [
        { key: 'displayName', label: 'List name', type: 'text', required: true },
        { key: 'description', label: 'Description', type: 'textarea', required: false }
      ],
      submissionTarget: {
        toolName: 'createList',
        serverId: 'mcp_SharePointListsTools',
        staticArgs: {
          siteId: 'contoso.sharepoint.com,stale-site,stale-web',
          list: {
            template: 'genericList'
          }
        },
        fieldToParamMap: {
          displayName: 'displayName',
          description: 'list.description'
        },
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
          siteName: 'Project Nova',
          source: 'explicit-user'
        }
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      {
        displayName: 'Project Tracking',
        description: 'Project tracking list created by Test via Grimoire'
      },
      {},
      createStore()
    );

    expect(result.success).toBe(true);
    expect(executeMock).toHaveBeenNthCalledWith(1, 'session-odsp', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/ProjectNova'
    });
    expect(executeMock).toHaveBeenNthCalledWith(2, 'session-odsp', 'createList', {
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      displayName: 'Project Tracking',
      list: {
        template: 'genericList',
        description: 'Project tracking list created by Test via Grimoire'
      }
    });
  });

  it('renders a friendly SharePoint column result block after generic form submission', async () => {
    getStateMock.mockReturnValue({
      mcpConnections: [createListsConnection()],
      removeMcpConnection: jest.fn(),
      getToken: jest.fn()
    });

    executeMock
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              id: 'contoso.sharepoint.com,site-project,web-project',
              displayName: 'Project Nova'
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              value: [
                {
                  id: 'list-testnellov1',
                  displayName: 'testnellov1',
                  webUrl: 'https://contoso.sharepoint.com/sites/ProjectNova/Lists/testnellov1',
                  list: {
                    template: 'genericList'
                  }
                }
              ]
            })
          }
        ]
      })
      .mockResolvedValueOnce({
        success: true,
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              message: 'Graph tool executed successfully.',
              response: JSON.stringify({
                '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#sites(\'contoso.sharepoint.com%2Csite-project%2Cweb-project\')/lists(\'list-testnellov1\')/columns/$entity',
                id: 'column-mynotes',
                columnGroup: 'Custom Columns',
                displayName: 'mynotes',
                name: 'mynotes',
                required: false,
                enforceUniqueValues: false,
                hidden: false,
                indexed: false,
                readOnly: false,
                text: {
                  allowMultipleLines: false
                }
              })
            })
          }
        ]
      });

    const store = createStore();
    const formData: IFormData = {
      kind: 'form',
      preset: 'generic',
      description: 'Create a SharePoint column',
      fields: [
        { key: 'columnType', label: 'Column type', type: 'dropdown', required: true, options: [{ key: 'Text', text: 'Text' }] },
        { key: 'displayName', label: 'Display name', type: 'text', required: true },
        { key: 'columnName', label: 'Internal name', type: 'text', required: false }
      ],
      submissionTarget: {
        toolName: 'createListColumn',
        serverId: 'mcp_SharePointListsTools',
        staticArgs: {
          siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
          listName: 'testnellov1'
        },
        fieldToParamMap: {
          columnType: 'columnType',
          displayName: 'displayName',
          columnName: 'columnName'
        },
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
          listName: 'testnellov1',
          source: 'explicit-user'
        }
      },
      status: 'editing'
    };

    const result = await executeFormSubmission(
      formData,
      {
        columnType: 'Text',
        displayName: 'mynotes',
        columnName: ''
      },
      {},
      store
    );

    expect(result.success).toBe(true);
    expect(result.message).toBe('Added Text column "mynotes" to list "testnellov1".');
    expect(executeMock).toHaveBeenNthCalledWith(1, 'session-odsp', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/ProjectNova'
    });
    expect(executeMock).toHaveBeenNthCalledWith(2, 'session-odsp', 'listLists', {
      siteId: 'contoso.sharepoint.com,site-project,web-project'
    });
    expect(executeMock).toHaveBeenNthCalledWith(3, 'session-odsp', 'createListColumn', {
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      listId: 'list-testnellov1',
      name: 'mynotes',
      displayName: 'mynotes',
      text: {}
    });
    expect(store.pushBlock).toHaveBeenCalledTimes(1);
    expect((store.pushBlock as jest.Mock).mock.calls[0][0]).toMatchObject({
      type: 'info-card',
      title: 'mynotes',
      data: expect.objectContaining({
        heading: 'mynotes',
        body: expect.stringContaining('Type: Text')
      })
    });
  });
});
