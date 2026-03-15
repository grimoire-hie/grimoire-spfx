import {
  ensureCatalogMcpConnection,
  executeCatalogMcpTool,
  extractHiddenMcpError,
  extractStructuredMcpPayload,
  resolveEffectiveMcpTargetContext
} from './McpExecutionAdapter';

describe('McpExecutionAdapter', () => {
  it('prefers explicit target context over HIE and current-page fallbacks', () => {
    const resolved = resolveEffectiveMcpTargetContext({
      explicitTargetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      },
      sourceContext: {
        selectedItems: [
          {
            title: 'copilot-test',
            targetContext: {
              siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
              siteName: 'copilot-test'
            }
          }
        ]
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/current-site'
    });

    expect(resolved.targetSource).toBe('explicit-user');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      siteName: 'copilot-test-cooking',
      source: 'explicit-user'
    }));
  });

  it('falls back to HIE-derived target context when no explicit target exists', () => {
    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: {
        selectedItems: [
          {
            title: 'Test User',
            targetContext: {
              personEmail: 'user@contoso.com',
              personDisplayName: 'Test User'
            }
          }
        ]
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/current-site'
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      personEmail: 'user@contoso.com',
      personDisplayName: 'Test User'
    }));
  });

  it('unwraps nested response payloads consistently', () => {
    const result = extractStructuredMcpPayload([
      {
        type: 'text',
        text: JSON.stringify({
          response: JSON.stringify({
            payload: {
              value: [
                { id: '1', displayName: 'Documents' }
              ]
            }
          })
        })
      }
    ]);

    expect(result.unwrapPath).toEqual(['response', 'payload']);
    expect(result.payload).toEqual({
      value: [
        { id: '1', displayName: 'Documents' }
      ]
    });
  });

  it('detects hidden MCP errors inside text success wrappers', () => {
    expect(extractHiddenMcpError([
      {
        type: 'text',
        text: 'Missing required parameters: siteId (string: The full site ID is required.)'
      }
    ])).toContain('Missing required parameters: siteId');
  });

  it('falls back to catalog tool metadata when live session metadata is missing', async () => {
    const execute = jest.fn().mockResolvedValue({
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
    });

    const result = await executeCatalogMcpTool({
      serverId: 'mcp_SharePointListsTools',
      serverName: 'SharePoint Lists',
      serverUrl: 'https://agent365.example.com/mcp/lists',
      toolName: 'getSiteByPath',
      rawArgs: {
        hostname: 'contoso.sharepoint.com',
        serverRelativePath: 'sites/copilot-test-cooking'
      },
      connections: [],
      getConnections: () => [],
      mcpClient: { execute } as never,
      sessionHelpers: {
        findExistingSession: () => undefined,
        connectToM365Server: jest.fn().mockResolvedValue('session-lists')
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(result.success).toBe(true);
    expect(result.realToolName).toBe('getSiteByPath');
    expect(result.recoverySteps).toContain('used catalog session metadata fallback');
    expect(result.recoverySteps).toContain('used catalog tool schema for getSiteByPath');
    expect(execute).toHaveBeenCalledWith('session-lists', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/copilot-test-cooking'
    });
  });

  it('normalizes legacy createList args and template aliases into the SharePoint Lists schema', async () => {
    const execute = jest.fn()
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

    const result = await executeCatalogMcpTool({
      serverId: 'mcp_SharePointListsTools',
      serverName: 'SharePoint Lists',
      serverUrl: 'https://agent365.example.com/mcp/lists',
      toolName: 'createList',
      rawArgs: {
        siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        listName: 'Project Tracking',
        template: 'Custom',
        description: 'Project tracking list created by Test via Grimoire'
      },
      connections: [],
      getConnections: () => [],
      mcpClient: { execute } as never,
      sessionHelpers: {
        findExistingSession: () => undefined,
        connectToM365Server: jest.fn().mockResolvedValue('session-lists')
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(result.success).toBe(true);
    expect(result.resolvedArgs).toEqual({
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      displayName: 'Project Tracking',
      list: {
        template: 'genericList',
        description: 'Project tracking list created by Test via Grimoire'
      }
    });
    expect(execute).toHaveBeenNthCalledWith(1, 'session-lists', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/ProjectNova'
    });
    expect(execute).toHaveBeenNthCalledWith(2, 'session-lists', 'createList', {
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      displayName: 'Project Tracking',
      list: {
        template: 'genericList',
        description: 'Project tracking list created by Test via Grimoire'
      }
    });
  });

  it('normalizes snake_case createList args before resolving siteId', async () => {
    const execute = jest.fn()
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

    const result = await executeCatalogMcpTool({
      serverId: 'mcp_SharePointListsTools',
      serverName: 'SharePoint Lists',
      serverUrl: 'https://agent365.example.com/mcp/lists',
      toolName: 'createList',
      rawArgs: {
        site_url: 'https://contoso.sharepoint.com/sites/ProjectNova',
        list_name: 'Project Tracking',
        template: 'Custom'
      },
      connections: [],
      getConnections: () => [],
      mcpClient: { execute } as never,
      sessionHelpers: {
        findExistingSession: () => undefined,
        connectToM365Server: jest.fn().mockResolvedValue('session-lists')
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(result.success).toBe(true);
    expect(result.resolvedArgs).toEqual({
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      displayName: 'Project Tracking',
      list: {
        template: 'genericList'
      }
    });
    expect(execute).toHaveBeenNthCalledWith(1, 'session-lists', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/ProjectNova'
    });
    expect(execute).toHaveBeenNthCalledWith(2, 'session-lists', 'createList', {
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      displayName: 'Project Tracking',
      list: {
        template: 'genericList'
      }
    });
  });

  it('maps createListColumn intent args into schema args and resolves the target listId', async () => {
    const execute = jest.fn()
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
              id: 'column-mynotes',
              name: 'mynotes',
              displayName: 'mynotes'
            })
          }
        ]
      });

    const result = await executeCatalogMcpTool({
      serverId: 'mcp_SharePointListsTools',
      serverName: 'SharePoint Lists',
      serverUrl: 'https://agent365.example.com/mcp/lists',
      toolName: 'createListColumn',
      rawArgs: {
        siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        listName: 'testnellov1',
        columnDisplayName: 'mynotes',
        columnType: 'Text'
      },
      connections: [],
      getConnections: () => [],
      mcpClient: { execute } as never,
      sessionHelpers: {
        findExistingSession: () => undefined,
        connectToM365Server: jest.fn().mockResolvedValue('session-lists')
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova'
    });

    expect(result.success).toBe(true);
    expect(result.resolvedArgs).toEqual({
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      listId: 'list-testnellov1',
      name: 'mynotes',
      displayName: 'mynotes',
      text: {}
    });
    expect(execute).toHaveBeenNthCalledWith(1, 'session-lists', 'getSiteByPath', {
      hostname: 'contoso.sharepoint.com',
      serverRelativePath: 'sites/ProjectNova'
    });
    expect(execute).toHaveBeenNthCalledWith(2, 'session-lists', 'listLists', {
      siteId: 'contoso.sharepoint.com,site-project,web-project'
    });
    expect(execute).toHaveBeenNthCalledWith(3, 'session-lists', 'createListColumn', {
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      listId: 'list-testnellov1',
      name: 'mynotes',
      displayName: 'mynotes',
      text: {}
    });
  });

  it('maps choice column form args into schema args with normalized choices', async () => {
    const execute = jest.fn()
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
              id: 'column-status',
              name: 'status',
              displayName: 'Status'
            })
          }
        ]
      });

    const result = await executeCatalogMcpTool({
      serverId: 'mcp_SharePointListsTools',
      serverName: 'SharePoint Lists',
      serverUrl: 'https://agent365.example.com/mcp/lists',
      toolName: 'createListColumn',
      rawArgs: {
        siteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        listName: 'testnellov1',
        columnDisplayName: 'Status',
        columnType: 'Choice',
        choiceValues: 'New\nIn Progress\nDone'
      },
      connections: [],
      getConnections: () => [],
      mcpClient: { execute } as never,
      sessionHelpers: {
        findExistingSession: () => undefined,
        connectToM365Server: jest.fn().mockResolvedValue('session-lists')
      },
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova'
    });

    expect(result.success).toBe(true);
    expect(result.resolvedArgs).toEqual({
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      listId: 'list-testnellov1',
      name: 'status',
      displayName: 'Status',
      choice: {
        choices: ['New', 'In Progress', 'Done']
      }
    });
    expect(execute).toHaveBeenNthCalledWith(3, 'session-lists', 'createListColumn', {
      siteId: 'contoso.sharepoint.com,site-project,web-project',
      listId: 'list-testnellov1',
      name: 'status',
      displayName: 'Status',
      choice: {
        choices: ['New', 'In Progress', 'Done']
      }
    });
  });

  it('reuses the shared connection helper to fall back to catalog session metadata', async () => {
    const connectToM365Server = jest.fn().mockResolvedValue('session-mail');

    const result = await ensureCatalogMcpConnection({
      serverId: 'mcp_MailTools',
      serverName: 'Outlook Mail',
      serverUrl: 'https://agent365.example.com/mcp/mail',
      connections: [],
      getConnections: () => [],
      mcpClient: {} as never,
      sessionHelpers: {
        findExistingSession: () => undefined,
        connectToM365Server
      }
    });

    expect(result.success).toBe(true);
    expect(result.sessionId).toBe('session-mail');
    expect(result.connection?.serverName).toBe('Outlook Mail');
    expect(result.recoverySteps).toContain('used catalog session metadata fallback');
    expect(connectToM365Server).toHaveBeenCalled();
  });
});
