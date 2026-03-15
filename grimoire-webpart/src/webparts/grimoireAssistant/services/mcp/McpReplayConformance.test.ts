import { createBlock } from '../../models/IBlock';
import type { IFormData } from '../../models/IBlock';
import { HybridInteractionEngine } from '../hie/HybridInteractionEngine';
import {
  applySchemaArgAliases,
  resolveEffectiveMcpTargetContext
} from './McpExecutionAdapter';

describe('MCP replay conformance', () => {
  it('replays selected site -> site lookup using the selected site instead of current page fallback', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-site-select',
      blockType: 'selection-list',
      action: 'select',
      timestamp: Date.now(),
      payload: {
        label: 'copilot-test-cooking',
        prompt: 'Found 3 sites',
        selectedItems: [
          {
            id: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
            label: 'copilot-test-cooking',
            description: 'https://contoso.sharepoint.com/sites/copilot-test-cooking'
          }
        ]
      }
    });

    const sourceContext = engine.captureCurrentSourceContext();
    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext,
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(resolved.targetSource).toBe('hie-selection');

    const normalized = applySchemaArgAliases(
      'getSiteByPath',
      {},
      {},
      {
        hostname: { type: 'string' },
        serverRelativePath: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test',
      resolved.targetContext
    );

    expect(normalized.hostname).toBe('contoso.sharepoint.com');
    expect(normalized.serverRelativePath).toBe('sites/copilot-test-cooking');
  });

  it('replays clicked person -> profile lookup using the selected person identity', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-user-card',
      blockType: 'user-card',
      action: 'click-user',
      timestamp: Date.now(),
      payload: {
        displayName: 'Test User',
        email: 'user@contoso.com'
      }
    });

    const sourceContext = engine.captureCurrentSourceContext();
    const resolved = resolveEffectiveMcpTargetContext({ sourceContext });

    expect(resolved.targetContext).toEqual(expect.objectContaining({
      personEmail: 'user@contoso.com',
      personDisplayName: 'Test User'
    }));

    const normalized = applySchemaArgAliases(
      'GetUserDetails',
      {},
      {},
      {
        userIdentifier: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );

    expect(normalized.userIdentifier).toBe('user@contoso.com');
  });

  it('replays explicit form target -> folder creation keeping the form target over current page context', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    const formBlock = createBlock('form', 'Create folder in Documents', {
      kind: 'form',
      preset: 'folder-create',
      fields: [],
      status: 'editing',
      submissionTarget: {
        toolName: 'createFolder',
        serverId: 'mcp_ODSPRemoteServer',
        staticArgs: {},
        targetContext: {
          siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
          siteName: 'copilot-test-cooking'
        }
      }
    } as IFormData);

    engine.onBlockCreated(formBlock);

    const sourceContext = engine.captureCurrentSourceContext();
    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext,
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      siteName: 'copilot-test-cooking'
    }));
  });

  it('replays site-info follow-up -> folder creation using the latest site artifact instead of current page fallback', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    const siteBlock = createBlock('site-info', 'copilot-test-cooking', {
      kind: 'site-info',
      siteName: 'copilot-test-cooking',
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      libraries: ['Dokumente']
    });
    engine.onBlockCreated(siteBlock);

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts(),
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      siteName: 'copilot-test-cooking'
    }));
  });

  it('replays selected search result -> share follow-up preserving the exact selected file target', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-search',
      blockType: 'search-results',
      action: 'click-result',
      timestamp: Date.now(),
      payload: {
        title: 'SPFx.pdf',
        url: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/SPFx.pdf',
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: spfx'
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts(),
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/SPFx.pdf',
      fileOrFolderName: 'SPFx.pdf'
    }));

    const shareArgs = applySchemaArgAliases(
      'shareFileOrFolder',
      {},
      {},
      {
        fileOrFolderUrl: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test',
      resolved.targetContext
    );

    expect(shareArgs.fileOrFolderUrl).toBe(
      'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/SPFx.pdf'
    );
  });

  it('replays clicked list row -> list follow-up preserving site and list identifiers', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-lists',
      blockType: 'list-items',
      action: 'click-list-row',
      timestamp: Date.now(),
      payload: {
        index: 1,
        rowData: {
          displayName: 'Events',
          id: '00000000-0000-0000-0000-000000000003',
          webUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Lists/Events',
          parentReference: JSON.stringify({
            siteId: 'contoso.sharepoint.com,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002'
          }),
          list: JSON.stringify({
            template: 'events'
          })
        }
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts(),
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking'
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
      siteId: 'contoso.sharepoint.com,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002',
      listId: '00000000-0000-0000-0000-000000000003',
      listUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Lists/Events',
      listName: 'Events'
    }));

    const normalized = applySchemaArgAliases(
      'listItems',
      {},
      {},
      {
        siteId: { type: 'string' },
        listId: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      resolved.targetContext
    );

    expect(normalized.siteId).toBe('contoso.sharepoint.com,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002');
    expect(normalized.listId).toBe('00000000-0000-0000-0000-000000000003');
  });

  it('replays clicked document-library file -> url-based follow-up preserving canonical file targets', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-library',
      blockType: 'document-library',
      action: 'click-file',
      timestamp: Date.now(),
      payload: {
        index: 3,
        name: 'Brotkultur_Deutschsprachige_Laender_DE.docx',
        title: 'Brotkultur_Deutschsprachige_Laender_DE.docx',
        url: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/Brotkultur_Deutschsprachige_Laender_DE.docx',
        fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/Brotkultur_Deutschsprachige_Laender_DE.docx',
        fileOrFolderId: 'drive-item-brotkultur',
        documentLibraryId: 'drive-cooking-docs',
        documentLibraryUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente',
        documentLibraryName: 'Freigegebene Dokumente',
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking',
        type: 'file',
        fileType: 'docx'
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts(),
      currentSiteUrl: 'https://contoso.sharepoint.com/sites/copilot-test'
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      documentLibraryId: 'drive-cooking-docs',
      documentLibraryUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente',
      documentLibraryName: 'Freigegebene Dokumente',
      fileOrFolderId: 'drive-item-brotkultur',
      fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/Brotkultur_Deutschsprachige_Laender_DE.docx',
      fileOrFolderName: 'Brotkultur_Deutschsprachige_Laender_DE.docx'
    }));

    const normalized = applySchemaArgAliases(
      'getFileOrFolderMetadataByUrl',
      {},
      {},
      {
        fileOrFolderUrl: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test',
      resolved.targetContext
    );

    expect(normalized.fileOrFolderUrl).toBe(
      'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/Brotkultur_Deutschsprachige_Laender_DE.docx'
    );

    const shareArgs = applySchemaArgAliases(
      'shareFileOrFolder',
      {},
      {},
      {
        fileOrFolderUrl: { type: 'string' }
      },
      'https://contoso.sharepoint.com/sites/copilot-test',
      resolved.targetContext
    );

    expect(shareArgs.fileOrFolderUrl).toBe(
      'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/Brotkultur_Deutschsprachige_Laender_DE.docx'
    );
  });

  it('replays selected document-library file -> mutation follow-ups preserving file and library ids', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-library',
      blockType: 'document-library',
      action: 'click-file',
      timestamp: Date.now(),
      payload: {
        index: 2,
        name: 'SPFx.pdf',
        title: 'SPFx.pdf',
        url: 'https://contoso.sharepoint.com/sites/copilot-test/Freigegebene%20Dokumente/SPFx.pdf',
        fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Freigegebene%20Dokumente/SPFx.pdf',
        fileOrFolderId: 'drive-item-spfx',
        documentLibraryId: 'drive-copilot-docs',
        documentLibraryUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Freigegebene%20Dokumente',
        documentLibraryName: 'Freigegebene Dokumente',
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
        siteName: 'copilot-test',
        type: 'file',
        fileType: 'pdf'
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts()
    });

    expect(resolved.targetContext).toEqual(expect.objectContaining({
      documentLibraryId: 'drive-copilot-docs',
      fileOrFolderId: 'drive-item-spfx'
    }));

    const renameArgs = applySchemaArgAliases(
      'renameFileOrFolder',
      { newFileOrFolderName: 'SPFx-renamed.pdf' },
      {},
      {
        documentLibraryId: { type: 'string' },
        fileOrFolderId: { type: 'string' },
        newFileOrFolderName: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(renameArgs).toEqual({
      documentLibraryId: 'drive-copilot-docs',
      fileOrFolderId: 'drive-item-spfx',
      newFileOrFolderName: 'SPFx-renamed.pdf'
    });

    const deleteArgs = applySchemaArgAliases(
      'deleteFileOrFolder',
      {},
      {},
      {
        documentLibraryId: { type: 'string' },
        fileOrFolderId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(deleteArgs).toEqual({
      documentLibraryId: 'drive-copilot-docs',
      fileOrFolderId: 'drive-item-spfx'
    });

    const moveArgs = applySchemaArgAliases(
      'moveSmallFile',
      { newParentFolderId: 'folder-target-1' },
      {},
      {
        documentLibraryId: { type: 'string' },
        fileId: { type: 'string' },
        newParentFolderId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(moveArgs).toEqual({
      documentLibraryId: 'drive-copilot-docs',
      fileId: 'drive-item-spfx',
      newParentFolderId: 'folder-target-1'
    });
  });

  it('replays selected document-library folder -> container follow-ups preserving parent folder id', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-library',
      blockType: 'document-library',
      action: 'click-folder',
      timestamp: Date.now(),
      payload: {
        index: 1,
        name: 'chedice',
        title: 'chedice',
        url: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/chedice',
        fileOrFolderUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente/chedice',
        fileOrFolderId: 'folder-chedice',
        documentLibraryId: 'drive-cooking-docs',
        documentLibraryUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking/Freigegebene%20Dokumente',
        documentLibraryName: 'Freigegebene Dokumente',
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking',
        type: 'folder'
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts()
    });

    const createFolderArgs = applySchemaArgAliases(
      'createFolder',
      { folderName: 'nested-folder' },
      {},
      {
        folderName: { type: 'string' },
        documentLibraryId: { type: 'string' },
        parentFolderId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(createFolderArgs).toEqual({
      folderName: 'nested-folder',
      documentLibraryId: 'drive-cooking-docs',
      parentFolderId: 'folder-chedice'
    });

    const listChildrenArgs = applySchemaArgAliases(
      'getFolderChildren',
      {},
      {},
      {
        documentLibraryId: { type: 'string' },
        parentFolderId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(listChildrenArgs).toEqual({
      documentLibraryId: 'drive-cooking-docs',
      parentFolderId: 'folder-chedice'
    });
  });

  it('replays clicked list item row -> update and delete follow-ups preserving site, list, and item ids', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-list-items',
      blockType: 'list-items',
      action: 'click-list-row',
      timestamp: Date.now(),
      payload: {
        index: 2,
        rowData: {
          Title: 'Release checklist',
          siteId: 'contoso.sharepoint.com,site-guid,web-guid',
          listId: 'list-guid-checklists',
          listName: 'Checklists',
          listUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Lists/Checklists',
          itemId: '42'
        }
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts()
    });

    expect(resolved.targetContext).toEqual(expect.objectContaining({
      siteId: 'contoso.sharepoint.com,site-guid,web-guid',
      listId: 'list-guid-checklists',
      listItemId: '42',
      listName: 'Checklists'
    }));

    const updateArgs = applySchemaArgAliases(
      'updateListItem',
      { fields: { Status: 'Done' } },
      {},
      {
        siteId: { type: 'string' },
        listId: { type: 'string' },
        itemId: { type: 'string' },
        fields: { type: 'object' }
      },
      undefined,
      resolved.targetContext
    );
    expect(updateArgs).toEqual({
      siteId: 'contoso.sharepoint.com,site-guid,web-guid',
      listId: 'list-guid-checklists',
      itemId: '42',
      fields: { Status: 'Done' }
    });

    const deleteArgs = applySchemaArgAliases(
      'deleteListItem',
      {},
      {},
      {
        siteId: { type: 'string' },
        listId: { type: 'string' },
        itemId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(deleteArgs).toEqual({
      siteId: 'contoso.sharepoint.com,site-guid,web-guid',
      listId: 'list-guid-checklists',
      itemId: '42'
    });
  });

  it('replays clicked mail row -> reply and forward follow-ups preserving the selected message id', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-mails',
      blockType: 'list-items',
      action: 'click-list-row',
      timestamp: Date.now(),
      payload: {
        index: 1,
        rowData: {
          subject: 'SPFx follow-up',
          messageId: 'msg-spfx-123',
          webUrl: 'https://outlook.office.com/mail/msg-spfx-123'
        }
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts()
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      mailItemId: 'msg-spfx-123'
    }));

    const replyArgs = applySchemaArgAliases(
      'ReplyToMessage',
      {},
      {},
      {
        id: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(replyArgs.id).toBe('msg-spfx-123');

    const forwardArgs = applySchemaArgAliases(
      'ForwardMessage',
      {},
      {},
      {
        messageId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );
    expect(forwardArgs.messageId).toBe('msg-spfx-123');
  });

  it('replays clicked calendar row -> event follow-up preserving the selected event id', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    engine.onBlockInteraction({
      blockId: 'block-events',
      blockType: 'list-items',
      action: 'click-list-row',
      timestamp: Date.now(),
      payload: {
        index: 1,
        rowData: {
          subject: 'SPFx demo',
          eventId: 'evt-spfx-456'
        }
      }
    });

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts()
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      calendarItemId: 'evt-spfx-456'
    }));

    const updateArgs = applySchemaArgAliases(
      'UpdateEvent',
      {},
      {},
      {
        eventId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );

    expect(updateArgs.eventId).toBe('evt-spfx-456');
  });

  it('replays explicit Teams channel destination -> channel post preserving team and channel ids', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });

    const formBlock = createBlock('form', 'Post to a Teams Channel', {
      kind: 'form',
      preset: 'share-teams-channel',
      fields: [],
      status: 'editing',
      submissionTarget: {
        toolName: 'share_teams_channel',
        serverId: 'internal-share',
        staticArgs: {},
        targetContext: {
          teamId: 'team-spfx',
          teamName: 'Engineering',
          channelId: 'channel-general',
          channelName: 'General'
        }
      }
    } as IFormData);

    engine.onBlockCreated(formBlock);

    const resolved = resolveEffectiveMcpTargetContext({
      sourceContext: engine.captureCurrentSourceContext(),
      taskContext: engine.getCurrentTaskContext(),
      artifacts: engine.getCurrentArtifacts()
    });

    expect(resolved.targetSource).toBe('hie-selection');
    expect(resolved.targetContext).toEqual(expect.objectContaining({
      teamId: 'team-spfx',
      teamName: 'Engineering',
      channelId: 'channel-general',
      channelName: 'General'
    }));

    const channelArgs = applySchemaArgAliases(
      'PostChannelMessage',
      {},
      {},
      {
        teamId: { type: 'string' },
        channelId: { type: 'string' }
      },
      undefined,
      resolved.targetContext
    );

    expect(channelArgs).toEqual({
      teamId: 'team-spfx',
      channelId: 'channel-general'
    });
  });
});
