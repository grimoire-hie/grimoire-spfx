import { createBlock } from '../../models/IBlock';
import { HybridInteractionEngine } from './HybridInteractionEngine';

describe('HybridInteractionEngine tool error feedback', () => {
  it('injects a response-triggering tool error when an error block is completed', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const errorBlock = createBlock('error', 'Email Search Error', {
      kind: 'error',
      message: 'MCP error -32001: Request timed out'
    });
    engine.onBlockCreated(errorBlock);
    engine.onToolComplete('search_emails', errorBlock.id, false, 0);

    const toolErrorCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool error:')
    );

    expect(toolErrorCall).toBeDefined();
    expect(toolErrorCall?.[0]).toContain('search_emails failed');
    expect(toolErrorCall?.[0]).toContain('MCP error -32001: Request timed out');
    expect(toolErrorCall?.[1]).toBe(true);
  });

  it('includes the selected permission target in injected tool error context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const errorBlock = createBlock('error', 'Permissions: Brotkultur_Deutschsprachige_Laender_DE.docx', {
      kind: 'error',
      message: 'I couldn\'t inspect permissions for "Brotkultur_Deutschsprachige_Laender_DE.docx" because the connected SharePoint & OneDrive MCP server does not expose a permission-inspection tool.'
    });
    engine.onBlockCreated(errorBlock);
    engine.onToolComplete('show_permissions', errorBlock.id, false, 0);

    const toolErrorCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool error:')
    );

    expect(toolErrorCall).toBeDefined();
    expect(toolErrorCall?.[0]).toContain('show_permissions failed');
    expect(toolErrorCall?.[0]).toContain('Brotkultur_Deutschsprachige_Laender_DE.docx');
  });

  it('does not auto-inject tool error context for non-error blocks', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const searchBlock = createBlock('search-results', 'Search', {
      kind: 'search-results',
      query: 'budget',
      results: [{ title: 'Q4 Budget', summary: 'Draft', url: 'https://contoso/share/q4-budget.docx' }],
      totalCount: 1,
      source: 'semantic'
    });
    engine.onBlockCreated(searchBlock);
    engine.onToolComplete('search_sharepoint', searchBlock.id, false, 0);

    const hasToolErrorCall = sendContextMessage.mock.calls.some(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool error:')
    );
    expect(hasToolErrorCall).toBe(false);
  });

  it('injects a generic tool error context when failure has no error block', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.onToolComplete('search_people', '', false, 0);

    const toolErrorCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool error:')
    );

    expect(toolErrorCall).toBeDefined();
    expect(toolErrorCall?.[0]).toContain('search_people failed');
    expect(toolErrorCall?.[1]).toBe(true);
  });

  it('respects context policy when tool error injection is disabled', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, {
      sendContextMessage,
      contextPolicy: {
        maxChars: 1500,
        debounceMs: 800,
        triggerRules: { toolError: false }
      }
    });

    engine.sendToolError('use_m365_capability', 'Request timed out');

    const hasToolErrorCall = sendContextMessage.mock.calls.some(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool error:')
    );
    expect(hasToolErrorCall).toBe(false);
  });
});

describe('HybridInteractionEngine async completion feedback', () => {
  beforeEach(() => {
    jest.useFakeTimers();
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  it('uses local callback and silent completion context for low-latency voice tools', () => {
    const sendContextMessage = jest.fn();
    const onAsyncToolCompletion = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage, onAsyncToolCompletion });
    engine.setVoicePathActive(true);

    const searchBlock = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'animals',
      results: [
        { title: 'Animal Facts', summary: 'Overview', url: 'https://contoso.sharepoint.com/doc1.docx' }
      ],
      totalCount: 1,
      source: 'semantic'
    });
    engine.onBlockCreated(searchBlock);
    engine.onToolComplete('search_sharepoint', searchBlock.id, true, 1);

    expect(onAsyncToolCompletion).toHaveBeenCalledWith(
      expect.objectContaining({
        toolName: 'search_sharepoint',
        blockId: searchBlock.id,
        itemCount: 1
      })
    );

    jest.advanceTimersByTime(300);

    const completionCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool completed:')
    );
    expect(completionCall).toBeDefined();
    expect(completionCall?.[1]).toBe(false);
  });

  it('falls back to model-triggered completion when no local callback is configured', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });
    engine.setVoicePathActive(true);

    const searchBlock = createBlock('search-results', 'Search: budget', {
      kind: 'search-results',
      query: 'budget',
      results: [
        { title: 'Q4 Budget', summary: 'Draft', url: 'https://contoso.sharepoint.com/q4-budget.docx' }
      ],
      totalCount: 1,
      source: 'semantic'
    });
    engine.onBlockCreated(searchBlock);
    engine.onToolComplete('search_sharepoint', searchBlock.id, true, 1);

    jest.advanceTimersByTime(300);

    const completionCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool completed:')
    );
    expect(completionCall).toBeDefined();
    expect(completionCall?.[1]).toBe(true);
  });

  it('uses local callback and silent completion context for compose forms in voice mode', () => {
    const sendContextMessage = jest.fn();
    const onAsyncToolCompletion = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage, onAsyncToolCompletion });
    engine.setVoicePathActive(true);

    const formBlock = createBlock('form', 'Share via Email', {
      kind: 'form',
      preset: 'email-compose',
      description: 'Share the current result',
      fields: [],
      submissionTarget: { toolName: 'SendEmailWithAttachments', serverId: 'mcp_MailTools', staticArgs: {} },
      status: 'editing'
    });
    engine.onBlockCreated(formBlock);
    engine.onToolComplete('show_compose_form', formBlock.id, true, 0);

    expect(onAsyncToolCompletion).toHaveBeenCalledWith(
      expect.objectContaining({
        toolName: 'show_compose_form',
        blockId: formBlock.id,
        itemCount: 0
      })
    );

    jest.advanceTimersByTime(300);

    const completionCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Tool completed:')
    );
    expect(completionCall).toBeDefined();
    expect(completionCall?.[1]).toBe(false);
  });
});

describe('HybridInteractionEngine derived HIE state', () => {
  it('replays queued pre-init shell events after initialization', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();

    engine.emitEvent({
      eventName: 'shell.logs.toggled',
      source: 'app-layout',
      surface: 'logs',
      correlationId: 'shell-pre-init',
      payload: { isOpen: true },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });

    expect(engine.getRecentEvents()).toHaveLength(0);

    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    expect(engine.getShellState().isLogsOpen).toBe(true);
    expect(engine.getRecentEvents()).toHaveLength(1);
    expect(sendContextMessage).not.toHaveBeenCalled();
  });

  it('stores shell events without sending them to the prompt', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.emitEvent({
      eventName: 'shell.logs.toggled',
      source: 'app-layout',
      surface: 'logs',
      correlationId: 'shell-1',
      payload: { isOpen: true },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });

    expect(engine.getShellState().isLogsOpen).toBe(true);
    expect(engine.getRecentEvents()).toHaveLength(1);
    expect(sendContextMessage).not.toHaveBeenCalled();
  });

  it('tracks recap artifacts as the current task context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.emitEvent({
      eventName: 'task.recap.requested',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'recap-1',
      payload: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx',
        derivedBlockId: 'block-recap',
        derivedBlockTitle: 'Recap: Search: SPFx'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-search',
      blockType: 'search-results'
    });

    engine.emitEvent({
      eventName: 'artifact.recap.ready',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'recap-1',
      payload: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx',
        derivedBlockId: 'block-recap',
        derivedBlockTitle: 'Recap: Search: SPFx'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-recap',
      blockType: 'info-card'
    });

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'recap',
      sourceBlockId: 'block-search',
      derivedBlockId: 'block-recap',
      derivedBlockTitle: 'Recap: Search: SPFx'
    });
    expect(engine.getCurrentArtifacts()['block-recap']).toMatchObject({
      artifactKind: 'recap',
      status: 'ready'
    });

    const taskContextCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].startsWith('[Task context:')
    );
    expect(taskContextCall).toBeDefined();
  });

  it('tracks visible summary cards as derived block artifacts', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const summaryBlock = createBlock('info-card', 'Summary: TeamsFx.pdf', {
      kind: 'info-card',
      heading: 'Summary: TeamsFx.pdf',
      body: 'TeamsFx summary content'
    });
    engine.onBlockCreated(summaryBlock);

    expect(engine.getCurrentArtifacts()[summaryBlock.id]).toMatchObject({
      artifactKind: 'summary',
      blockType: 'info-card',
      title: 'Summary: TeamsFx.pdf',
      status: 'ready'
    });
  });

  it('promotes top-level search results into the current task context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });
    engine.setCurrentTurnId('turn-spfx');

    const searchBlock = createBlock('search-results', 'Search: spfx', {
      kind: 'search-results',
      query: 'spfx',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx guidance',
          summary: 'Implementation guidance for SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
          sources: ['copilot-search']
        }
      ]
    });

    engine.onBlockCreated(searchBlock);

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'search',
      sourceBlockId: searchBlock.id,
      sourceBlockTitle: 'Search: spfx',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-spfx'
    });
  });

  it('does not inherit the previous task context when a new root turn creates a top-level search block', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.beginUserTurn({ turnId: 'turn-animals', mode: 'new-root' });
    const animalBlock = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'animals',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'Animal facts',
          summary: 'Animal facts.',
          url: 'https://tenant.sharepoint.com/sites/dev/animals.docx',
          sources: ['copilot-search']
        }
      ]
    });
    engine.onBlockCreated(animalBlock);

    engine.beginUserTurn({ turnId: 'turn-spfx', mode: 'new-root' });
    const spfxBlock = createBlock('search-results', 'Search: spfx', {
      kind: 'search-results',
      query: 'spfx',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx guidance',
          summary: 'Implementation guidance for SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
          sources: ['copilot-search']
        }
      ]
    });
    engine.onBlockCreated(spfxBlock);

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'search',
      sourceBlockId: spfxBlock.id,
      sourceBlockTitle: 'Search: spfx',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-spfx'
    });
    expect(engine.getBlockTracker().get(spfxBlock.id)).toMatchObject({
      turnId: 'turn-spfx',
      rootTurnId: 'turn-spfx'
    });
  });

  it('links derived summary blocks back to the producing interaction task context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });
    engine.setCurrentTurnId('turn-spfx');

    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'SPFx',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx guidance',
          summary: 'Implementation guidance for SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
          sources: ['copilot-search']
        }
      ]
    });
    engine.onBlockCreated(searchBlock);

    engine.onBlockInteraction({
      blockId: searchBlock.id,
      blockType: searchBlock.type,
      action: 'summarize',
      eventName: 'block.interaction.summarize',
      exposurePolicy: { mode: 'response-triggering', relevance: 'foreground' },
      source: 'hover-action',
      surface: 'action-panel',
      payload: {
        title: 'SPFx guidance',
        url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
        selectedItems: [{ index: 1, title: 'SPFx guidance' }]
      },
      timestamp: Date.now()
    });

    const summaryBlock = createBlock('info-card', 'Summary: SPFx guidance', {
      kind: 'info-card',
      heading: 'Summary: SPFx guidance',
      body: 'SPFx guidance summary content'
    });
    engine.onBlockCreated(summaryBlock);

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'summarize',
      sourceBlockId: searchBlock.id,
      sourceBlockTitle: 'Search: SPFx',
      derivedBlockId: summaryBlock.id,
      derivedBlockTitle: 'Summary: SPFx guidance'
    });
    expect(engine.getCurrentArtifacts()[summaryBlock.id]).toMatchObject({
      artifactKind: 'summary',
      sourceBlockId: searchBlock.id,
      sourceTaskKind: 'summarize',
      sourceTurnId: 'turn-spfx'
    });
  });

  it('tracks non-list result blocks as derived block artifacts', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const previewBlock = createBlock('file-preview', 'File: SPFx.pdf', {
      kind: 'file-preview',
      fileName: 'SPFx.pdf',
      fileUrl: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
      fileType: 'pdf'
    });
    engine.onBlockCreated(previewBlock);

    expect(engine.getCurrentArtifacts()[previewBlock.id]).toMatchObject({
      artifactKind: 'preview',
      blockType: 'file-preview',
      title: 'File: SPFx.pdf',
      status: 'ready'
    });
  });

  it('inherits source turn lineage onto derived form artifacts', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.setCurrentTurnId('turn-spfx');
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'SPFx',
      results: [
        { title: 'SPFx guidance', summary: 'Overview', url: 'https://contoso.sharepoint.com/spfx.docx' }
      ],
      totalCount: 1,
      source: 'semantic'
    });
    engine.onBlockCreated(searchBlock);

    const formBlock = createBlock('form', 'Share via Email', {
      kind: 'form',
      preset: 'email-compose',
      description: 'Share the current result',
      fields: [],
      submissionTarget: { toolName: 'SendEmailWithAttachments', serverId: 'mcp_MailTools', staticArgs: {} },
      status: 'editing'
    }, true, undefined, { originTool: `share-form:${searchBlock.id}` });
    engine.onBlockCreated(formBlock);

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'form',
      turnId: 'turn-spfx',
      sourceBlockId: searchBlock.id,
      derivedBlockId: formBlock.id
    });
    expect(engine.getCurrentArtifacts()[formBlock.id]).toMatchObject({
      artifactKind: 'form',
      sourceBlockId: searchBlock.id,
      sourceTurnId: 'turn-spfx'
    });
  });

  it('captures source event correlation from the current task context for later producers', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.emitEvent({
      eventName: 'task.focused',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'focus-1',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-root',
      payload: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-search',
      blockType: 'search-results'
    });

    expect(engine.captureCurrentSourceContext()).toEqual(expect.objectContaining({
      sourceBlockId: 'block-search',
      sourceTaskKind: 'focus',
      sourceEventName: 'task.focused',
      correlationId: 'focus-1',
      sourceTurnId: 'turn-spfx',
      sourceRootTurnId: 'turn-root'
    }));
  });

  it('tracks centralized action-panel selections as current HIE selection context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.emitEvent({
      eventName: 'task.selection.updated',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'selection-1',
      turnId: 'turn-mail',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-root',
      payload: {
        sourceBlockId: 'block-mail',
        sourceBlockType: 'markdown',
        sourceBlockTitle: 'Inbox results',
        selectedItems: [
          {
            index: 1,
            title: 'Budget follow-up',
            kind: 'email',
            itemType: 'email',
            targetContext: {
              mailItemId: 'mail-123',
              source: 'hie-selection'
            }
          }
        ]
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-mail',
      blockType: 'markdown'
    });

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'select',
      sourceBlockId: 'block-mail',
      selectedItems: [
        expect.objectContaining({
          index: 1,
          title: 'Budget follow-up',
          kind: 'email'
        })
      ],
      targetContext: expect.objectContaining({
        mailItemId: 'mail-123'
      })
    });
    expect(engine.captureCurrentSourceContext()).toEqual(expect.objectContaining({
      sourceBlockId: 'block-mail',
      sourceTaskKind: 'select',
      targetContext: expect.objectContaining({
        mailItemId: 'mail-123'
      })
    }));

    const selectionCall = sendContextMessage.mock.calls.find(
      (call) => typeof call[0] === 'string' && call[0].includes('The user selected options from "Inbox results".')
    );
    expect(selectionCall).toBeDefined();
    expect(selectionCall?.[1]).toBe(false);
  });

  it('removes dismissed forms from projected current state', () => {
    jest.useFakeTimers();
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    try {
      engine.setCurrentTurnId('turn-spfx');
      const formBlock = createBlock('form', 'Share via Email', {
        kind: 'form',
        preset: 'email-compose',
        description: 'Share the current result',
        fields: [],
        submissionTarget: { toolName: 'SendEmailWithAttachments', serverId: 'mcp_MailTools', staticArgs: {} },
        status: 'editing'
      }, true);
      engine.onBlockCreated(formBlock, {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx',
        sourceTurnId: 'turn-spfx',
        sourceRootTurnId: 'turn-spfx'
      });

      expect(engine.getProjectedCurrentStateSummary()).toContain('Current form: Share via Email.');

      sendContextMessage.mockClear();
      engine.onBlockRemoved(formBlock.id);

      expect(engine.getCurrentTaskContext()).toMatchObject({
        kind: 'form',
        derivedBlockId: formBlock.id,
        formStatus: 'dismissed'
      });
      expect(engine.getCurrentArtifacts()[formBlock.id]).toMatchObject({
        status: 'dismissed'
      });
      expect(engine.getProjectedCurrentStateSummary()).not.toContain('Current form: Share via Email.');
      expect(engine.getProjectedCurrentStateSummary()).toContain('no longer visible');

      const dismissalCall = sendContextMessage.mock.calls.find(
        (call) => typeof call[0] === 'string' && call[0].includes('The email-compose form was closed.')
      );
      expect(dismissalCall).toBeDefined();
      expect(dismissalCall?.[1]).toBe(false);
      jest.runOnlyPendingTimers();
    } finally {
      jest.useRealTimers();
    }
  });

  it('derives MCP target context from selection-list site choices', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.onBlockInteraction({
      blockId: 'block-site-select',
      blockType: 'selection-list',
      action: 'select',
      eventName: 'block.interaction.select',
      exposurePolicy: { mode: 'response-triggering', relevance: 'foreground' },
      source: 'block-ui',
      surface: 'action-panel',
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
      },
      timestamp: Date.now()
    });

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'select',
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      }
    });
    expect(engine.captureCurrentSourceContext()).toEqual(expect.objectContaining({
      targetContext: expect.objectContaining({
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking'
      })
    }));
  });

  it('keeps explicit form submission target context in the current HIE source context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const formBlock = createBlock('form', 'Create folder in Documents', {
      kind: 'form',
      preset: 'folder-create',
      description: 'Create a folder in the selected site',
      fields: [],
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
    });

    engine.onBlockCreated(formBlock);

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'form',
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking'
      }
    });
    expect(engine.captureCurrentSourceContext()).toEqual(expect.objectContaining({
      targetContext: expect.objectContaining({
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking'
      })
    }));
  });

  it('promotes target-bearing site-info artifacts into the current HIE source context', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    const siteBlock = createBlock('site-info', 'copilot-test-cooking', {
      kind: 'site-info',
      siteName: 'copilot-test-cooking',
      siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
      libraries: ['Dokumente']
    });

    engine.onBlockCreated(siteBlock);

    expect(engine.getCurrentArtifacts()[siteBlock.id]).toMatchObject({
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      }
    });
    expect(engine.captureCurrentSourceContext()).toEqual(expect.objectContaining({
      sourceArtifactId: siteBlock.id,
      sourceBlockId: siteBlock.id,
      sourceBlockType: 'site-info',
      sourceBlockTitle: 'copilot-test-cooking',
      targetContext: expect.objectContaining({
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking'
      })
    }));
  });

  it('derives MCP target context from clicked list rows', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.onBlockInteraction({
      blockId: 'block-lists',
      blockType: 'list-items',
      action: 'click-list-row',
      eventName: 'block.interaction.click-list-row',
      exposurePolicy: { mode: 'response-triggering', relevance: 'foreground' },
      source: 'block-ui',
      surface: 'action-panel',
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
      },
      timestamp: Date.now()
    });

    expect(engine.getCurrentTaskContext()).toMatchObject({
      kind: 'click-list-row',
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
        siteId: 'contoso.sharepoint.com,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002',
        listId: '00000000-0000-0000-0000-000000000003',
        listUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Lists/Events',
        listName: 'Events'
      }
    });
    expect(engine.captureCurrentSourceContext()).toEqual(expect.objectContaining({
      targetContext: expect.objectContaining({
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test',
        listId: '00000000-0000-0000-0000-000000000003',
        listName: 'Events'
      })
    }));
  });

  it('creates follow-up user turns that keep the current root thread', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.emitEvent({
      eventName: 'task.focused',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'focus-1',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-root',
      payload: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-search',
      blockType: 'search-results'
    });

    const followUpTurn = engine.beginUserTurn('turn-share');

    expect(followUpTurn).toEqual({
      turnId: 'turn-share',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-spfx'
    });
    expect(engine.getTurnLineage('turn-share')).toEqual(followUpTurn);
    expect(engine.getCurrentTurnLineage()).toEqual(followUpTurn);
  });

  it('starts a new root thread for explicit reset phrases in auto mode', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    engine.emitEvent({
      eventName: 'task.focused',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'focus-2',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-root',
      payload: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-search',
      blockType: 'search-results'
    });

    const resetTurn = engine.beginUserTurn({
      turnId: 'turn-invoices',
      mode: 'auto',
      text: 'new topic: search invoices'
    });

    expect(resetTurn).toEqual({
      turnId: 'turn-invoices',
      rootTurnId: 'turn-invoices',
      parentTurnId: undefined
    });
    expect(engine.getCurrentTurnLineage()).toEqual(resetTurn);
    expect(engine.getRecentEvents().slice(-1)[0]).toMatchObject({
      eventName: 'thread.reset',
      turnId: 'turn-invoices',
      rootTurnId: 'turn-invoices',
      payload: expect.objectContaining({
        requestedMode: 'auto',
        resolvedMode: 'new-root',
        reason: 'explicit-reset-phrase'
      })
    });
  });

  it('replays queued block interactions through the shared projector path', () => {
    const sendContextMessage = jest.fn();
    const engine = new HybridInteractionEngine();

    engine.onBlockInteraction({
      blockId: 'block-1',
      blockType: 'search-results',
      action: 'summarize',
      payload: {
        title: 'Q4 Budget Plan',
        fileType: 'docx',
        url: 'https://contoso.sharepoint.com/sites/finance/Shared%20Documents/Q4-Budget.docx'
      },
      timestamp: Date.now()
    });

    expect(engine.getRecentEvents()).toHaveLength(0);

    engine.initialize(() => { /* no-op */ }, { sendContextMessage });

    expect(sendContextMessage).toHaveBeenCalledTimes(1);
    expect(sendContextMessage.mock.calls[0][0]).toContain('Trusted action: The user clicked Summarize on a search result.');
    expect(sendContextMessage.mock.calls[0][1]).toBe(true);
    expect(engine.getRecentEvents()[0]).toMatchObject({
      eventName: 'block.interaction.summarize',
      blockId: 'block-1',
      blockType: 'search-results'
    });
  });
});
