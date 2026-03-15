jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('../nano/NanoService', () => ({
  getNanoService: jest.fn()
}));

import { createBlock } from '../../models/IBlock';
import type { ITranscriptEntry } from '../../store/useGrimoireStore';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { SessionShareFormatter, hasShareableSessionContent } from './SessionShareFormatter';

describe('SessionShareFormatter', () => {
  beforeEach(() => {
    hybridInteractionEngine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });
  });

  afterEach(() => {
    hybridInteractionEngine.reset();
  });

  it('excludes system transcript, forms, and confirmation dialogs from shared content', () => {
    const transcript: ITranscriptEntry[] = [
      { role: 'system', text: 'Internal context', timestamp: new Date('2026-03-07T10:00:00.000Z') },
      { role: 'user', text: 'Please share the current work.', timestamp: new Date('2026-03-07T10:01:00.000Z') },
      { role: 'assistant', text: 'Here is the current session summary.', timestamp: new Date('2026-03-07T10:02:00.000Z') }
    ];
    const searchBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'SPFx',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx guidance',
          summary: 'Implementation guidance for SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx migration',
          summary: 'Migration notes for older solutions.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFxMigration.docx',
          sources: ['copilot-search']
        }
      ]
    });
    const formBlock = createBlock('form', 'Email compose', {
      kind: 'form',
      preset: 'email-compose',
      description: 'Internal form',
      fields: [],
      submissionTarget: { toolName: 'SendEmailWithAttachments', serverId: 'mcp_MailTools', staticArgs: {} },
      status: 'editing'
    });
    const confirmationBlock = createBlock('confirmation-dialog', 'Confirm', {
      kind: 'confirmation-dialog',
      message: 'Internal confirmation',
      confirmLabel: 'Yes',
      cancelLabel: 'No',
      onConfirmAction: 'confirm'
    });

    const content = new SessionShareFormatter().format({
      blocks: [searchBlock, formBlock, confirmationBlock],
      transcript,
      activeBlockId: searchBlock.id
    });

    expect(content.subject).toBe('Grimoire share — Search: SPFx');
    expect(content.plainText).toContain('User: Please share the current work.');
    expect(content.plainText).toContain('Assistant: Here is the current session summary.');
    expect(content.plainText).not.toContain('Internal context');
    expect(content.plainText).toContain('Visible results');
    expect(content.plainText).toContain('Search: SPFx');
    expect(content.plainText).not.toContain('Email compose');
    expect(content.plainText).not.toContain('Confirm');
    expect(content.markdown).toContain('# Grimoire share');
    expect(content.markdown).toContain('## Conversation');
    expect(content.markdown).toContain('## Visible Results');
  });

  it('falls back to a generic session subject when there is no shareable block', () => {
    const content = new SessionShareFormatter().format({
      blocks: [
        createBlock('form', 'Email compose', {
          kind: 'form',
          preset: 'email-compose',
          fields: [],
          submissionTarget: { toolName: 'SendEmailWithAttachments', serverId: 'mcp_MailTools', staticArgs: {} },
          status: 'editing'
        })
      ],
      transcript: [
        { role: 'user', text: 'Only transcript content', timestamp: new Date('2026-03-07T10:00:00.000Z') }
      ]
    });

    expect(content.subject).toBe('Grimoire share — Session summary');
    expect(content.plainText).toContain('Only transcript content');
  });

  it('builds detailed share content and attachments from the selected visible items', () => {
    const selectedBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'SPFx',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx guidance',
          summary: 'Implementation guidance for SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
          author: 'Test User',
          fileType: 'docx',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx migration',
          summary: 'Migration notes for older solutions.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFxMigration.docx',
          author: 'SharePoint App',
          fileType: 'docx',
          sources: ['sharepoint-search']
        }
      ]
    });
    const selectedContent = new SessionShareFormatter().format({
      blocks: [selectedBlock],
      transcript: [],
      activeBlockId: selectedBlock.id,
      selectedActionIndices: [2]
    });

    expect(selectedContent.detailedPlainText).toContain('Current visible content');
    expect(selectedContent.detailedPlainText).toContain('SPFx migration');
    expect(selectedContent.detailedPlainText).not.toContain('SPFx guidance');
    expect(selectedContent.emailPlainText).toBe(selectedContent.detailedPlainText);
    expect(selectedContent.attachmentUris).toEqual([
      'https://tenant.sharepoint.com/sites/dev/SPFxMigration.docx'
    ]);
  });

  it('prefers the current HIE recap artifact and exports its body', () => {
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
    const recapBlock = createBlock('info-card', 'Recap: Search: SPFx', {
      kind: 'info-card',
      heading: 'Recap: Search: SPFx',
      body: 'SPFx results focus on the SPFx framework overview in multiple language variants.'
    }, true, undefined, { originTool: `block-recap:${searchBlock.id}` });

    hybridInteractionEngine.emitEvent({
      eventName: 'artifact.recap.ready',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-recap-1',
      payload: {
        sourceBlockId: searchBlock.id,
        sourceBlockType: searchBlock.type,
        sourceBlockTitle: searchBlock.title,
        derivedBlockId: recapBlock.id,
        derivedBlockTitle: recapBlock.title
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: recapBlock.id,
      blockType: recapBlock.type
    });

    const content = new SessionShareFormatter().format({
      blocks: [searchBlock, recapBlock],
      transcript: [],
      activeBlockId: searchBlock.id
    });

    expect(content.subject).toBe('Grimoire share — Recap: Search: SPFx');
    expect(content.plainText).toContain('Recap: Search: SPFx');
    expect(content.plainText).toContain('SPFx results focus on the SPFx framework overview');
    expect(content.detailedPlainText).toContain('SPFx results focus on the SPFx framework overview');
    expect(content.attachmentUris).toEqual([]);
  });

  it('prefers HIE-derived summary artifacts over the stale active source block', () => {
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
    hybridInteractionEngine.onBlockCreated(searchBlock);

    hybridInteractionEngine.onBlockInteraction({
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
    hybridInteractionEngine.onBlockCreated(summaryBlock);

    const content = new SessionShareFormatter().format({
      blocks: [searchBlock, summaryBlock],
      transcript: [],
      activeBlockId: searchBlock.id
    });

    expect(content.subject).toBe('Grimoire share — Summary: SPFx guidance');
    expect(content.plainText).toContain('Summary: SPFx guidance');
    expect(content.plainText).toContain('SPFx guidance summary content');
    expect(content.detailedPlainText).toContain('SPFx guidance summary content');
    expect(content.detailedPlainText).not.toContain('Implementation guidance for SharePoint Framework.');
  });

  it('keeps recap scope when a share form is opened from that recap', () => {
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
    const recapBlock = createBlock('info-card', 'Recap: Search: SPFx', {
      kind: 'info-card',
      heading: 'Recap: Search: SPFx',
      body: 'SPFx results focus on the SPFx framework overview in multiple language variants.'
    }, true, undefined, { originTool: `block-recap:${searchBlock.id}` });
    const shareFormBlock = createBlock('form', 'Share via Email', {
      kind: 'form',
      preset: 'email-compose',
      description: 'Share recap',
      fields: [],
      submissionTarget: { toolName: 'SendEmailWithAttachments', serverId: 'mcp_MailTools', staticArgs: {} },
      status: 'editing'
    }, true, undefined, { originTool: `share-form:${searchBlock.id}` });

    hybridInteractionEngine.emitEvent({
      eventName: 'artifact.recap.ready',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-recap-form-1',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-root',
      payload: {
        sourceBlockId: searchBlock.id,
        sourceBlockType: searchBlock.type,
        sourceBlockTitle: searchBlock.title,
        derivedBlockId: recapBlock.id,
        derivedBlockTitle: recapBlock.title,
        sourceTurnId: 'turn-spfx',
        sourceRootTurnId: 'turn-root',
        sourceParentTurnId: 'turn-root'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: recapBlock.id,
      blockType: recapBlock.type
    });

    hybridInteractionEngine.emitEvent({
      eventName: 'form.opened',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-recap-form-2',
      turnId: 'turn-share',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-spfx',
      payload: {
        blockId: shareFormBlock.id,
        blockTitle: shareFormBlock.title,
        preset: 'email-compose',
        sourceBlockId: recapBlock.id,
        sourceBlockType: recapBlock.type,
        sourceBlockTitle: recapBlock.title,
        sourceArtifactId: recapBlock.id,
        sourceTaskKind: 'recap',
        linkedSourceBlockId: searchBlock.id,
        sourceTurnId: 'turn-spfx',
        sourceRootTurnId: 'turn-root',
        sourceParentTurnId: 'turn-root'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: shareFormBlock.id,
      blockType: shareFormBlock.type
    });

    const transcript: ITranscriptEntry[] = [
      { role: 'user', text: 'i am searching for doc about animals', timestamp: new Date('2026-03-07T10:00:00.000Z'), turnId: 'turn-root', rootTurnId: 'turn-root' },
      { role: 'assistant', text: 'I found 3 documents. They are in the panel.', timestamp: new Date('2026-03-07T10:00:05.000Z'), turnId: 'turn-root', rootTurnId: 'turn-root' },
      { role: 'user', text: 'what about spfx ?', timestamp: new Date('2026-03-07T10:01:00.000Z'), turnId: 'turn-spfx', rootTurnId: 'turn-root', parentTurnId: 'turn-root' },
      { role: 'assistant', text: 'I opened a capability overview in the panel.', timestamp: new Date('2026-03-07T10:01:05.000Z'), turnId: 'turn-spfx', rootTurnId: 'turn-root', parentTurnId: 'turn-root' },
      { role: 'user', text: 'share it by email', timestamp: new Date('2026-03-07T10:03:00.000Z'), turnId: 'turn-share', rootTurnId: 'turn-root', parentTurnId: 'turn-spfx' },
      { role: 'assistant', text: 'I opened the compose form.', timestamp: new Date('2026-03-07T10:03:05.000Z'), turnId: 'turn-share', rootTurnId: 'turn-root', parentTurnId: 'turn-spfx' }
    ];

    const content = new SessionShareFormatter().format({
      blocks: [searchBlock, recapBlock, shareFormBlock],
      transcript,
      activeBlockId: searchBlock.id
    });

    expect(content.subject).toBe('Grimoire share — Recap: Search: SPFx');
    expect(content.plainText).toContain('Recap: Search: SPFx');
    expect(content.plainText).toContain('SPFx results focus on the SPFx framework overview');
    expect(content.plainText).not.toContain('Implementation guidance for SharePoint Framework.');
    expect(content.plainText).toContain('User: what about spfx ?');
    expect(content.plainText).toContain('Assistant: I opened a capability overview in the panel.');
    expect(content.plainText).not.toContain('share it by email');
    expect(content.detailedPlainText).toContain('SPFx results focus on the SPFx framework overview');
    expect(content.attachmentUris).toEqual([]);
  });

  it('uses exact transcript turn IDs before falling back to transcript heuristics', () => {
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
    const recapBlock = createBlock('info-card', 'Recap: Search: SPFx', {
      kind: 'info-card',
      heading: 'Recap: Search: SPFx',
      body: 'SPFx results focus on the SPFx framework overview in multiple language variants.'
    }, true, undefined, { originTool: `block-recap:${searchBlock.id}` });
    const transcript: ITranscriptEntry[] = [
      { role: 'user', text: 'show me something else', timestamp: new Date('2026-03-07T10:00:00.000Z'), turnId: 'turn-old' },
      { role: 'assistant', text: 'Here are other documents.', timestamp: new Date('2026-03-07T10:00:05.000Z'), turnId: 'turn-old' },
      { role: 'user', text: 'what about that framework again?', timestamp: new Date('2026-03-07T10:01:00.000Z'), turnId: 'turn-spfx' },
      { role: 'assistant', text: 'I opened the relevant framework documents.', timestamp: new Date('2026-03-07T10:01:05.000Z'), turnId: 'turn-spfx' },
      { role: 'user', text: 'share it by email', timestamp: new Date('2026-03-07T10:03:00.000Z'), turnId: 'turn-share' },
      { role: 'assistant', text: 'I opened the compose form.', timestamp: new Date('2026-03-07T10:03:05.000Z'), turnId: 'turn-share' }
    ];

    hybridInteractionEngine.emitEvent({
      eventName: 'artifact.recap.ready',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-turn-1',
      turnId: 'turn-spfx',
      payload: {
        sourceBlockId: searchBlock.id,
        sourceBlockType: searchBlock.type,
        sourceBlockTitle: searchBlock.title,
        derivedBlockId: recapBlock.id,
        derivedBlockTitle: recapBlock.title,
        sourceTurnId: 'turn-spfx'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: recapBlock.id,
      blockType: recapBlock.type
    });

    const content = new SessionShareFormatter().format({
      blocks: [searchBlock, recapBlock],
      transcript,
      activeBlockId: searchBlock.id
    });

    expect(content.plainText).toContain('User: what about that framework again?');
    expect(content.plainText).toContain('Assistant: I opened the relevant framework documents.');
    expect(content.plainText).not.toContain('share it by email');
    expect(content.plainText).not.toContain('show me something else');
  });

  it('exports the full root thread up to the source turn when root turn lineage exists', () => {
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
    const recapBlock = createBlock('info-card', 'Recap: Search: SPFx', {
      kind: 'info-card',
      heading: 'Recap: Search: SPFx',
      body: 'SPFx results focus on the SPFx framework overview in multiple language variants.'
    }, true, undefined, { originTool: `block-recap:${searchBlock.id}` });
    const transcript: ITranscriptEntry[] = [
      {
        role: 'user',
        text: 'i am searching for doc about animals',
        timestamp: new Date('2026-03-07T10:00:00.000Z'),
        turnId: 'turn-root',
        rootTurnId: 'turn-root'
      },
      {
        role: 'assistant',
        text: 'I found 3 documents. They are in the panel.',
        timestamp: new Date('2026-03-07T10:00:05.000Z'),
        turnId: 'turn-root',
        rootTurnId: 'turn-root'
      },
      {
        role: 'user',
        text: 'what about spfx ?',
        timestamp: new Date('2026-03-07T10:01:00.000Z'),
        turnId: 'turn-spfx',
        rootTurnId: 'turn-root',
        parentTurnId: 'turn-root'
      },
      {
        role: 'assistant',
        text: 'I opened the relevant framework documents.',
        timestamp: new Date('2026-03-07T10:01:05.000Z'),
        turnId: 'turn-spfx',
        rootTurnId: 'turn-root',
        parentTurnId: 'turn-root'
      },
      {
        role: 'user',
        text: 'share it by email',
        timestamp: new Date('2026-03-07T10:03:00.000Z'),
        turnId: 'turn-share',
        rootTurnId: 'turn-root',
        parentTurnId: 'turn-spfx'
      },
      {
        role: 'assistant',
        text: 'I opened the compose form.',
        timestamp: new Date('2026-03-07T10:03:05.000Z'),
        turnId: 'turn-share',
        rootTurnId: 'turn-root',
        parentTurnId: 'turn-spfx'
      }
    ];

    hybridInteractionEngine.emitEvent({
      eventName: 'artifact.recap.ready',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-thread-1',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-root',
      payload: {
        sourceBlockId: searchBlock.id,
        sourceBlockType: searchBlock.type,
        sourceBlockTitle: searchBlock.title,
        derivedBlockId: recapBlock.id,
        derivedBlockTitle: recapBlock.title,
        sourceTurnId: 'turn-spfx',
        sourceRootTurnId: 'turn-root',
        sourceParentTurnId: 'turn-root'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: recapBlock.id,
      blockType: recapBlock.type
    });

    const content = new SessionShareFormatter().format({
      blocks: [searchBlock, recapBlock],
      transcript,
      activeBlockId: searchBlock.id
    });

    expect(content.plainText).toContain('User: i am searching for doc about animals');
    expect(content.plainText).toContain('Assistant: I found 3 documents. They are in the panel.');
    expect(content.plainText).toContain('User: what about spfx ?');
    expect(content.plainText).toContain('Assistant: I opened the relevant framework documents.');
    expect(content.plainText).not.toContain('share it by email');
    expect(content.plainText).not.toContain('I opened the compose form.');
  });

  it('prefers the current HIE task context over a stale active block', () => {
    const animalBlock = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'animals',
      totalCount: 1,
      source: 'copilot-search',
      results: [
        {
          title: 'Animal facts',
          summary: 'Animal summary.',
          url: 'https://tenant.sharepoint.com/sites/dev/animals.docx',
          sources: ['copilot-search']
        }
      ]
    });
    const spfxBlock = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'SPFx',
      totalCount: 2,
      source: 'copilot-search',
      results: [
        {
          title: 'SPFx guidance',
          summary: 'Implementation guidance for SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.docx',
          sources: ['copilot-search']
        },
        {
          title: 'SPFx migration',
          summary: 'Migration notes for older solutions.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFxMigration.docx',
          sources: ['copilot-search']
        }
      ]
    });

    hybridInteractionEngine.emitEvent({
      eventName: 'task.focused',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-focused-1',
      payload: {
        sourceBlockId: spfxBlock.id,
        sourceBlockType: spfxBlock.type,
        sourceBlockTitle: spfxBlock.title,
        selectedItems: [{ index: 2, title: 'SPFx migration' }]
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: spfxBlock.id,
      blockType: spfxBlock.type
    });

    const content = new SessionShareFormatter().format({
      blocks: [animalBlock, spfxBlock],
      transcript: [],
      activeBlockId: animalBlock.id,
      selectedActionIndices: [1]
    });

    expect(content.subject).toBe('Grimoire share — Search: SPFx');
    expect(content.plainText).toContain('Search: SPFx');
    expect(content.plainText).not.toContain('Animal summary.');
    expect(content.detailedPlainText).toContain('SPFx migration');
    expect(content.detailedPlainText).not.toContain('SPFx guidance');
  });

  it('falls back to visible blocks when HIE context points outside the current session', () => {
    const visibleBlock = createBlock('search-results', 'Search: SPFx', {
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

    hybridInteractionEngine.emitEvent({
      eventName: 'task.focused',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'share-stale-1',
      payload: {
        sourceBlockId: 'missing-block',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Missing block'
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'missing-block',
      blockType: 'search-results'
    });

    const content = new SessionShareFormatter().format({
      blocks: [visibleBlock],
      transcript: [],
      activeBlockId: visibleBlock.id
    });

    expect(content.subject).toBe('Grimoire share — Search: SPFx');
    expect(content.plainText).toContain('Search: SPFx');
    expect(content.detailedPlainText).toContain('Current visible content');
    expect(content.detailedPlainText).toContain('SPFx guidance');
  });

  it('does not reuse stale tracker lineage when HIE cannot resolve the current session', () => {
    const spfxBlock = createBlock('search-results', 'Search: SPFx', {
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
    const transcript: ITranscriptEntry[] = [
      {
        role: 'user',
        text: 'i am searching for doc about animals',
        timestamp: new Date('2026-03-07T10:00:00.000Z'),
        turnId: 'turn-animals',
        rootTurnId: 'turn-animals'
      },
      {
        role: 'assistant',
        text: 'I found 3 animal documents.',
        timestamp: new Date('2026-03-07T10:00:05.000Z'),
        turnId: 'turn-animals',
        rootTurnId: 'turn-animals'
      },
      {
        role: 'user',
        text: 'searching for doc about spfx',
        timestamp: new Date('2026-03-07T10:01:00.000Z'),
        turnId: 'turn-spfx',
        rootTurnId: 'turn-spfx'
      },
      {
        role: 'assistant',
        text: 'I found 3 SPFx documents.',
        timestamp: new Date('2026-03-07T10:01:05.000Z'),
        turnId: 'turn-spfx',
        rootTurnId: 'turn-spfx'
      }
    ];

    const content = new SessionShareFormatter().format({
      blocks: [spfxBlock],
      transcript,
      activeBlockId: spfxBlock.id
    });

    expect(content.subject).toBe('Grimoire share — Search: SPFx');
    expect(content.plainText).toContain('Search: SPFx');
    expect(content.plainText).toContain('User: searching for doc about spfx');
    expect(content.plainText).toContain('Assistant: I found 3 SPFx documents.');
    expect(content.plainText).toContain('i am searching for doc about animals');
  });

  it('falls back to the recent visible session when no exact lineage is available', () => {
    const visibleBlock = createBlock('search-results', 'Search: SPFx', {
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
    const transcript: ITranscriptEntry[] = [
      { role: 'user', text: 'older turn 1', timestamp: new Date('2026-03-07T10:00:00.000Z') },
      { role: 'assistant', text: 'older reply 1', timestamp: new Date('2026-03-07T10:00:05.000Z') },
      { role: 'user', text: 'older turn 2', timestamp: new Date('2026-03-07T10:01:00.000Z') },
      { role: 'assistant', text: 'older reply 2', timestamp: new Date('2026-03-07T10:01:05.000Z') },
      { role: 'user', text: 'recent turn 3', timestamp: new Date('2026-03-07T10:02:00.000Z') },
      { role: 'assistant', text: 'recent reply 3', timestamp: new Date('2026-03-07T10:02:05.000Z') },
      { role: 'user', text: 'recent turn 4', timestamp: new Date('2026-03-07T10:03:00.000Z') },
      { role: 'assistant', text: 'recent reply 4', timestamp: new Date('2026-03-07T10:03:05.000Z') }
    ];

    const content = new SessionShareFormatter().format({
      blocks: [visibleBlock],
      transcript,
      activeBlockId: visibleBlock.id
    });

    expect(content.plainText).not.toContain('older turn 1');
    expect(content.plainText).not.toContain('older reply 1');
    expect(content.plainText).toContain('older turn 2');
    expect(content.plainText).toContain('recent turn 4');
  });

  it('detects whether a session has anything shareable', () => {
    const transcript: ITranscriptEntry[] = [
      { role: 'system', text: 'system only', timestamp: new Date('2026-03-07T10:00:00.000Z') }
    ];

    expect(hasShareableSessionContent([], transcript)).toBe(false);
    expect(hasShareableSessionContent([], [
      ...transcript,
      { role: 'assistant', text: 'User-facing output', timestamp: new Date('2026-03-07T10:01:00.000Z') }
    ])).toBe(true);
    expect(hasShareableSessionContent([
      createBlock('info-card', 'Summary', {
        kind: 'info-card',
        heading: 'Summary',
        body: 'Visible summary'
      })
    ], transcript)).toBe(true);
  });
});
