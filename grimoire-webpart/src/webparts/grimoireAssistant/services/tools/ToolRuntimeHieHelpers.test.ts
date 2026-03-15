import { createBlock } from '../../models/IBlock';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { trackCreatedBlock } from './ToolRuntimeHieHelpers';

describe('ToolRuntimeHieHelpers', () => {
  beforeEach(() => {
    hybridInteractionEngine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });
  });

  afterEach(() => {
    hybridInteractionEngine.reset();
  });

  it('promotes derived runtime outputs into explicit HIE artifacts from source context', () => {
    const store = {
      pushBlock: jest.fn()
    };
    const summaryBlock = createBlock('info-card', 'Summary: SPFx guidance', {
      kind: 'info-card',
      heading: 'Summary: SPFx guidance',
      body: 'SPFx guidance summary content'
    });

    trackCreatedBlock(store, summaryBlock, {
      sourceContext: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx',
        sourceTaskKind: 'summarize',
        sourceEventName: 'block.interaction.summarize',
        correlationId: 'iact-1',
        sourceTurnId: 'turn-spfx',
        sourceRootTurnId: 'turn-root',
        sourceParentTurnId: 'turn-root'
      }
    });

    expect(store.pushBlock).toHaveBeenCalledWith(summaryBlock);
    expect(hybridInteractionEngine.getCurrentArtifacts()[summaryBlock.id]).toMatchObject({
      artifactKind: 'summary',
      blockType: 'info-card',
      sourceBlockId: 'block-search',
      sourceTaskKind: 'summarize',
      sourceEventName: 'block.interaction.summarize',
      sourceTurnId: 'turn-spfx',
      sourceRootTurnId: 'turn-root'
    });
    expect(hybridInteractionEngine.getCurrentTaskContext()).toMatchObject({
      kind: 'summarize',
      sourceBlockId: 'block-search',
      derivedBlockId: summaryBlock.id,
      derivedBlockTitle: 'Summary: SPFx guidance',
      turnId: 'turn-spfx',
      rootTurnId: 'turn-root'
    });
  });

  it('promotes file preview outputs as preview artifacts', () => {
    const store = {
      pushBlock: jest.fn()
    };
    const previewBlock = createBlock('file-preview', 'SPFx.pdf', {
      kind: 'file-preview',
      fileName: 'SPFx.pdf',
      fileUrl: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
      fileType: 'pdf',
      metadata: {}
    });

    trackCreatedBlock(store, previewBlock, {
      sourceContext: {
        sourceBlockId: 'block-search',
        sourceBlockType: 'search-results',
        sourceBlockTitle: 'Search: SPFx',
        sourceTaskKind: 'look',
        sourceEventName: 'block.interaction.look',
        correlationId: 'iact-2',
        sourceTurnId: 'turn-spfx',
        sourceRootTurnId: 'turn-root'
      }
    });

    expect(hybridInteractionEngine.getCurrentArtifacts()[previewBlock.id]).toMatchObject({
      artifactKind: 'preview',
      blockType: 'file-preview',
      sourceBlockId: 'block-search',
      sourceTaskKind: 'look'
    });
  });
});
