import type { IContextMessage, IHieEvent } from '../../services/hie/HIETypes';
import {
  buildStateCompassCards,
  buildHieLinkageNodes,
  buildHieTargetContextSummary,
  buildSituationPills,
  buildThreadNodes,
  describeHieEvent,
  formatHieContextPreview,
  getDefaultExpandedCompassCard,
  formatHieTaskMeta
} from './hieInspectorHelpers';

describe('hieInspectorHelpers', () => {
  it('builds distinct root, parent, and current thread nodes', () => {
    expect(buildThreadNodes({
      turnId: 'turn-current',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-parent'
    })).toEqual([
      { key: 'root-turn-root', role: 'Root', value: 'turn-root' },
      { key: 'parent-turn-parent', role: 'Parent', value: 'turn-parent' },
      { key: 'current-turn-current', role: 'Current', value: 'turn-current' }
    ]);
  });

  it('formats thread lifecycle events with resolved mode and reason', () => {
    const event: IHieEvent = {
      eventId: 'evt-1',
      eventName: 'thread.reset',
      source: 'hie',
      surface: 'unknown',
      correlationId: 'turn-2',
      turnId: 'turn-2',
      rootTurnId: 'turn-2',
      timestamp: Date.now(),
      payload: {
        resolvedMode: 'new-root',
        reason: 'explicit-reset-phrase'
      },
      exposurePolicy: { mode: 'store-only', relevance: 'contextual' }
    };

    expect(describeHieEvent(event)).toBe('thread.reset · store-only · new-root · explicit-reset-phrase');
  });

  it('unwraps and clips projected context previews', () => {
    const message: IContextMessage = {
      contextType: 'visual',
      text: '[Task context: The user focused items from "Search: SPFx" in the action panel.]',
      blockIds: ['block-1'],
      sentAt: Date.now()
    };

    expect(formatHieContextPreview(message)).toBe('Task context: The user focused items from "Search: SPFx" in the action panel.');
  });

  it('builds linkage nodes from source block to derived artifacts', () => {
    expect(buildHieLinkageNodes({
      kind: 'form',
      eventName: 'form.opened',
      correlationId: 'form-1',
      turnId: 'turn-share',
      sourceBlockId: 'block-search',
      sourceBlockTitle: 'Search: SPFx',
      derivedBlockId: 'block-form',
      derivedBlockTitle: 'Share via Email',
      updatedAt: Date.now()
    }, {
      'block-recap': {
        artifactId: 'block-recap',
        artifactKind: 'recap',
        blockId: 'block-recap',
        blockType: 'info-card',
        title: 'Recap: Search: SPFx',
        sourceEventName: 'artifact.recap.ready',
        correlationId: 'recap-1',
        sourceTurnId: 'turn-spfx',
        status: 'ready',
        updatedAt: 1
      },
      'block-form': {
        artifactId: 'block-form',
        artifactKind: 'form',
        sourceArtifactId: 'block-recap',
        blockId: 'block-form',
        blockType: 'form',
        title: 'Share via Email',
        sourceEventName: 'form.opened',
        correlationId: 'form-1',
        sourceTurnId: 'turn-share',
        status: 'opened',
        updatedAt: 2
      }
    })).toEqual([
      { key: 'source-block-search', label: 'Source', value: 'Search: SPFx', meta: 'form.opened · corr form-1 · turn turn-share' },
      { key: 'artifact-block-recap', label: 'Recap', value: 'Recap: Search: SPFx', meta: 'artifact.recap.ready · corr recap-1 · turn turn-spfx' },
      { key: 'artifact-block-form', label: 'Form', value: 'Share via Email', meta: 'form.opened · corr form-1 · turn turn-share' }
    ]);
  });

  it('falls back to the latest artifact linkage when there is no active task', () => {
    expect(buildHieLinkageNodes(undefined, {
      'block-summary': {
        artifactId: 'block-summary',
        artifactKind: 'summary',
        blockId: 'block-summary',
        blockType: 'info-card',
        title: 'Summary: TeamsFx.pdf',
        status: 'ready',
        updatedAt: 10
      }
    })).toEqual([
      { key: 'artifact-block-summary', label: 'Summary', value: 'Summary: TeamsFx.pdf', meta: undefined }
    ]);
  });

  it('formats task metadata with correlation and turn details', () => {
    expect(formatHieTaskMeta({
      kind: 'focus',
      eventName: 'task.focused',
      correlationId: 'focus-1',
      turnId: 'turn-12345678901234567890',
      updatedAt: Date.now()
    })).toBe('task.focused · corr focus-1 · turn turn-12345...7890');
  });

  it('builds compass cards with merged detail from thread to artifacts', () => {
    const cards = buildStateCompassCards({
      turnId: 'turn-share',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-spfx'
    }, {
      eventId: 'evt-thread',
      eventName: 'thread.continued',
      source: 'hie',
      surface: 'unknown',
      correlationId: 'turn-share',
      turnId: 'turn-share',
      rootTurnId: 'turn-root',
      parentTurnId: 'turn-spfx',
      timestamp: Date.now(),
      payload: {
        resolvedMode: 'inherit',
        reason: 'active-task-context'
      },
      exposurePolicy: { mode: 'store-only', relevance: 'contextual' }
    }, {
      kind: 'form',
      eventName: 'form.opened',
      correlationId: 'form-1',
      turnId: 'turn-share',
      sourceBlockId: 'block-search',
      sourceBlockTitle: 'Search: SPFx',
      derivedBlockId: 'block-form',
      derivedBlockTitle: 'Share via Email',
      updatedAt: Date.now()
    }, {
      'block-form': {
        artifactId: 'block-form',
        artifactKind: 'form',
        sourceArtifactId: 'block-recap',
        blockId: 'block-form',
        blockType: 'form',
        title: 'Share via Email',
        sourceEventName: 'form.opened',
        correlationId: 'form-1',
        sourceTurnId: 'turn-share',
        status: 'opened',
        updatedAt: 2
      },
      'block-recap': {
        artifactId: 'block-recap',
        artifactKind: 'recap',
        blockId: 'block-recap',
        blockType: 'info-card',
        title: 'Recap: Search: SPFx',
        sourceEventName: 'artifact.recap.ready',
        correlationId: 'recap-1',
        sourceTurnId: 'turn-spfx',
        status: 'ready',
        updatedAt: 1
      }
    }, {
      snapshotId: 'snapshot-1',
      correlationId: 'corr-1',
      createdAt: 1,
      summary: 'Search and recap blocks are visible.',
      flowSummary: 'User is reviewing SPFx results and a recap.',
      hasDuplicateTypes: false,
      referencesNewestFirst: [],
      blocks: [
        {
          blockId: 'block-search',
          blockType: 'search-results',
          title: 'Search: SPFx',
          state: 'ready',
          summary: 'Found four matching documents.',
          itemCount: 4,
          references: [],
          createdAt: 1,
          updatedAt: 1
        },
        {
          blockId: 'block-recap',
          blockType: 'info-card',
          title: 'Recap: Search: SPFx',
          state: 'ready',
          summary: 'SPFx results recap.',
          itemCount: 0,
          references: [],
          createdAt: 2,
          updatedAt: 2
        }
      ]
    });

    expect(cards).toHaveLength(4);
    expect(cards[0].key).toBe('thread');
    expect(cards[0].value).toBe('Follow-up thread');
    expect(cards[0].isEmpty).toBe(false);
    expect(cards[0].detail).toContain('Continues the current thread');

    expect(cards[1].key).toBe('task');
    expect(cards[1].value).toBe('Form: Share via Email (opened)');
    expect(cards[1].isEmpty).toBe(false);
    expect(cards[1].detail).toContain('Search: SPFx');

    expect(cards[2].key).toBe('content');
    expect(cards[2].value).toBe('Search: SPFx');
    expect(cards[2].isEmpty).toBe(false);

    expect(cards[3].key).toBe('artifacts');
    expect(cards[3].value).toBe('Share via Email');
    expect(cards[3].isEmpty).toBe(false);
  });

  it('marks empty compass cards and picks the first meaningful card', () => {
    const emptyCards = buildStateCompassCards(undefined, undefined, undefined, {}, {
      snapshotId: 'snapshot-empty',
      correlationId: 'corr-empty',
      createdAt: 1,
      summary: '',
      flowSummary: '',
      hasDuplicateTypes: false,
      referencesNewestFirst: [],
      blocks: []
    });

    expect(emptyCards[0].isEmpty).toBe(true);
    expect(emptyCards[1].isEmpty).toBe(true);
    expect(emptyCards[2].isEmpty).toBe(true);
    expect(emptyCards[3].isEmpty).toBe(true);
    expect(getDefaultExpandedCompassCard(emptyCards)).toBeUndefined();

    const meaningfulCards = buildStateCompassCards({
      turnId: 'turn-1',
      rootTurnId: 'turn-1'
    }, undefined, {
      kind: 'focus',
      eventName: 'task.focused',
      turnId: 'turn-1',
      sourceBlockTitle: 'Search: SPFx',
      updatedAt: 1
    });

    expect(getDefaultExpandedCompassCard(meaningfulCards)).toBe('thread');
  });

  it('builds situation pills from flow, verbosity, expression, and target', () => {
    const pills = buildSituationPills({
      flowState: { flowName: 'Search and Drill Down', stepIndex: 2, totalSteps: 3 },
      verbosity: 'minimal',
      expressionTrigger: { triggerId: 'search-many', expression: 'happy', firedAt: Date.now() },
      targetSummary: { value: 'Documents', description: 'Library: Documents' }
    });

    expect(pills).toHaveLength(4);
    expect(pills[0].key).toBe('flow');
    expect(pills[0].label).toContain('Search and Drill Down');
    expect(pills[1].key).toBe('verbosity');
    expect(pills[1].label).toBe('Verbosity: minimal');
    expect(pills[2].key).toBe('expression');
    expect(pills[2].label).toContain('happy');
    expect(pills[3].key).toBe('target');
    expect(pills[3].label).toContain('Documents');
  });

  it('shows idle pill when no situation data is active', () => {
    const pills = buildSituationPills({
      verbosity: 'normal'
    });

    expect(pills).toHaveLength(1);
    expect(pills[0].key).toBe('idle');
    expect(pills[0].label).toBe('Idle');
  });

  it('summarizes the current MCP target context from task and artifact lineage', () => {
    expect(buildHieTargetContextSummary({
      kind: 'form',
      eventName: 'form.opened',
      correlationId: 'form-1',
      sourceBlockTitle: 'Create folder in Documents',
      targetContext: {
        siteUrl: 'https://contoso.sharepoint.com/sites/copilot-test-cooking',
        siteName: 'copilot-test-cooking',
        source: 'explicit-user'
      },
      updatedAt: 1
    }, {
      'block-form': {
        artifactId: 'block-form',
        artifactKind: 'form',
        blockId: 'block-form',
        blockType: 'form',
        title: 'Create folder in Documents',
        targetContext: {
          siteId: 'contoso.sharepoint.com,site-id,web-id',
          documentLibraryId: 'drive-123',
          documentLibraryName: 'Documents'
        },
        status: 'opened',
        updatedAt: 2
      }
    })).toEqual({
      value: 'Documents',
      description: 'Site: copilot-test-cooking · Library: Documents',
      meta: 'explicit user target · site id ready · library id ready'
    });
  });
});
