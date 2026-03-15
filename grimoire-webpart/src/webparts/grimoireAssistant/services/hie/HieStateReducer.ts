import type {
  IHieArtifactRecord,
  IHieDerivedState,
  IHieEvent,
  IHieSelectedItem,
  IHieShellState,
  IHieTaskContext
} from './HIETypes';
import type { BlockType } from '../../models/IBlock';
import { isHieArtifactKind, resolveArtifactKindFromBlockContext } from './HieArtifactKindResolver';
import {
  deriveMcpTargetContextFromUnknown,
  mergeMcpTargetContexts,
  type IMcpTargetContext
} from '../mcp/McpTargetContext';

const NON_DERIVED_BLOCK_ARTIFACT_TYPES: ReadonlySet<BlockType> = new Set<BlockType>([
  'search-results',
  'selection-list',
  'confirmation-dialog',
  'form'
]);

function pickString(payload: Record<string, unknown>, key: string): string | undefined {
  const value = payload[key];
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

function pickBoolean(payload: Record<string, unknown>, key: string): boolean | undefined {
  const value = payload[key];
  return typeof value === 'boolean' ? value : undefined;
}

function pickBlockType(payload: Record<string, unknown>, key: string): BlockType | undefined {
  const value = payload[key];
  return typeof value === 'string' ? value as BlockType : undefined;
}

function pickSelectedItems(payload: Record<string, unknown>, key: string): IHieSelectedItem[] | undefined {
  const value = payload[key];
  if (!Array.isArray(value)) return undefined;

  const items = value
    .filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object')
    .map((entry) => ({
      index: typeof entry.index === 'number' ? entry.index : undefined,
      title: typeof entry.title === 'string'
        ? entry.title
        : (typeof entry.label === 'string' ? entry.label : undefined),
      url: typeof entry.url === 'string'
        ? entry.url
        : (typeof entry.id === 'string' && /^https?:\/\//i.test(entry.id)
          ? entry.id
          : (typeof entry.description === 'string' && /^https?:\/\//i.test(entry.description)
            ? entry.description
            : undefined)),
      kind: typeof entry.kind === 'string' ? entry.kind : undefined,
      itemType: typeof entry.itemType === 'string' ? entry.itemType : undefined,
      targetContext: deriveMcpTargetContextFromUnknown(entry, 'hie-selection')
    }))
    .filter((entry) => entry.index !== undefined || entry.title || entry.url || entry.targetContext);

  return items.length > 0 ? items : undefined;
}

function deriveTargetContextFromEvent(
  event: IHieEvent,
  selectedItems?: IHieSelectedItem[],
  previousTargetContext?: IMcpTargetContext
): IMcpTargetContext | undefined {
  const payloadContext = deriveMcpTargetContextFromUnknown(event.payload, 'hie-selection');
  const selectedItemsContext = deriveMcpTargetContextFromUnknown(
    selectedItems?.map((item) => item.targetContext || item),
    'hie-selection'
  );

  return mergeMcpTargetContexts(previousTargetContext, payloadContext, selectedItemsContext);
}

function buildTaskContext(
  event: IHieEvent,
  kind: string,
  overrides?: Partial<IHieTaskContext>
): IHieTaskContext {
  const payload = event.payload;
  const selectedItems = overrides?.selectedItems || pickSelectedItems(payload, 'selectedItems');
  const targetContext = overrides?.targetContext || deriveTargetContextFromEvent(
    event,
    selectedItems,
    overrides?.targetContext
  );
  return {
    kind,
    eventName: event.eventName,
    correlationId: overrides?.correlationId || event.correlationId,
    turnId: overrides?.turnId || pickString(payload, 'sourceTurnId') || event.turnId,
    rootTurnId: overrides?.rootTurnId || pickString(payload, 'sourceRootTurnId') || event.rootTurnId || overrides?.turnId || pickString(payload, 'sourceTurnId') || event.turnId,
    parentTurnId: overrides?.parentTurnId || pickString(payload, 'sourceParentTurnId') || event.parentTurnId,
    sourceBlockId: overrides?.sourceBlockId || pickString(payload, 'sourceBlockId') || event.blockId,
    sourceBlockType: overrides?.sourceBlockType || pickBlockType(payload, 'sourceBlockType') || event.blockType,
    sourceBlockTitle: overrides?.sourceBlockTitle || pickString(payload, 'sourceBlockTitle') || pickString(payload, 'blockTitle'),
    derivedBlockId: overrides?.derivedBlockId || pickString(payload, 'derivedBlockId') || pickString(payload, 'blockId'),
    derivedBlockType: overrides?.derivedBlockType || pickBlockType(payload, 'derivedBlockType') || event.blockType,
    derivedBlockTitle: overrides?.derivedBlockTitle || pickString(payload, 'derivedBlockTitle') || pickString(payload, 'title'),
    formPreset: overrides?.formPreset || pickString(payload, 'preset'),
    formStatus: overrides?.formStatus || pickString(payload, 'formStatus'),
    selectedItems,
    targetContext,
    updatedAt: event.timestamp
  };
}

function resolveInteractionTaskKind(eventName: string): string | undefined {
  switch (eventName) {
    case 'block.interaction.click-result':
    case 'block.interaction.click-folder':
    case 'block.interaction.click-file':
    case 'block.interaction.click-user':
    case 'block.interaction.click-list-row':
    case 'block.interaction.click-permission':
    case 'block.interaction.click-activity':
    case 'block.interaction.select':
    case 'block.interaction.click-item':
    case 'block.interaction.look':
    case 'block.interaction.summarize':
    case 'block.interaction.chat-about':
    case 'block.interaction.open-external':
    case 'block.interaction.navigate':
      return eventName.slice('block.interaction.'.length);
    default:
      return undefined;
  }
}

function shouldPromoteSearchBlockToTask(
  event: IHieEvent,
  nextState: IHieDerivedState
): boolean {
  const blockType = pickBlockType(event.payload, 'blockType') || event.blockType;
  if (blockType !== 'search-results') {
    return false;
  }

  if (pickString(event.payload, 'sourceTaskKind')) {
    return false;
  }

  const searchTurnId = pickString(event.payload, 'sourceTurnId') || event.turnId;
  const currentTask = nextState.taskContext;
  if (!currentTask) {
    return true;
  }

  if (currentTask.kind === 'search') {
    return true;
  }

  if (currentTask.sourceBlockId && currentTask.sourceBlockId === (pickString(event.payload, 'blockId') || event.blockId)) {
    return true;
  }

  return !!searchTurnId && currentTask.turnId !== searchTurnId;
}

function upsertArtifact(
  artifacts: Record<string, IHieArtifactRecord>,
  artifact: IHieArtifactRecord
): Record<string, IHieArtifactRecord> {
  return {
    ...artifacts,
    [artifact.artifactId]: artifact
  };
}

function parseSourceBlockId(originTool?: string): string | undefined {
  if (!originTool) return undefined;
  if (originTool.startsWith('share-form:')) {
    return originTool.slice('share-form:'.length);
  }
  if (originTool.startsWith('block-recap:')) {
    return originTool.slice('block-recap:'.length);
  }
  return undefined;
}

function deriveArtifactFromBlockEvent(
  event: IHieEvent,
  previous?: IHieArtifactRecord
): IHieArtifactRecord | undefined {
  const blockId = pickString(event.payload, 'blockId') || event.blockId;
  const blockType = pickBlockType(event.payload, 'blockType') || event.blockType;
  const blockTitle = pickString(event.payload, 'blockTitle');
  const originTool = pickString(event.payload, 'originTool');

  if (!blockId || !blockType || NON_DERIVED_BLOCK_ARTIFACT_TYPES.has(blockType)) {
    return undefined;
  }

  const isRecapArtifact = originTool?.startsWith('block-recap:') === true;
  const sourceBlockId = previous?.sourceBlockId
    || pickString(event.payload, 'sourceBlockId')
    || parseSourceBlockId(originTool);
  const sourceArtifactId = previous?.sourceArtifactId || pickString(event.payload, 'sourceArtifactId');
  const sourceTaskKind = previous?.sourceTaskKind
    || pickString(event.payload, 'sourceTaskKind')
    || (isRecapArtifact ? 'recap' : undefined);
  const sourceTurnId = previous?.sourceTurnId || pickString(event.payload, 'sourceTurnId') || event.turnId;
  const sourceRootTurnId = previous?.sourceRootTurnId || pickString(event.payload, 'sourceRootTurnId') || event.rootTurnId || sourceTurnId;
  const sourceParentTurnId = previous?.sourceParentTurnId || pickString(event.payload, 'sourceParentTurnId') || event.parentTurnId;

  const resolvedArtifactKind = resolveArtifactKindFromBlockContext(blockType, {
    title: blockTitle,
    sourceTaskKind,
    originTool
  });
  const targetContext = mergeMcpTargetContexts(
    previous?.targetContext,
    deriveTargetContextFromEvent(event, pickSelectedItems(event.payload, 'selectedItems'))
  );

  return {
    artifactId: blockId,
    artifactKind: previous?.artifactKind === 'summary'
      || previous?.artifactKind === 'preview'
      || previous?.artifactKind === 'lookup'
      || previous?.artifactKind === 'error'
      || previous?.artifactKind === 'recap'
      || previous?.artifactKind === 'form'
      || previous?.artifactKind === 'share'
      || previous?.artifactKind === 'generic'
      ? previous.artifactKind
      : resolvedArtifactKind,
    sourceBlockId,
    sourceArtifactId,
    sourceTaskKind,
    sourceEventName: previous?.sourceEventName || pickString(event.payload, 'sourceEventName'),
    correlationId: previous?.correlationId || pickString(event.payload, 'sourceCorrelationId') || event.correlationId,
    sourceTurnId,
    sourceRootTurnId,
    sourceParentTurnId,
    blockId,
    blockType,
    title: blockTitle || previous?.title,
    targetContext,
    status: previous?.status === 'loading' && event.eventName === 'block.created'
      ? previous.status
      : (isRecapArtifact && event.eventName === 'block.created' ? 'loading' : 'ready'),
    updatedAt: event.timestamp
  };
}

function deriveArtifactFromExplicitResultEvent(event: IHieEvent): IHieArtifactRecord | undefined {
  const blockId = pickString(event.payload, 'blockId') || event.blockId;
  const blockType = pickBlockType(event.payload, 'blockType') || event.blockType;
  if (!blockId || !blockType) {
    return undefined;
  }

  const explicitArtifactKind = pickString(event.payload, 'artifactKind');
  const sourceTaskKind = pickString(event.payload, 'sourceTaskKind');
  const blockTitle = pickString(event.payload, 'blockTitle') || pickString(event.payload, 'derivedBlockTitle');
  const resolvedArtifactKind = isHieArtifactKind(explicitArtifactKind)
    ? explicitArtifactKind
    : resolveArtifactKindFromBlockContext(blockType, {
      title: blockTitle,
      sourceTaskKind
    });
  const targetContext = deriveTargetContextFromEvent(event, pickSelectedItems(event.payload, 'selectedItems'));

  return {
    artifactId: blockId,
    artifactKind: resolvedArtifactKind,
    sourceBlockId: pickString(event.payload, 'sourceBlockId'),
    sourceArtifactId: pickString(event.payload, 'sourceArtifactId'),
    sourceTaskKind,
    sourceEventName: pickString(event.payload, 'sourceEventName'),
    correlationId: pickString(event.payload, 'sourceCorrelationId') || event.correlationId,
    sourceTurnId: pickString(event.payload, 'sourceTurnId') || event.turnId,
    sourceRootTurnId: pickString(event.payload, 'sourceRootTurnId') || event.rootTurnId || pickString(event.payload, 'sourceTurnId') || event.turnId,
    sourceParentTurnId: pickString(event.payload, 'sourceParentTurnId') || event.parentTurnId,
    blockId,
    blockType,
    title: blockTitle,
    targetContext,
    status: 'ready',
    updatedAt: event.timestamp
  };
}

function reduceShellState(shellState: IHieShellState, event: IHieEvent): IHieShellState {
  const payload = event.payload;

  switch (event.eventName) {
    case 'shell.logs.toggled':
      return {
        ...shellState,
        isLogsOpen: pickBoolean(payload, 'isOpen'),
        updatedAt: event.timestamp
      };
    case 'shell.settings.toggled':
      return {
        ...shellState,
        isSettingsOpen: pickBoolean(payload, 'isOpen'),
        updatedAt: event.timestamp
      };
    case 'shell.app.visibility':
      return {
        ...shellState,
        isAppOpen: pickBoolean(payload, 'isOpen'),
        updatedAt: event.timestamp
      };
    default:
      return shellState;
  }
}

export function createInitialHieDerivedState(): IHieDerivedState {
  return {
    artifacts: {},
    shellState: {}
  };
}

export function reduceHieDerivedState(
  previousState: IHieDerivedState,
  event: IHieEvent
): IHieDerivedState {
  const nextState: IHieDerivedState = {
    ...previousState,
    artifacts: { ...previousState.artifacts },
    shellState: reduceShellState(previousState.shellState, event)
  };

  switch (event.eventName) {
    case 'block.created':
    case 'block.updated': {
      const blockId = pickString(event.payload, 'blockId') || event.blockId;
      const blockType = pickBlockType(event.payload, 'blockType') || event.blockType;
      const artifact = deriveArtifactFromBlockEvent(
        event,
        blockId ? nextState.artifacts[blockId] : undefined
      );
      if (artifact) {
        nextState.artifacts = upsertArtifact(nextState.artifacts, artifact);
        if (artifact.sourceTaskKind) {
          nextState.taskContext = buildTaskContext(event, artifact.sourceTaskKind, {
            sourceBlockId: artifact.sourceBlockId,
            sourceBlockType: pickBlockType(event.payload, 'sourceBlockType'),
            sourceBlockTitle: pickString(event.payload, 'sourceBlockTitle'),
            derivedBlockId: artifact.blockId,
            derivedBlockType: artifact.blockType,
            derivedBlockTitle: artifact.title,
            turnId: artifact.sourceTurnId,
            rootTurnId: artifact.sourceRootTurnId,
            parentTurnId: artifact.sourceParentTurnId,
            selectedItems: pickSelectedItems(event.payload, 'selectedItems')
              || nextState.taskContext?.selectedItems,
            targetContext: mergeMcpTargetContexts(artifact.targetContext, nextState.taskContext?.targetContext)
          });
        }
      } else if (blockId && blockType === 'search-results' && shouldPromoteSearchBlockToTask(event, nextState)) {
        nextState.taskContext = buildTaskContext(event, 'search', {
          sourceBlockId: blockId,
          sourceBlockType: blockType,
          sourceBlockTitle: pickString(event.payload, 'blockTitle'),
          derivedBlockId: undefined,
          derivedBlockType: undefined,
          derivedBlockTitle: undefined,
          selectedItems: undefined,
          targetContext: deriveTargetContextFromEvent(event)
        });
      }
      return nextState;
    }

    case 'block.removed': {
      const removedBlockId = pickString(event.payload, 'blockId') || event.blockId;
      if (!removedBlockId) {
        return nextState;
      }

      const previousArtifact = nextState.artifacts[removedBlockId];
      if (previousArtifact) {
        nextState.artifacts = upsertArtifact(nextState.artifacts, {
          ...previousArtifact,
          status: previousArtifact.status === 'submitted' ? previousArtifact.status : 'dismissed',
          updatedAt: event.timestamp
        });
      }

      if (nextState.taskContext?.kind === 'form' && nextState.taskContext.derivedBlockId === removedBlockId) {
        nextState.taskContext = {
          ...nextState.taskContext,
          formStatus: 'dismissed',
          updatedAt: event.timestamp
        };
      }

      return nextState;
    }

    case 'block.interaction.click-result':
    case 'block.interaction.click-folder':
    case 'block.interaction.click-file':
    case 'block.interaction.click-user':
    case 'block.interaction.click-list-row':
    case 'block.interaction.click-permission':
    case 'block.interaction.click-activity':
    case 'block.interaction.select':
    case 'block.interaction.click-item':
    case 'block.interaction.look':
    case 'block.interaction.summarize':
    case 'block.interaction.chat-about':
    case 'block.interaction.open-external':
    case 'block.interaction.navigate': {
      const interactionKind = resolveInteractionTaskKind(event.eventName);
      if (interactionKind) {
        nextState.taskContext = buildTaskContext(event, interactionKind, {
          derivedBlockId: undefined,
          derivedBlockType: undefined,
          derivedBlockTitle: undefined,
          targetContext: mergeMcpTargetContexts(
            nextState.taskContext?.targetContext,
            deriveTargetContextFromEvent(event, pickSelectedItems(event.payload, 'selectedItems'))
          )
        });
      }
      return nextState;
    }

    case 'task.focused':
      nextState.taskContext = buildTaskContext(event, 'focus', {
        derivedBlockId: undefined,
        derivedBlockType: undefined,
        derivedBlockTitle: undefined,
        targetContext: mergeMcpTargetContexts(
          nextState.taskContext?.targetContext,
          deriveTargetContextFromEvent(event, pickSelectedItems(event.payload, 'selectedItems'))
        )
      });
      return nextState;

    case 'task.selection.updated':
      nextState.taskContext = buildTaskContext(event, 'select', {
        derivedBlockId: undefined,
        derivedBlockType: undefined,
        derivedBlockTitle: undefined
      });
      return nextState;

    case 'task.recap.requested': {
      const recapBlockId = pickString(event.payload, 'derivedBlockId');
      nextState.taskContext = buildTaskContext(event, 'recap', {
        derivedBlockId: recapBlockId,
        derivedBlockType: 'info-card',
        derivedBlockTitle: pickString(event.payload, 'derivedBlockTitle')
      });
      if (recapBlockId) {
        nextState.artifacts = upsertArtifact(nextState.artifacts, {
          artifactId: recapBlockId,
          artifactKind: 'recap',
          sourceBlockId: nextState.taskContext.sourceBlockId,
          sourceEventName: nextState.taskContext.eventName,
          correlationId: nextState.taskContext.correlationId,
          sourceTurnId: nextState.taskContext.turnId,
          sourceRootTurnId: nextState.taskContext.rootTurnId,
          sourceParentTurnId: nextState.taskContext.parentTurnId,
          blockId: recapBlockId,
          blockType: 'info-card',
          title: nextState.taskContext.derivedBlockTitle,
          status: 'loading',
          updatedAt: event.timestamp
        });
      }
      return nextState;
    }

    case 'artifact.recap.ready':
    case 'artifact.recap.failed': {
      const status = event.eventName === 'artifact.recap.ready' ? 'ready' : 'error';
      const recapBlockId = pickString(event.payload, 'derivedBlockId') || event.blockId;
      nextState.taskContext = buildTaskContext(event, 'recap', {
        derivedBlockId: recapBlockId,
        derivedBlockType: 'info-card',
        derivedBlockTitle: pickString(event.payload, 'derivedBlockTitle')
      });
      if (recapBlockId) {
        nextState.artifacts = upsertArtifact(nextState.artifacts, {
          artifactId: recapBlockId,
          artifactKind: 'recap',
          sourceBlockId: nextState.taskContext.sourceBlockId,
          sourceEventName: nextState.taskContext.eventName,
          correlationId: nextState.taskContext.correlationId,
          sourceTurnId: nextState.taskContext.turnId,
          sourceRootTurnId: nextState.taskContext.rootTurnId,
          sourceParentTurnId: nextState.taskContext.parentTurnId,
          blockId: recapBlockId,
          blockType: 'info-card',
          title: nextState.taskContext.derivedBlockTitle,
          status,
          updatedAt: event.timestamp
        });
      }
      return nextState;
    }

    case 'artifact.result.ready': {
      const artifact = deriveArtifactFromExplicitResultEvent(event);
      if (!artifact) {
        return nextState;
      }

      nextState.artifacts = upsertArtifact(nextState.artifacts, artifact);
      if (artifact.sourceTaskKind) {
        nextState.taskContext = buildTaskContext(event, artifact.sourceTaskKind, {
          sourceBlockId: artifact.sourceBlockId,
          sourceBlockType: pickBlockType(event.payload, 'sourceBlockType'),
          sourceBlockTitle: pickString(event.payload, 'sourceBlockTitle'),
          derivedBlockId: artifact.blockId,
          derivedBlockType: artifact.blockType,
          derivedBlockTitle: artifact.title,
          turnId: artifact.sourceTurnId,
          rootTurnId: artifact.sourceRootTurnId,
          parentTurnId: artifact.sourceParentTurnId,
          selectedItems: pickSelectedItems(event.payload, 'selectedItems')
            || nextState.taskContext?.selectedItems
        });
      }
      return nextState;
    }

    case 'form.opened': {
      const formBlockId = pickString(event.payload, 'blockId') || event.blockId;
      const formPreset = pickString(event.payload, 'preset');
      const sourceArtifactId = pickString(event.payload, 'sourceArtifactId');
      const sourceTaskKind = pickString(event.payload, 'sourceTaskKind');
      const sourceTurnId = pickString(event.payload, 'sourceTurnId') || event.turnId;
      const sourceRootTurnId = pickString(event.payload, 'sourceRootTurnId') || event.rootTurnId || sourceTurnId;
      const sourceParentTurnId = pickString(event.payload, 'sourceParentTurnId') || event.parentTurnId;
      const taskContext = buildTaskContext(event, 'form', {
        derivedBlockId: formBlockId,
        derivedBlockType: 'form',
        derivedBlockTitle: pickString(event.payload, 'blockTitle'),
        formPreset,
        formStatus: 'opened',
        turnId: sourceTurnId,
        rootTurnId: sourceRootTurnId,
        parentTurnId: sourceParentTurnId,
        targetContext: mergeMcpTargetContexts(
          nextState.taskContext?.targetContext,
          deriveTargetContextFromEvent(event, pickSelectedItems(event.payload, 'selectedItems'))
        )
      });
      nextState.taskContext = taskContext;
      if (formBlockId) {
        nextState.artifacts = upsertArtifact(nextState.artifacts, {
          artifactId: formBlockId,
          artifactKind: formPreset && formPreset.startsWith('share-') ? 'share' : 'form',
          sourceBlockId: taskContext.sourceBlockId,
          sourceArtifactId,
          sourceTaskKind,
          sourceEventName: taskContext.eventName,
          correlationId: taskContext.correlationId,
          sourceTurnId: sourceTurnId || undefined,
          sourceRootTurnId: sourceRootTurnId || undefined,
          sourceParentTurnId: sourceParentTurnId || undefined,
          blockId: formBlockId,
          blockType: 'form',
          title: taskContext.derivedBlockTitle,
          preset: formPreset,
          targetContext: taskContext.targetContext,
          status: 'opened',
          updatedAt: event.timestamp
        });
      }
      return nextState;
    }

    case 'form.submitted':
    case 'form.dismissed':
    case 'form.cancelled':
    case 'form.execution.failed': {
      const formBlockId = pickString(event.payload, 'blockId') || event.blockId;
      const formPreset = pickString(event.payload, 'preset');
      const formStatus = event.eventName === 'form.submitted'
        ? 'submitted'
        : event.eventName === 'form.dismissed'
          ? 'dismissed'
          : event.eventName === 'form.execution.failed'
            ? 'failed'
            : 'cancelled';
      nextState.taskContext = buildTaskContext(event, 'form', {
        derivedBlockId: formBlockId,
        derivedBlockType: 'form',
        formPreset,
        formStatus,
        turnId: pickString(event.payload, 'sourceTurnId') || event.turnId,
        rootTurnId: pickString(event.payload, 'sourceRootTurnId') || event.rootTurnId || pickString(event.payload, 'sourceTurnId') || event.turnId,
        parentTurnId: pickString(event.payload, 'sourceParentTurnId') || event.parentTurnId,
        targetContext: mergeMcpTargetContexts(
          nextState.artifacts[formBlockId || '']?.targetContext,
          nextState.taskContext?.targetContext,
          deriveTargetContextFromEvent(event, pickSelectedItems(event.payload, 'selectedItems'))
        )
      });
      if (formBlockId) {
        const previous = nextState.artifacts[formBlockId];
        nextState.artifacts = upsertArtifact(nextState.artifacts, {
          artifactId: formBlockId,
          artifactKind: previous?.artifactKind || (formPreset && formPreset.startsWith('share-') ? 'share' : 'form'),
          sourceBlockId: nextState.taskContext.sourceBlockId,
          sourceArtifactId: previous?.sourceArtifactId || pickString(event.payload, 'sourceArtifactId'),
          sourceTaskKind: previous?.sourceTaskKind || pickString(event.payload, 'sourceTaskKind'),
          sourceEventName: previous?.sourceEventName || pickString(event.payload, 'sourceEventName') || nextState.taskContext.eventName,
          correlationId: previous?.correlationId || pickString(event.payload, 'sourceCorrelationId') || nextState.taskContext.correlationId,
          sourceTurnId: previous?.sourceTurnId || pickString(event.payload, 'sourceTurnId') || nextState.taskContext.turnId,
          sourceRootTurnId: previous?.sourceRootTurnId || pickString(event.payload, 'sourceRootTurnId') || nextState.taskContext.rootTurnId,
          sourceParentTurnId: previous?.sourceParentTurnId || pickString(event.payload, 'sourceParentTurnId') || nextState.taskContext.parentTurnId,
          blockId: formBlockId,
          blockType: 'form',
          title: previous?.title || nextState.taskContext.derivedBlockTitle,
          preset: formPreset,
          targetContext: mergeMcpTargetContexts(previous?.targetContext, nextState.taskContext.targetContext),
          status: formStatus === 'failed' ? 'error' : formStatus,
          updatedAt: event.timestamp
        });
      }
      return nextState;
    }

    case 'thread.started':
    case 'thread.reset': {
      const preview = pickString(event.payload, 'userTextPreview');
      if (preview) {
        nextState.userTextPreview = preview;
      }
      return nextState;
    }

    case 'thread.continued': {
      const preview = pickString(event.payload, 'userTextPreview');
      if (preview) {
        nextState.userTextPreview = preview;
      }
      return nextState;
    }

    default:
      return nextState;
  }
}
