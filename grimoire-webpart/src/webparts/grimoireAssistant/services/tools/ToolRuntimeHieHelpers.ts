import type { BlockType, IBlock } from '../../models/IBlock';
import { createCorrelationId } from '../hie/HAEContracts';
import { resolveArtifactKindFromBlockContext } from '../hie/HieArtifactKindResolver';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import type { IHieSourceContext } from '../hie/HIETypes';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';

const EXPLICIT_RUNTIME_ARTIFACT_BLOCK_TYPES: ReadonlySet<BlockType> = new Set<BlockType>([
  'info-card',
  'markdown',
  'error',
  'file-preview',
  'site-info',
  'user-card',
  'list-items',
  'permissions-view',
  'activity-feed',
  'chart',
  'document-library'
]);

function shouldEmitExplicitArtifactResult(
  block: IBlock,
  sourceContext?: IHieSourceContext
): boolean {
  if (!sourceContext) {
    return false;
  }

  const hasSourceLink = !!(
    sourceContext.sourceBlockId
    || sourceContext.sourceArtifactId
    || sourceContext.sourceTaskKind
    || sourceContext.sourceEventName
  );
  if (!hasSourceLink) {
    return false;
  }

  return EXPLICIT_RUNTIME_ARTIFACT_BLOCK_TYPES.has(block.type);
}

function emitArtifactResultReady(block: IBlock, sourceContext?: IHieSourceContext): void {
  if (!shouldEmitExplicitArtifactResult(block, sourceContext)) {
    return;
  }

  const artifactKind = resolveArtifactKindFromBlockContext(block.type, {
    title: block.title,
    sourceTaskKind: sourceContext?.sourceTaskKind,
    originTool: block.originTool
  });

  hybridInteractionEngine.emitEvent({
    eventName: 'artifact.result.ready',
    source: 'tool-runtime',
    surface: 'tool-runtime',
    correlationId: createCorrelationId('artifact'),
    payload: {
      blockId: block.id,
      blockType: block.type,
      blockTitle: block.title,
      artifactKind,
      sourceBlockId: sourceContext?.sourceBlockId,
      sourceBlockType: sourceContext?.sourceBlockType,
      sourceBlockTitle: sourceContext?.sourceBlockTitle,
      sourceArtifactId: sourceContext?.sourceArtifactId,
      sourceTaskKind: sourceContext?.sourceTaskKind,
      sourceEventName: sourceContext?.sourceEventName,
      sourceCorrelationId: sourceContext?.correlationId,
      sourceTurnId: sourceContext?.sourceTurnId,
      sourceRootTurnId: sourceContext?.sourceRootTurnId,
      sourceParentTurnId: sourceContext?.sourceParentTurnId,
      selectedItems: sourceContext?.selectedItems
    },
    exposurePolicy: { mode: 'store-only', relevance: 'contextual' },
    turnId: sourceContext?.sourceTurnId,
    rootTurnId: sourceContext?.sourceRootTurnId,
    parentTurnId: sourceContext?.sourceParentTurnId,
    blockId: block.id,
    blockType: block.type
  });
}

export function trackCreatedBlock(
  store: Pick<IFunctionCallStore, 'pushBlock'>,
  block: IBlock,
  deps: Pick<IToolRuntimeHandlerDeps, 'sourceContext'>
): void {
  store.pushBlock(block);
  hybridInteractionEngine.onBlockCreated(block, deps.sourceContext);
  emitArtifactResultReady(block, deps.sourceContext);
}

export function trackUpdatedBlock<TUpdates extends Record<string, unknown>>(
  store: { updateBlock: (blockId: string, updates: TUpdates) => void },
  blockId: string,
  updates: TUpdates,
  nextBlock: IBlock,
  deps: Pick<IToolRuntimeHandlerDeps, 'sourceContext'>
): void {
  store.updateBlock(blockId, updates);
  hybridInteractionEngine.onBlockUpdated(blockId, nextBlock, deps.sourceContext);
  emitArtifactResultReady(nextBlock, deps.sourceContext);
}

export function trackToolCompletion(
  toolName: string,
  blockId: string,
  success: boolean,
  itemCount: number,
  deps: Pick<IToolRuntimeHandlerDeps, 'sourceContext'>
): void {
  hybridInteractionEngine.onToolComplete(toolName, blockId, success, itemCount, deps.sourceContext);
}
