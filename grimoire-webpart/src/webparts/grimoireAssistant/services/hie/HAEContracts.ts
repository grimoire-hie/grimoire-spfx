/**
 * HAEContracts
 * Formal contract types for Hybrid Architecture Engine state, interactions,
 * tool lifecycle, and context injections.
 */

import type { BlockType } from '../../models/IBlock';
import type {
  BlockInteractionAction,
  ContextType,
  HieEventSource,
  HieExposureMode,
  HieSurface,
  IBlockReference,
  TrackedBlockState
} from './HIETypes';

export type CorrelationId = string;

export interface IVisualBlockSnapshot {
  blockId: string;
  blockType: BlockType;
  title: string;
  state: TrackedBlockState;
  summary: string;
  itemCount: number;
  references: IBlockReference[];
  createdAt: number;
  updatedAt: number;
}

export interface IVisualStateSnapshot {
  snapshotId: string;
  correlationId: CorrelationId;
  createdAt: number;
  summary: string;
  flowSummary: string;
  hasDuplicateTypes: boolean;
  blocks: IVisualBlockSnapshot[];
  referencesNewestFirst: Array<{
    blockId: string;
    blockType: BlockType;
    references: IBlockReference[];
  }>;
}

export interface IInteractionEnvelope {
  envelopeId: string;
  correlationId: CorrelationId;
  createdAt: number;
  source: HieEventSource;
  blockId: string;
  blockType: BlockType;
  action: BlockInteractionAction;
  triggerResponse: boolean;
  schemaId?: string;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
  payload: Record<string, unknown>;
}

export interface IHieEventEnvelope {
  envelopeId: string;
  eventId: string;
  correlationId: CorrelationId;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
  createdAt: number;
  eventName: string;
  source: HieEventSource;
  surface: HieSurface;
  exposureMode: HieExposureMode;
  blockId?: string;
  blockType?: BlockType;
  payload: Record<string, unknown>;
}

export interface IToolExecutionEnvelope {
  envelopeId: string;
  correlationId: CorrelationId;
  createdAt: number;
  toolName: string;
  phase: 'dispatch' | 'complete' | 'error';
  awaitAsync: boolean;
  blockId?: string;
  success?: boolean;
  itemCount?: number;
  durationMs?: number;
  args?: Record<string, unknown>;
  message?: string;
}

export interface IContextEnvelope {
  envelopeId: string;
  correlationId: CorrelationId;
  createdAt: number;
  contextType: ContextType;
  triggerResponse: boolean;
  blockIds: string[];
  text: string;
  source: 'hie';
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
}

/**
 * Lightweight correlation ID generator suitable for client-side tracing.
 */
export function createCorrelationId(prefix: string = 'hae'): CorrelationId {
  const rand = Math.floor(Math.random() * 0xfffff).toString(36);
  return `${prefix}-${Date.now().toString(36)}-${rand}`;
}
