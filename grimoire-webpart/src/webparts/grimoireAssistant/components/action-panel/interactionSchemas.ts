/**
 * interactionSchemas
 * Declarative block interaction mapping + unified emitter adapter.
 */

import type { BlockType } from '../../models/IBlock';
import type {
  BlockInteractionAction,
  HieSurface,
  IBlockInteraction,
  IHieExposurePolicy
} from '../../services/hie/HIETypes';
import { DEFAULT_HIE_EXPOSURE_POLICY } from '../../services/hie/HIETypes';
import { hybridInteractionEngine } from '../../services/hie/HybridInteractionEngine';
import { createCorrelationId } from '../../services/hie/HAEContracts';

const RESPONSE_TRIGGERING_POLICY: IHieExposurePolicy = {
  mode: 'response-triggering',
  relevance: 'foreground'
};

export interface IInteractionSchema {
  id: string;
  blockType: BlockType;
  action: BlockInteractionAction;
  eventName: string;
  exposurePolicy: IHieExposurePolicy;
  source: 'block-ui' | 'hover-action' | 'form' | 'system';
  surface: HieSurface;
  buildPayload?: (payload: Record<string, unknown>) => Record<string, unknown>;
}

const DEFAULT_PAYLOAD = (payload: Record<string, unknown>): Record<string, unknown> => payload;

const SCHEMAS: IInteractionSchema[] = [
  { id: 'search-results.click-result', blockType: 'search-results', action: 'click-result', eventName: 'block.interaction.click-result', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'document-library.click-folder', blockType: 'document-library', action: 'click-folder', eventName: 'block.interaction.click-folder', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'document-library.click-file', blockType: 'document-library', action: 'click-file', eventName: 'block.interaction.click-file', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'file-preview.open-external', blockType: 'file-preview', action: 'open-external', eventName: 'block.interaction.open-external', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'site-info.open-external', blockType: 'site-info', action: 'open-external', eventName: 'block.interaction.open-external', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'user-card.click-user', blockType: 'user-card', action: 'click-user', eventName: 'block.interaction.click-user', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'list-items.click-list-row', blockType: 'list-items', action: 'click-list-row', eventName: 'block.interaction.click-list-row', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'permissions-view.click-permission', blockType: 'permissions-view', action: 'click-permission', eventName: 'block.interaction.click-permission', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'activity-feed.click-activity', blockType: 'activity-feed', action: 'click-activity', eventName: 'block.interaction.click-activity', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'confirmation.confirm', blockType: 'confirmation-dialog', action: 'confirm', eventName: 'block.interaction.confirm', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'confirmation.cancel', blockType: 'confirmation-dialog', action: 'cancel', eventName: 'block.interaction.cancel', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'selection.select', blockType: 'selection-list', action: 'select', eventName: 'block.interaction.select', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'error.retry', blockType: 'error', action: 'retry', eventName: 'block.interaction.retry', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'block-ui', surface: 'action-panel' },
  { id: 'form.submit', blockType: 'form', action: 'submit-form', eventName: 'block.interaction.submit-form', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'form', surface: 'action-panel' },
  { id: 'form.cancel', blockType: 'form', action: 'cancel-form', eventName: 'block.interaction.cancel-form', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'form', surface: 'action-panel' },
  { id: 'hover.look', blockType: 'search-results', action: 'look', eventName: 'block.interaction.look', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'hover-action', surface: 'action-panel', buildPayload: DEFAULT_PAYLOAD },
  { id: 'hover.summarize', blockType: 'search-results', action: 'summarize', eventName: 'block.interaction.summarize', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'hover-action', surface: 'action-panel', buildPayload: DEFAULT_PAYLOAD },
  { id: 'hover.chat-about', blockType: 'search-results', action: 'chat-about', eventName: 'block.interaction.chat-about', exposurePolicy: RESPONSE_TRIGGERING_POLICY, source: 'hover-action', surface: 'action-panel', buildPayload: DEFAULT_PAYLOAD }
];

const SCHEMA_BY_ID: Record<string, IInteractionSchema> = {};
const SCHEMA_BY_BLOCK_ACTION: Record<string, IInteractionSchema> = {};

for (let i = 0; i < SCHEMAS.length; i++) {
  const s = SCHEMAS[i];
  SCHEMA_BY_ID[s.id] = s;
  SCHEMA_BY_BLOCK_ACTION[`${s.blockType}:${s.action}`] = s;
}

export function resolveInteractionSchemaId(
  blockType: BlockType,
  action: BlockInteractionAction
): string | undefined {
  const exact = SCHEMA_BY_BLOCK_ACTION[`${blockType}:${action}`];
  if (exact) return exact.id;

  // Hover actions are generic and intentionally apply to all block types.
  if (action === 'look' || action === 'summarize' || action === 'chat-about') {
    return `hover.${action}`;
  }
  return undefined;
}

export function getInteractionSchema(
  blockType: BlockType,
  action: BlockInteractionAction
): IInteractionSchema | undefined {
  const id = resolveInteractionSchemaId(blockType, action);
  if (!id) return undefined;
  return SCHEMA_BY_ID[id] || SCHEMA_BY_BLOCK_ACTION[`${blockType}:${action}`];
}

export interface IEmitInteractionParams {
  blockId?: string;
  blockType: BlockType;
  action: BlockInteractionAction;
  payload: Record<string, unknown>;
  schemaId?: string;
  timestamp?: number;
}

/**
 * Unified adapter used by block components to emit interactions.
 * Returns false when the interaction cannot be emitted (missing blockId).
 */
export function emitBlockInteraction(params: IEmitInteractionParams): boolean {
  const { blockId, blockType, action, payload, schemaId, timestamp } = params;
  if (!blockId) return false;

  const schema = schemaId
    ? SCHEMA_BY_ID[schemaId] || getInteractionSchema(blockType, action)
    : getInteractionSchema(blockType, action);

  const normalizedPayload = schema?.buildPayload ? schema.buildPayload(payload) : payload;
  const interaction: IBlockInteraction = {
    blockId,
    blockType,
    action,
    payload: normalizedPayload,
    timestamp: timestamp || Date.now(),
    eventName: schema?.eventName || `block.interaction.${action}`,
    schemaId: schema?.id || schemaId,
    source: schema?.source || 'block-ui',
    surface: schema?.surface || 'action-panel',
    exposurePolicy: schema?.exposurePolicy || DEFAULT_HIE_EXPOSURE_POLICY,
    correlationId: createCorrelationId('iact')
  };

  hybridInteractionEngine.onBlockInteraction(interaction);
  return true;
}

export function getInteractionSchemas(): IInteractionSchema[] {
  return SCHEMAS.slice();
}
