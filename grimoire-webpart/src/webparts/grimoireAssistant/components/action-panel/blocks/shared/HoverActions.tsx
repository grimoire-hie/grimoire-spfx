/**
 * HoverActions — Shared hover-reveal action buttons for block rows.
 *
 * Deprecated: header-level actions are now the primary interaction model.
 * This module remains to keep compatibility with existing block renderers.
 */

import * as React from 'react';
import type { BlockInteractionAction } from '../../../../services/hie/HIETypes';
import type { BlockType } from '../../../../models/IBlock';

// ─── Action Definitions ─────────────────────────────────────────

type HoverActionType = Extract<BlockInteractionAction, 'look' | 'summarize' | 'chat-about'>;

export interface IHoverAction {
  action: HoverActionType;
  icon: string;
  label: string;
}

export const LOOK_ACTION: IHoverAction = { action: 'look', icon: 'View', label: 'Focus' };
export const SUMMARIZE_ACTION: IHoverAction = { action: 'summarize', icon: 'TextDocument', label: 'Summarize' };
export const CHAT_ACTION: IHoverAction = { action: 'chat-about', icon: 'Chat', label: 'Chat' };

/** Pre-built action arrays to avoid per-render allocation */
export const ACTIONS_LSC: IHoverAction[] = [LOOK_ACTION, SUMMARIZE_ACTION, CHAT_ACTION];
export const ACTIONS_LOOK_ONLY: IHoverAction[] = [LOOK_ACTION];
export const ACTIONS_LC: IHoverAction[] = [LOOK_ACTION, CHAT_ACTION];
export const ACTIONS_SC: IHoverAction[] = [SUMMARIZE_ACTION, CHAT_ACTION];

export interface IRenderHoverActionsOptions {
  /** Keep actions visible even when row hover styles normally hide them. */
  alwaysVisible?: boolean;
}

// ─── CSS prefix helper ──────────────────────────────────────────

/** Returns class names for a given hover prefix. Single source of truth. */
export function hoverClasses(prefix: string): { row: string; actions: string } {
  return { row: `grim-${prefix}-row`, actions: `grim-${prefix}-actions` };
}

// ─── Style injection ────────────────────────────────────────────

/**
 * Inject CSS rules for hover-reveal pattern: actions hidden by default,
 * shown on row hover. Call once per prefix in a useEffect.
 * Uses document.getElementById as the idempotency guard (same pattern as
 * ActionPanel.injectStyleOnce).
 */
export function injectHoverStyles(prefix: string): void {
  if (typeof document === 'undefined') return;
  const id = `grim-hover-${prefix}`;
  if (document.getElementById(id)) return;

  const css = [
    `.grim-${prefix}-row .grim-${prefix}-actions { opacity: 0; pointer-events: none; transition: opacity 0.15s ease; }`,
    `.grim-${prefix}-row:hover .grim-${prefix}-actions { opacity: 1; pointer-events: auto; }`
  ].join('\n');

  const style = document.createElement('style');
  style.id = id;
  style.textContent = css;
  document.head.appendChild(style);
}

// ─── Render function ────────────────────────────────────────────

/**
 * Render nothing: row-level hover actions have been replaced by header actions.
 */
export function renderHoverActions(
  _actions: IHoverAction[],
  _blockId: string | undefined,
  _blockType: BlockType,
  _payload: Record<string, unknown>,
  _cssClass: string,
  _options?: IRenderHoverActionsOptions
): React.ReactNode {
  return null;
}
