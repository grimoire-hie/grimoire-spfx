/**
 * HIETypes — Type definitions for the Hybrid Interaction Engine.
 */

import { BlockType } from '../../models/IBlock';
import { Expression } from '../avatar/ExpressionEngine';
import type { IMcpTargetContext } from '../mcp/McpTargetContext';

// ─── Configuration ──────────────────────────────────────────────

export interface IHIEConfig {
  contextInjectionEnabled: boolean;
  interactionNotificationsEnabled: boolean;
  dataExpressionsEnabled: boolean;
  flowOrchestrationEnabled: boolean;
  /** Enable voice feedback after async tool completion */
  asyncToolResponseEnabled: boolean;
  /** Debounce ms for batching rapid block updates before context injection */
  contextDebounceMs: number;
  /** Max character length for a single context injection message */
  maxContextLength: number;
}

export interface IHIEContextPolicy {
  maxChars: number;
  debounceMs: number;
  triggerRules?: {
    visual?: boolean;
    interaction?: boolean;
    flow?: boolean;
    toolError?: boolean;
    toolCompletion?: boolean;
  };
}

export const DEFAULT_HIE_CONFIG: IHIEConfig = {
  contextInjectionEnabled: true,
  interactionNotificationsEnabled: true,
  dataExpressionsEnabled: true,
  flowOrchestrationEnabled: true,
  asyncToolResponseEnabled: true,
  contextDebounceMs: 800,
  maxContextLength: 1500
};

// ─── Block Tracking ─────────────────────────────────────────────

export type TrackedBlockState = 'loading' | 'ready' | 'acknowledged' | 'interacted' | 'dismissed';

export interface ITrackedBlock {
  id: string;
  type: BlockType;
  title: string;
  originTool?: string;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
  summary: string;
  itemCount: number;
  /** Numbered item references for positional resolution ("the second one") */
  references: IBlockReference[];
  state: TrackedBlockState;
  contextInjected: boolean;
  createdAt: number;
  updatedAt: number;
}

// ─── Generic HIE Events ────────────────────────────────────────

export type HieExposureMode = 'store-only' | 'silent-context' | 'response-triggering';
export type HieEventRelevance = 'background' | 'contextual' | 'foreground';
export type HieEventSource =
  | 'block-ui'
  | 'hover-action'
  | 'form'
  | 'system'
  | 'action-panel'
  | 'avatar-panel'
  | 'app-layout'
  | 'log-sidebar'
  | 'tool-runtime'
  | 'hie';
export type HieSurface =
  | 'action-panel'
  | 'avatar-panel'
  | 'settings'
  | 'logs'
  | 'app-shell'
  | 'tool-runtime'
  | 'unknown';

export interface IHieExposurePolicy {
  mode: HieExposureMode;
  relevance?: HieEventRelevance;
}

export interface IHieTurnLineage {
  turnId: string;
  rootTurnId: string;
  parentTurnId?: string;
}

export type HieTurnStartMode = 'inherit' | 'new-root' | 'auto';

export interface IHieTurnStartOptions {
  turnId?: string;
  mode?: HieTurnStartMode;
  text?: string;
  reason?: string;
}

export interface IHieSourceContext {
  sourceBlockId?: string;
  sourceBlockType?: BlockType;
  sourceBlockTitle?: string;
  sourceArtifactId?: string;
  sourceTaskKind?: string;
  sourceEventName?: string;
  correlationId?: string;
  sourceTurnId?: string;
  sourceRootTurnId?: string;
  sourceParentTurnId?: string;
  selectedItems?: IHieSelectedItem[];
  targetContext?: IMcpTargetContext;
}

export interface IHieEvent {
  eventId: string;
  eventName: string;
  source: HieEventSource;
  surface: HieSurface;
  correlationId: string;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
  timestamp: number;
  payload: Record<string, unknown>;
  exposurePolicy: IHieExposurePolicy;
  blockId?: string;
  blockType?: BlockType;
}

export interface IHieSelectedItem {
  index?: number;
  title?: string;
  url?: string;
  kind?: string;
  itemType?: string;
  targetContext?: IMcpTargetContext;
}

export interface IHieTaskContext {
  kind: string;
  eventName: string;
  correlationId?: string;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
  sourceBlockId?: string;
  sourceBlockType?: BlockType;
  sourceBlockTitle?: string;
  derivedBlockId?: string;
  derivedBlockType?: BlockType;
  derivedBlockTitle?: string;
  formPreset?: string;
  formStatus?: string;
  selectedItems?: IHieSelectedItem[];
  targetContext?: IMcpTargetContext;
  updatedAt: number;
}

export type HieArtifactKind =
  | 'block'
  | 'summary'
  | 'preview'
  | 'lookup'
  | 'error'
  | 'recap'
  | 'form'
  | 'share'
  | 'generic';

export interface IHieArtifactRecord {
  artifactId: string;
  artifactKind: HieArtifactKind;
  sourceBlockId?: string;
  sourceArtifactId?: string;
  sourceTaskKind?: string;
  sourceEventName?: string;
  correlationId?: string;
  sourceTurnId?: string;
  sourceRootTurnId?: string;
  sourceParentTurnId?: string;
  blockId?: string;
  blockType?: BlockType;
  title?: string;
  preset?: string;
  targetContext?: IMcpTargetContext;
  status: 'opened' | 'requested' | 'loading' | 'ready' | 'error' | 'submitted' | 'cancelled' | 'dismissed';
  updatedAt: number;
}

export interface IHieShellState {
  isLogsOpen?: boolean;
  isSettingsOpen?: boolean;
  isAppOpen?: boolean;
  updatedAt?: number;
}

export interface IHieDerivedState {
  taskContext?: IHieTaskContext;
  artifacts: Record<string, IHieArtifactRecord>;
  shellState: IHieShellState;
  userTextPreview?: string;
}

export const DEFAULT_HIE_EXPOSURE_POLICY: IHieExposurePolicy = {
  mode: 'response-triggering',
  relevance: 'foreground'
};

// ─── Block Interactions ─────────────────────────────────────────

export type BlockInteractionAction =
  | 'click-result'
  | 'click-folder'
  | 'click-file'
  | 'click-user'
  | 'click-list-row'
  | 'click-permission'
  | 'click-activity'
  | 'confirm'
  | 'cancel'
  | 'select'
  | 'dismiss'
  | 'open-external'
  | 'navigate'
  | 'retry'
  | 'click-item'
  | 'submit-form'
  | 'cancel-form'
  | 'look'
  | 'summarize'
  | 'chat-about';

export interface IBlockInteraction {
  blockId: string;
  blockType: BlockType;
  action: BlockInteractionAction;
  payload: Record<string, unknown>;
  timestamp: number;
  schemaId?: string;
  source?: HieEventSource;
  surface?: HieSurface;
  eventName?: string;
  exposurePolicy?: IHieExposurePolicy;
  correlationId?: string;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
}

// ─── Context Injection ──────────────────────────────────────────

export type ContextType = 'visual' | 'interaction' | 'flow';

export interface IContextMessage {
  contextType: ContextType;
  text: string;
  blockIds: string[];
  sentAt: number;
}

// ─── Flow Orchestration ─────────────────────────────────────────

export type FlowId = 'search-then-drill' | 'browse-then-open' | 'confirm-before-action' | 'select-then-act' | 'compose-then-submit';

export interface IFlowStep {
  name: string;
  description: string;
}

export interface IFlowDefinition {
  id: FlowId;
  name: string;
  steps: IFlowStep[];
  triggerTools: string[];
}

export interface IFlowInstance {
  definition: IFlowDefinition;
  currentStep: number;
  context: Record<string, unknown>;
  blockIds: string[];
  startedAt: number;
}

// ─── Data-Aware Expressions ─────────────────────────────────────

export interface IExpressionRule {
  id: string;
  trigger: string;
  expression: Expression;
  revertMs: number;
  priority: number;
}

// ─── Numbered References ────────────────────────────────────────

export interface IBlockReference {
  /** 1-based index within the block */
  index: number;
  /** Display title of the item */
  title: string;
  /** URL if applicable (search result, file, site) */
  url?: string;
  /** Item type hint (e.g. 'file', 'folder', 'person', 'site') */
  itemType?: string;
  /** Extra detail string (e.g. author, department, file type) */
  detail?: string;
}

// ─── Verbosity ──────────────────────────────────────────────

export type VerbosityLevel = 'minimal' | 'brief' | 'normal' | 'detailed';

// ─── Destructive Tools ──────────────────────────────────────────

export const DESTRUCTIVE_TOOL_NAMES: ReadonlySet<string> = new Set([
  // ODSP
  'deleteFileOrFolder',
  // Lists
  'deleteListItem',
  'deleteListColumn',
  // Mail
  'DeleteMessage',
  'DeleteAttachment',
  // Calendar
  'DeleteEventById',
  'CancelEvent',
  // Teams
  'DeleteChat',
  'DeleteChatMessage'
]);
