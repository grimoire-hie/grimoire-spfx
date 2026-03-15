/**
 * HybridInteractionEngine — Core singleton orchestrating visual context injection,
 * block interaction handling, data-aware expressions, and multi-step flow detection.
 *
 * This is the coordination engine that closes the three gaps between
 * Grimoire's voice/text LLM and its traditional UI controls:
 *   Gap 1: LLM never sees block contents → ContextInjector
 *   Gap 2: Block clicks don't feed back to LLM → onBlockInteraction
 *   Gap 3: No multi-step flow coordination → FlowOrchestrator
 */

import { IBlock } from '../../models/IBlock';
import { logService } from '../logging/LogService';
import { BlockTracker } from './BlockTracker';
import { ContextInjector } from './ContextInjector';
import type { SendContextFn } from './ContextInjector';
import { DataExpressionDirector, SetExpressionFn, IExpressionTriggerSnapshot } from './DataExpressionDirector';
import { FlowOrchestrator } from './FlowOrchestrator';
import { HieContextProjector } from './HieContextProjector';
import type { IHiePromptMessage } from './HiePromptProtocol';
import { resolveLatestArtifactContext } from './HieArtifactLinkage';
import { createInitialHieDerivedState, reduceHieDerivedState } from './HieStateReducer';
import { VerbosityDirector } from './VerbosityDirector';
import {
  DEFAULT_HIE_CONFIG,
  DEFAULT_HIE_EXPOSURE_POLICY,
  HieTurnStartMode,
  IBlockInteraction,
  IContextMessage,
  IHIEConfig,
  IHIEContextPolicy,
  IHieArtifactRecord,
  IHieDerivedState,
  IHieEvent,
  IHieShellState,
  IHieSourceContext,
  IHieTaskContext,
  IHieTurnLineage,
  IHieTurnStartOptions,
  VerbosityLevel
} from './HIETypes';
import type { NanoService } from '../nano/NanoService';
import {
  createCorrelationId,
  IHieEventEnvelope,
  IInteractionEnvelope,
  IToolExecutionEnvelope,
  IVisualStateSnapshot
} from './HAEContracts';
import {
  deriveMcpTargetContextFromUnknown,
  mergeMcpTargetContexts
} from '../mcp/McpTargetContext';

export type SetGazeFn = (target: 'none' | 'action-panel') => void;
export interface IAsyncToolCompletionFeedback {
  toolName: string;
  blockId: string;
  itemCount: number;
  summary?: string;
}

export interface IHIEInitOptions {
  config?: Partial<IHIEConfig>;
  setGazeFn?: SetGazeFn;
  nanoService?: NanoService;
  sendContextMessage?: SendContextFn;
  contextPolicy?: IHIEContextPolicy;
  onAsyncToolCompletion?: (feedback: IAsyncToolCompletionFeedback) => void;
}

/** Number of LLM response turns between automatic visual state re-injections */
const GROUNDING_INTERVAL_TURNS: number = 5;

/** Tools that run asynchronously (return placeholder text immediately, push UI blocks later) */
const ASYNC_TOOLS: ReadonlySet<string> = new Set([
  'research_public_web',
  'search_sharepoint', 'search_people', 'search_sites',
  'browse_document_library', 'show_file_details', 'show_site_info', 'show_list_items',
  'show_compose_form',
  'connect_mcp_server', 'call_mcp_tool', 'use_m365_capability',
  'get_my_profile', 'get_recent_documents', 'get_trending_documents',
  'recall_notes', 'search_emails', 'read_file_content'
]);

/**
 * Async tools where local voice feedback is preferred over another model turn.
 * These primarily push visual blocks and don't need delayed LLM narration.
 */
const LOCAL_COMPLETION_FEEDBACK_TOOLS: ReadonlySet<string> = new Set([
  'research_public_web',
  'search_sharepoint',
  'search_people',
  'search_sites',
  'search_emails',
  'browse_document_library',
  'show_compose_form',
  'show_file_details',
  'show_site_info',
  'show_list_items',
  'get_my_profile',
  'get_recent_documents',
  'get_trending_documents',
  'recall_notes'
]);

/** Tools that already inject explicit detailed error context via sendToolError() */
const EXPLICIT_TOOL_ERROR_CONTEXT_TOOLS: ReadonlySet<string> = new Set([
  'call_mcp_tool',
  'use_m365_capability'
]);
const HIE_EVENT_HISTORY_LIMIT = 250;
const HIE_PENDING_ACTION_LIMIT = 200;

export class HybridInteractionEngine {
  private tracker: BlockTracker;
  private injector: ContextInjector;
  private expressionDirector: DataExpressionDirector;
  private flowOrchestrator: FlowOrchestrator;
  private verbosityDirector: VerbosityDirector;
  private config: IHIEConfig;
  private initialized: boolean = false;
  /** Counts LLM response turns since last visual state grounding injection */
  private turnsSinceGrounding: number = 0;
  private setGaze: SetGazeFn = () => { /* no-op until initialized */ };
  private gazeRevertTimer: ReturnType<typeof setTimeout> | undefined;
  /** True when voice WebRTC path is active (not text-only) */
  private voicePathActive: boolean = false;
  /** Gate for tool error context injection from contextPolicy.triggerRules.toolError */
  private toolErrorInjectionEnabled: boolean = true;
  /** Optional local callback for low-latency async completion acknowledgments. */
  private onAsyncToolCompletion: ((feedback: IAsyncToolCompletionFeedback) => void) | undefined;
  private recentEvents: IHieEvent[] = [];
  private derivedState: IHieDerivedState = createInitialHieDerivedState();
  private contextProjector: HieContextProjector = new HieContextProjector();
  private pendingInitializationActions: Array<{ label: string; action: () => void }> = [];
  private currentTurn: IHieTurnLineage | undefined;
  private turnLineageById: Map<string, IHieTurnLineage> = new Map();

  constructor() {
    this.config = { ...DEFAULT_HIE_CONFIG };
    this.tracker = new BlockTracker();
    this.injector = new ContextInjector(this.config, this.tracker, () => { /* no-op until initialized */ });
    this.expressionDirector = new DataExpressionDirector(this.config, () => { /* no-op until initialized */ });
    this.flowOrchestrator = new FlowOrchestrator(this.config);
    this.verbosityDirector = new VerbosityDirector(this.config, this.tracker);
  }

  /**
   * Initialize the engine when a voice session connects.
   * Must be called with the store's setExpression function.
   */
  public initialize(setExpression: SetExpressionFn, options?: IHIEInitOptions): void {
    const { config, setGazeFn, nanoService, sendContextMessage, contextPolicy } = options || {};
    if (config) {
      Object.assign(this.config, config);
    }
    if (contextPolicy) {
      this.config.maxContextLength = contextPolicy.maxChars;
      this.config.contextDebounceMs = contextPolicy.debounceMs;
      if (contextPolicy.triggerRules) {
        if (contextPolicy.triggerRules.visual !== undefined) {
          this.config.contextInjectionEnabled = contextPolicy.triggerRules.visual;
        }
        if (contextPolicy.triggerRules.interaction !== undefined) {
          this.config.interactionNotificationsEnabled = contextPolicy.triggerRules.interaction;
        }
        if (contextPolicy.triggerRules.flow !== undefined) {
          this.config.flowOrchestrationEnabled = contextPolicy.triggerRules.flow;
        }
        if (contextPolicy.triggerRules.toolCompletion !== undefined) {
          this.config.asyncToolResponseEnabled = contextPolicy.triggerRules.toolCompletion;
        }
        this.toolErrorInjectionEnabled = contextPolicy.triggerRules.toolError !== false;
      } else {
        this.toolErrorInjectionEnabled = true;
      }
    } else {
      this.toolErrorInjectionEnabled = true;
    }
    this.tracker = new BlockTracker();
    const sendFn: SendContextFn = sendContextMessage || (() => { /* no-op fallback */ });
    this.injector = new ContextInjector(this.config, this.tracker, sendFn, nanoService);
    this.expressionDirector = new DataExpressionDirector(this.config, setExpression);
    this.flowOrchestrator = new FlowOrchestrator(this.config);
    this.verbosityDirector = new VerbosityDirector(this.config, this.tracker);
    if (setGazeFn) this.setGaze = setGazeFn;
    this.turnsSinceGrounding = 0;
    this.recentEvents = [];
    this.derivedState = createInitialHieDerivedState();
    this.initialized = true;
    this.onAsyncToolCompletion = options?.onAsyncToolCompletion;
    this.flushPendingInitializationActions();

    logService.info('system', `HIE initialized${nanoService ? ' (with Nano)' : ''}`);
  }

  /**
   * Access the block tracker for querying tracked block metadata.
   */
  public getBlockTracker(): BlockTracker {
    return this.tracker;
  }

  public getRecentEvents(): ReadonlyArray<IHieEvent> {
    return this.recentEvents;
  }

  public getCurrentTaskContext(): Readonly<IHieTaskContext> | undefined {
    return this.derivedState.taskContext;
  }

  public getCurrentArtifacts(): Readonly<Record<string, IHieArtifactRecord>> {
    return this.derivedState.artifacts;
  }

  public getShellState(): Readonly<IHieShellState> {
    return this.derivedState.shellState;
  }

  public setCurrentTurnId(turnId?: string): void {
    if (turnId) {
      this.currentTurn = this.turnLineageById.get(turnId) || this.rememberTurnLineage({
        turnId,
        rootTurnId: turnId
      });
    }
  }

  public getCurrentTurnId(): string | undefined {
    return this.currentTurn?.turnId;
  }

  public getCurrentTurnLineage(): Readonly<IHieTurnLineage> | undefined {
    return this.currentTurn || this.getTaskContextTurnLineage(this.derivedState.taskContext);
  }

  public getContextHistory(): ReadonlyArray<IContextMessage> {
    return this.injector.getHistory();
  }

  public getActiveFlowState(): { flowName: string; stepIndex: number; totalSteps: number } | undefined {
    const summary = this.flowOrchestrator.getActiveFlowSummary();
    if (!summary) {
      return undefined;
    }
    // Parse from "Flow Name: step X/Y (description)" format
    const match = summary.match(/^(.+?):\s*step\s+(\d+)\/(\d+)/);
    if (match) {
      return {
        flowName: match[1],
        stepIndex: Number(match[2]),
        totalSteps: Number(match[3])
      };
    }
    return { flowName: summary, stepIndex: 1, totalSteps: 1 };
  }

  public getVerbosityLevel(): VerbosityLevel {
    const hint = this.verbosityDirector.getVerbosityHint();
    return hint ? hint.level : 'normal';
  }

  public getLastExpressionTrigger(): Readonly<IExpressionTriggerSnapshot> | undefined {
    return this.expressionDirector.getLastTrigger();
  }

  public getTurnLineage(turnId?: string): Readonly<IHieTurnLineage> | undefined {
    if (!turnId) {
      return undefined;
    }
    return this.turnLineageById.get(turnId);
  }

  public beginUserTurn(turnIdOrOptions?: string | IHieTurnStartOptions): IHieTurnLineage {
    const options: IHieTurnStartOptions = typeof turnIdOrOptions === 'string'
      ? { turnId: turnIdOrOptions, mode: 'inherit' }
      : (turnIdOrOptions || {});
    const resolvedTurnId = options.turnId || createCorrelationId('turn');
    const previousTurn = this.getCurrentContextTurnLineage();
    const turnStartDecision = this.resolveTurnStartDecision(options, previousTurn);
    const lineage = this.rememberTurnLineage({
      turnId: resolvedTurnId,
      rootTurnId: turnStartDecision.mode === 'new-root'
        ? resolvedTurnId
        : (previousTurn?.rootTurnId || previousTurn?.turnId || resolvedTurnId),
      parentTurnId: turnStartDecision.mode === 'inherit' && previousTurn?.turnId && previousTurn.turnId !== resolvedTurnId
        ? previousTurn.turnId
        : undefined
    });
    this.currentTurn = lineage;
    this.emitThreadLifecycleEvent(previousTurn, lineage, turnStartDecision.mode, options.mode || 'inherit', turnStartDecision.reason, options.text);
    return lineage;
  }

  public captureCurrentSourceContext(): IHieSourceContext | undefined {
    const taskContext = this.derivedState.taskContext;
    const taskArtifact = taskContext?.derivedBlockId
      ? this.derivedState.artifacts[taskContext.derivedBlockId]
      : undefined;
    const latestArtifactContext = resolveLatestArtifactContext(this.derivedState.artifacts);
    const currentArtifact = taskArtifact || latestArtifactContext.currentArtifact;
    const currentTurn = this.getCurrentContextTurnLineage();
    const turnId = currentArtifact?.sourceTurnId || taskContext?.turnId || currentTurn?.turnId;
    const rootTurnId = currentArtifact?.sourceRootTurnId || taskContext?.rootTurnId || currentTurn?.rootTurnId || turnId;
    const parentTurnId = currentArtifact?.sourceParentTurnId || taskContext?.parentTurnId || currentTurn?.parentTurnId;

    if (!taskContext && !currentArtifact && !currentTurn) {
      return undefined;
    }

    return {
      sourceBlockId: currentArtifact?.sourceBlockId || taskContext?.sourceBlockId || currentArtifact?.blockId,
      sourceBlockType: taskContext?.sourceBlockType || currentArtifact?.blockType,
      sourceBlockTitle: taskContext?.sourceBlockTitle || currentArtifact?.title,
      sourceArtifactId: currentArtifact?.artifactId
        || (taskContext?.derivedBlockId && taskContext.derivedBlockType ? taskContext.derivedBlockId : undefined),
      sourceTaskKind: currentArtifact?.sourceTaskKind || taskContext?.kind,
      sourceEventName: currentArtifact?.sourceEventName || taskContext?.eventName,
      correlationId: currentArtifact?.correlationId || taskContext?.correlationId,
      sourceTurnId: turnId,
      sourceRootTurnId: rootTurnId,
      sourceParentTurnId: parentTurnId,
      selectedItems: taskContext?.selectedItems,
      targetContext: mergeMcpTargetContexts(
        latestArtifactContext.targetContext,
        currentArtifact?.targetContext,
        taskContext?.targetContext
      )
    };
  }

  public emitEvent(
    event: Omit<IHieEvent, 'eventId' | 'timestamp'> & { eventId?: string; timestamp?: number }
  ): void {
    const resolvedTurnLineage = this.resolveEventTurnLineage(event);
    const normalizedEvent: IHieEvent = {
      ...event,
      eventId: event.eventId || createCorrelationId('hieevt'),
      timestamp: event.timestamp || Date.now(),
      turnId: event.turnId || resolvedTurnLineage?.turnId,
      rootTurnId: event.rootTurnId || resolvedTurnLineage?.rootTurnId,
      parentTurnId: event.parentTurnId || resolvedTurnLineage?.parentTurnId,
      exposurePolicy: event.exposurePolicy || DEFAULT_HIE_EXPOSURE_POLICY
    };

    this.runOrQueue(`event:${normalizedEvent.eventName}`, () => {
      this.recordEvent(normalizedEvent);
      this.projectEventToPrompt(normalizedEvent);
    });
  }

  /**
   * Called after store.pushBlock() — tracks the new block.
   * For blocks with immediate data, schedules context injection.
   */
  public onBlockCreated(block: IBlock, sourceContext?: IHieSourceContext): void {
    this.runOrQueue(`block-created:${block.id}`, () => {
      const sourceBlockId = this.parseSourceBlockId(block.originTool);
      const trackedSourceBlock = sourceBlockId ? this.tracker.get(sourceBlockId) : undefined;
      const currentTaskContext = this.derivedState.taskContext;
      const sourceTurnLineage = this.buildTurnLineage(
        sourceContext?.sourceTurnId,
        sourceContext?.sourceRootTurnId,
        sourceContext?.sourceParentTurnId
      );
      const blockTurnLineage = sourceTurnLineage
        || this.getTrackedBlockTurnLineage(trackedSourceBlock)
        || this.currentTurn
        || this.getTaskContextTurnLineage(currentTaskContext);

      this.tracker.track(block, blockTurnLineage);
      this.recordEvent({
        eventId: createCorrelationId('hieevt'),
        eventName: 'block.created',
        source: 'hie',
        surface: 'action-panel',
        correlationId: createCorrelationId('ctx'),
        turnId: blockTurnLineage?.turnId,
        rootTurnId: blockTurnLineage?.rootTurnId,
        parentTurnId: blockTurnLineage?.parentTurnId,
        timestamp: Date.now(),
        payload: this.buildBlockEventPayload(block, trackedSourceBlock, currentTaskContext, blockTurnLineage, sourceContext),
        exposurePolicy: { mode: 'store-only', relevance: 'background' },
        blockId: block.id,
        blockType: block.type
      });

      const tracked = this.tracker.get(block.id);
      if (tracked && tracked.state === 'ready') {
        this.injector.scheduleInjection(block.id);
      }

      if (block.type === 'form') {
        const formData = block.data as { preset?: string };
        const formTargetContext = (block.data as { submissionTarget?: { targetContext?: unknown } }).submissionTarget?.targetContext;
        const sourceArtifactId = currentTaskContext?.derivedBlockId
          && currentTaskContext.derivedBlockType
          && currentTaskContext.derivedBlockType !== 'form'
          ? currentTaskContext.derivedBlockId
          : undefined;
        const contextualSourceBlockId = sourceArtifactId
          || currentTaskContext?.sourceBlockId
          || sourceBlockId;
        const contextualSourceBlockType = sourceArtifactId
          ? currentTaskContext?.derivedBlockType
          : (currentTaskContext?.sourceBlockType || trackedSourceBlock?.type || undefined);
        const contextualSourceBlockTitle = sourceArtifactId
          ? currentTaskContext?.derivedBlockTitle
          : (currentTaskContext?.sourceBlockTitle || trackedSourceBlock?.title || undefined);
        this.emitEvent({
          eventName: 'form.opened',
          source: 'action-panel',
          surface: 'action-panel',
          correlationId: createCorrelationId('form'),
          payload: {
            blockId: block.id,
            blockTitle: block.title,
            preset: formData.preset,
            sourceBlockId: contextualSourceBlockId,
            sourceBlockType: contextualSourceBlockType,
            sourceBlockTitle: contextualSourceBlockTitle,
            sourceArtifactId,
            sourceTaskKind: currentTaskContext?.kind,
            sourceEventName: currentTaskContext?.eventName,
            sourceCorrelationId: currentTaskContext?.correlationId,
            sourceTurnId: trackedSourceBlock?.turnId || currentTaskContext?.turnId || blockTurnLineage?.turnId,
            sourceRootTurnId: trackedSourceBlock?.rootTurnId || currentTaskContext?.rootTurnId || blockTurnLineage?.rootTurnId,
            sourceParentTurnId: trackedSourceBlock?.parentTurnId || currentTaskContext?.parentTurnId || blockTurnLineage?.parentTurnId,
            linkedSourceBlockId: sourceBlockId,
            targetContext: formTargetContext || currentTaskContext?.targetContext,
            selectedItems: currentTaskContext?.selectedItems
          },
          exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
          turnId: trackedSourceBlock?.turnId || currentTaskContext?.turnId || blockTurnLineage?.turnId,
          rootTurnId: trackedSourceBlock?.rootTurnId || currentTaskContext?.rootTurnId || blockTurnLineage?.rootTurnId,
          parentTurnId: trackedSourceBlock?.parentTurnId || currentTaskContext?.parentTurnId || blockTurnLineage?.parentTurnId,
          blockId: block.id,
          blockType: block.type
        });
      }

      this.triggerGazeToActionPanel();
      logService.debug('system', `HIE: Block created: ${block.type} (${block.id})`);
    });
  }

  /**
   * Called after store.updateBlock() in async callbacks.
   * The block now has real data — update tracker and schedule context injection.
   */
  public onBlockUpdated(blockId: string, block: IBlock, sourceContext?: IHieSourceContext): void {
    this.runOrQueue(`block-updated:${blockId}`, () => {
      const sourceBlockId = this.parseSourceBlockId(block.originTool);
      const trackedSourceBlock = sourceBlockId ? this.tracker.get(sourceBlockId) : undefined;
      const currentTaskContext = this.derivedState.taskContext;
      this.tracker.update(blockId, block);
      const tracked = this.tracker.get(blockId);
      const trackedTurnLineage = this.getTrackedBlockTurnLineage(tracked);
      const sourceTurnLineage = this.buildTurnLineage(
        sourceContext?.sourceTurnId,
        sourceContext?.sourceRootTurnId,
        sourceContext?.sourceParentTurnId
      );
      this.recordEvent({
        eventId: createCorrelationId('hieevt'),
        eventName: 'block.updated',
        source: 'hie',
        surface: 'action-panel',
        correlationId: createCorrelationId('ctx'),
        turnId: sourceTurnLineage?.turnId || tracked?.turnId || this.currentTurn?.turnId,
        rootTurnId: sourceTurnLineage?.rootTurnId || tracked?.rootTurnId || this.currentTurn?.rootTurnId,
        parentTurnId: sourceTurnLineage?.parentTurnId || tracked?.parentTurnId || this.currentTurn?.parentTurnId,
        timestamp: Date.now(),
        payload: this.buildBlockEventPayload(block, trackedSourceBlock, currentTaskContext, trackedTurnLineage, sourceContext),
        exposurePolicy: { mode: 'store-only', relevance: 'background' },
        blockId,
        blockType: block.type
      });
      this.injector.scheduleInjection(blockId);
      this.triggerGazeToActionPanel();

      logService.debug('system', `HIE: Block updated: ${block.type} (${blockId})`);
    });
  }

  /**
   * Called when the user mouses over a fresh block, acknowledging it.
   * Updates the tracker state from 'ready' to 'acknowledged'.
   */
  public onBlockAcknowledged(blockId: string): void {
    this.runOrQueue(`block-acknowledged:${blockId}`, () => {
      this.tracker.markAcknowledged(blockId);
    });
  }

  /**
   * Called when a block is dismissed by the user.
   */
  public onBlockRemoved(blockId: string): void {
    this.runOrQueue(`block-removed:${blockId}`, () => {
      const tracked = this.tracker.get(blockId);
      const artifact = this.derivedState.artifacts[blockId];
      const currentTaskContext = this.derivedState.taskContext;
      const sourceBlockId = artifact?.sourceBlockId || currentTaskContext?.sourceBlockId;
      const trackedSourceBlock = sourceBlockId ? this.tracker.get(sourceBlockId) : undefined;
      this.tracker.remove(blockId);

      if (tracked?.type === 'form') {
        this.emitEvent({
          eventName: 'form.dismissed',
          source: 'action-panel',
          surface: 'action-panel',
          correlationId: createCorrelationId('form'),
          turnId: artifact?.sourceTurnId || currentTaskContext?.turnId || tracked.turnId || this.currentTurn?.turnId,
          rootTurnId: artifact?.sourceRootTurnId || currentTaskContext?.rootTurnId || tracked.rootTurnId || this.currentTurn?.rootTurnId,
          parentTurnId: artifact?.sourceParentTurnId || currentTaskContext?.parentTurnId || tracked.parentTurnId || this.currentTurn?.parentTurnId,
          payload: {
            blockId,
            blockTitle: tracked.title,
            preset: artifact?.preset,
            formStatus: 'dismissed',
            sourceBlockId,
            sourceBlockType: trackedSourceBlock?.type || currentTaskContext?.sourceBlockType,
            sourceBlockTitle: trackedSourceBlock?.title || currentTaskContext?.sourceBlockTitle,
            sourceArtifactId: artifact?.sourceArtifactId,
            sourceTaskKind: artifact?.sourceTaskKind || currentTaskContext?.kind,
            sourceEventName: artifact?.sourceEventName || currentTaskContext?.eventName,
            sourceCorrelationId: artifact?.correlationId || currentTaskContext?.correlationId,
            sourceTurnId: artifact?.sourceTurnId || currentTaskContext?.turnId || tracked.turnId || this.currentTurn?.turnId,
            sourceRootTurnId: artifact?.sourceRootTurnId || currentTaskContext?.rootTurnId || tracked.rootTurnId || this.currentTurn?.rootTurnId,
            sourceParentTurnId: artifact?.sourceParentTurnId || currentTaskContext?.parentTurnId || tracked.parentTurnId || this.currentTurn?.parentTurnId,
            targetContext: artifact?.targetContext || currentTaskContext?.targetContext,
            selectedItems: currentTaskContext?.selectedItems
          },
          exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
          blockId,
          blockType: tracked.type
        });
      } else if (tracked) {
        this.emitEvent({
          eventName: 'block.removed',
          source: 'action-panel',
          surface: 'action-panel',
          correlationId: createCorrelationId('ctx'),
          turnId: tracked.turnId || this.currentTurn?.turnId,
          rootTurnId: tracked.rootTurnId || this.currentTurn?.rootTurnId,
          parentTurnId: tracked.parentTurnId || this.currentTurn?.parentTurnId,
          payload: {
            blockId,
            blockTitle: tracked.title,
            blockType: tracked.type
          },
          exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
          blockId,
          blockType: tracked.type
        });
      }

      if (!tracked || tracked.type === 'form') {
        this.recordEvent({
          eventId: createCorrelationId('hieevt'),
          eventName: 'block.removed',
          source: 'hie',
          surface: 'action-panel',
          correlationId: createCorrelationId('ctx'),
          turnId: tracked?.turnId || this.currentTurn?.turnId,
          rootTurnId: tracked?.rootTurnId || this.currentTurn?.rootTurnId,
          parentTurnId: tracked?.parentTurnId || this.currentTurn?.parentTurnId,
          timestamp: Date.now(),
          payload: { blockId, blockTitle: tracked?.title, blockType: tracked?.type },
          exposurePolicy: { mode: 'store-only', relevance: 'background' },
          blockId,
          blockType: tracked?.type
        });
      }
    });
  }

  /**
   * Unified entry point for ALL block clicks/selections.
   * Handles context injection, expression updates, and flow advancement.
   */
  public onBlockInteraction(interaction: IBlockInteraction): void {
    this.runOrQueue(`interaction:${interaction.action}:${interaction.blockId}`, () => {
      this.tracker.markInteracted(interaction.blockId);
      this.verbosityDirector.recordInteraction();
      const trackedBlock = this.tracker.get(interaction.blockId);
      const interactionPayload: Record<string, unknown> = {
        ...interaction.payload,
        sourceBlockId: typeof interaction.payload.sourceBlockId === 'string' && interaction.payload.sourceBlockId.trim()
          ? interaction.payload.sourceBlockId
          : interaction.blockId,
        sourceBlockType: typeof interaction.payload.sourceBlockType === 'string' && interaction.payload.sourceBlockType.trim()
          ? interaction.payload.sourceBlockType
          : interaction.blockType,
        sourceBlockTitle: typeof interaction.payload.sourceBlockTitle === 'string' && interaction.payload.sourceBlockTitle.trim()
          ? interaction.payload.sourceBlockTitle
          : trackedBlock?.title,
        ...(this.derivedState.userTextPreview ? { userTextPreview: this.derivedState.userTextPreview } : {})
      };
      const interactionTurnLineage = this.resolveEventTurnLineage({
        eventName: interaction.eventName || `block.interaction.${interaction.action}`,
        source: interaction.source || 'block-ui',
        surface: interaction.surface || 'action-panel',
        correlationId: interaction.correlationId || createCorrelationId('iact'),
        turnId: interaction.turnId,
        rootTurnId: interaction.rootTurnId,
        parentTurnId: interaction.parentTurnId,
        payload: interactionPayload,
        exposurePolicy: interaction.exposurePolicy || DEFAULT_HIE_EXPOSURE_POLICY,
        blockId: interaction.blockId,
        blockType: interaction.blockType
      });
      const event: IHieEvent = {
        eventId: createCorrelationId('hieevt'),
        eventName: interaction.eventName || `block.interaction.${interaction.action}`,
        source: interaction.source || 'block-ui',
        surface: interaction.surface || 'action-panel',
        correlationId: interaction.correlationId || createCorrelationId('iact'),
        turnId: interaction.turnId || interactionTurnLineage?.turnId,
        rootTurnId: interaction.rootTurnId || interactionTurnLineage?.rootTurnId,
        parentTurnId: interaction.parentTurnId || interactionTurnLineage?.parentTurnId,
        timestamp: interaction.timestamp || Date.now(),
        payload: interactionPayload,
        exposurePolicy: interaction.exposurePolicy || DEFAULT_HIE_EXPOSURE_POLICY,
        blockId: interaction.blockId,
        blockType: interaction.blockType
      };
      this.recordEvent(event);
      this.projectEventToPrompt(event);

      this.expressionDirector.onInteraction(interaction.action);

      const flowUpdate = this.flowOrchestrator.onInteraction(interaction);
      if (flowUpdate) {
        this.injector.injectFlowUpdate(flowUpdate, [interaction.blockId]);
      }

      const envelope: IInteractionEnvelope = {
        envelopeId: createCorrelationId('iactenv'),
        correlationId: event.correlationId,
        createdAt: Date.now(),
        source: interaction.source || 'block-ui',
        blockId: interaction.blockId,
        blockType: interaction.blockType,
        action: interaction.action,
        triggerResponse: (interaction.exposurePolicy || DEFAULT_HIE_EXPOSURE_POLICY).mode === 'response-triggering',
        schemaId: interaction.schemaId,
        turnId: event.turnId,
        rootTurnId: event.rootTurnId,
        parentTurnId: event.parentTurnId,
        payload: interaction.payload
      };

      if (interaction.action === 'submit-form' || interaction.action === 'cancel-form') {
        this.emitEvent({
          eventName: interaction.action === 'submit-form' ? 'form.submitted' : 'form.cancelled',
          source: interaction.source || 'form',
          surface: interaction.surface || 'action-panel',
          correlationId: event.correlationId,
          payload: {
            blockId: interaction.blockId,
            preset: typeof interaction.payload.preset === 'string' ? interaction.payload.preset : undefined,
            formStatus: interaction.action === 'submit-form' ? 'submitted' : 'cancelled',
            sourceBlockId: this.derivedState.taskContext?.sourceBlockId,
            sourceBlockTitle: this.derivedState.taskContext?.sourceBlockTitle,
            sourceEventName: this.derivedState.taskContext?.eventName,
            sourceCorrelationId: this.derivedState.taskContext?.correlationId,
            sourceTurnId: this.derivedState.taskContext?.turnId || interactionTurnLineage?.turnId,
            sourceRootTurnId: this.derivedState.taskContext?.rootTurnId || interactionTurnLineage?.rootTurnId,
            sourceParentTurnId: this.derivedState.taskContext?.parentTurnId || interactionTurnLineage?.parentTurnId,
            targetContext: this.derivedState.taskContext?.targetContext,
            selectedItems: this.derivedState.taskContext?.selectedItems
          },
          exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
          turnId: interaction.turnId || interactionTurnLineage?.turnId,
          rootTurnId: interaction.rootTurnId || interactionTurnLineage?.rootTurnId,
          parentTurnId: interaction.parentTurnId || interactionTurnLineage?.parentTurnId,
          blockId: interaction.blockId,
          blockType: interaction.blockType
        });
      }

      logService.info('system', `HIE: Interaction: ${interaction.action} on ${interaction.blockType} (${interaction.blockId})`);
      logService.debug('system', 'HIE interaction envelope', JSON.stringify(envelope));
    });
  }

  /**
   * Set whether the voice (WebRTC) path is active.
   * When true, async tool completions will trigger the LLM to speak about results.
   */
  public setVoicePathActive(active: boolean): void {
    this.voicePathActive = active;
    logService.debug('system', `HIE: voicePathActive = ${active}`);
  }

  public isVoicePathActive(): boolean {
    return this.voicePathActive;
  }

  /**
   * Configure a local callback for async tool completion acknowledgments.
   * Useful to keep voice/chat aligned with visual updates without extra LLM latency.
   */
  public setAsyncToolCompletionHandler(
    handler?: (feedback: IAsyncToolCompletionFeedback) => void
  ): void {
    this.onAsyncToolCompletion = handler;
  }

  /**
   * Called when an async tool operation finishes.
   * Triggers data-aware expressions, flow detection, and voice feedback.
   */
  public onToolComplete(
    toolName: string,
    blockId: string,
    success: boolean,
    itemCount: number,
    sourceContext?: IHieSourceContext
  ): void {
    this.runOrQueue(`tool-complete:${toolName}:${blockId || 'none'}`, () => {
      this.expressionDirector.onToolComplete(toolName, success, itemCount);
      const tracked = blockId ? this.tracker.get(blockId) : undefined;
      const resolvedSourceContext = sourceContext || this.captureCurrentSourceContext();
      const toolTurnLineage = this.buildTurnLineage(
        resolvedSourceContext?.sourceTurnId,
        resolvedSourceContext?.sourceRootTurnId,
        resolvedSourceContext?.sourceParentTurnId
      );

      this.emitEvent({
        eventName: success ? 'tool.execution.completed' : 'tool.execution.failed',
        source: 'tool-runtime',
        surface: 'tool-runtime',
        correlationId: createCorrelationId('tool'),
        payload: {
          toolName,
          blockId: blockId || tracked?.id,
          blockTitle: tracked?.title,
          blockType: tracked?.type,
          success,
          itemCount,
          sourceBlockId: resolvedSourceContext?.sourceBlockId,
          sourceBlockType: resolvedSourceContext?.sourceBlockType,
          sourceBlockTitle: resolvedSourceContext?.sourceBlockTitle,
          sourceArtifactId: resolvedSourceContext?.sourceArtifactId,
          sourceTaskKind: resolvedSourceContext?.sourceTaskKind,
          sourceEventName: resolvedSourceContext?.sourceEventName,
          sourceCorrelationId: resolvedSourceContext?.correlationId,
          sourceTurnId: resolvedSourceContext?.sourceTurnId,
          sourceRootTurnId: resolvedSourceContext?.sourceRootTurnId,
          sourceParentTurnId: resolvedSourceContext?.sourceParentTurnId,
          targetContext: resolvedSourceContext?.targetContext
        },
        exposurePolicy: { mode: 'store-only', relevance: 'contextual' },
        turnId: toolTurnLineage?.turnId,
        rootTurnId: toolTurnLineage?.rootTurnId,
        parentTurnId: toolTurnLineage?.parentTurnId,
        blockId: blockId || tracked?.id,
        blockType: tracked?.type
      });

      const flowUpdate = this.flowOrchestrator.onToolInvoked(toolName, blockId);
      if (flowUpdate) {
        this.injector.injectFlowUpdate(flowUpdate, [blockId]);
      }

      if (!success && blockId) {
        if (tracked && tracked.type === 'error') {
          const errorText = tracked.summary.replace(/^Error:\s*/i, '').trim();
          this.sendToolError(toolName, errorText || 'Tool execution failed');
        }
      }
      if (!success && !blockId && !EXPLICIT_TOOL_ERROR_CONTEXT_TOOLS.has(toolName)) {
        this.sendToolError(toolName, 'The operation failed. Check the error card in the action panel if present.');
      }

      if (this.voicePathActive && this.config.asyncToolResponseEnabled && success && ASYNC_TOOLS.has(toolName)) {
        const canUseLocalFeedback = !!this.onAsyncToolCompletion && LOCAL_COMPLETION_FEEDBACK_TOOLS.has(toolName);

        if (canUseLocalFeedback) {
          this.injector.injectToolCompletion({ toolName, blockId, itemCount, triggerResponse: false });
          this.onAsyncToolCompletion?.({
            toolName,
            blockId,
            itemCount,
            summary: tracked?.summary
          });
        } else {
          this.injector.injectToolCompletion({ toolName, blockId, itemCount, triggerResponse: true });
        }
      }

      const envelope: IToolExecutionEnvelope = {
        envelopeId: createCorrelationId('toolenv'),
        correlationId: createCorrelationId('tool'),
        createdAt: Date.now(),
        toolName,
        phase: success ? 'complete' : 'error',
        awaitAsync: true,
        blockId,
        success,
        itemCount
      };

      logService.debug('system', `HIE: Tool complete: ${toolName} (success=${success}, items=${itemCount})`, JSON.stringify(envelope));
    });
  }

  /**
   * Called when the LLM explicitly sets an expression via set_expression tool.
   * Ensures the LLM's explicit choice takes priority over data-aware rules.
   */
  public onLlmExpression(): void {
    if (!this.initialized) return;
    this.expressionDirector.onLlmExpression();
  }

  /**
   * Called when the user sends a text message (text input path).
   * Pre-injects visual state context so the LLM has grounding before responding.
   * For the voice path, grounding happens periodically via onLlmResponse().
   */
  public onUserMessage(turnId?: string): void {
    this.setCurrentTurnId(turnId);
    if (!this.initialized) return;
    if (this.tracker.getSize() === 0) return;

    // Always inject current visual state before a text message
    const summary = this.getProjectedCurrentStateSummary();
    if (summary) {
      const snapshot = this.getVisualStateSnapshot(turnId);
      this.injector.injectVisualStateReminder(summary);
      this.turnsSinceGrounding = 0;
      logService.debug('system', 'HIE: Visual state grounding injected (text input)');
      logService.debug('system', 'HIE visual snapshot', JSON.stringify(snapshot));
    }
  }

  /**
   * Called when the LLM produces a response (voice or text).
   * Tracks turns and periodically re-injects visual state to prevent context drift.
   */
  public onLlmResponse(): void {
    if (!this.initialized) return;
    if (this.tracker.getSize() === 0) return;

    this.turnsSinceGrounding++;

    if (this.turnsSinceGrounding >= GROUNDING_INTERVAL_TURNS) {
      const summary = this.getProjectedCurrentStateSummary();
      if (summary) {
        this.injector.injectVisualStateReminder(summary);
        this.turnsSinceGrounding = 0;
        logService.debug('system', 'HIE: Periodic visual state grounding injected');
      }
    }
  }

  /**
   * Returns a compact text summary of the current visual state.
   * Can be used for system prompt enrichment.
   */
  public getVisualStateSummary(): string {
    if (!this.initialized) return '';

    const blockSummary = this.tracker.getActiveSummary();
    const flowSummary = this.flowOrchestrator.getActiveFlowSummary();

    const parts: string[] = [];
    if (blockSummary) parts.push(blockSummary);
    if (flowSummary) parts.push(`Active flows: ${flowSummary}`);

    return parts.join(' | ');
  }

  public getProjectedCurrentStateSummary(): string {
    if (!this.initialized) return '';
    const visualSummary = this.getVisualStateSummaryWithReferences();
    return this.contextProjector.buildCurrentStateReminder(
      visualSummary,
      this.derivedState,
      this.config.maxContextLength
    );
  }

  /**
   * Build a formal visual workspace snapshot for observability and deterministic
   * replay of interaction context decisions.
   */
  public getVisualStateSnapshot(correlationId?: string): IVisualStateSnapshot {
    const blocks = this.tracker.getActiveBlocks();
    const referencesNewestFirst = this.tracker.getAllReferences();
    const flowSummary = this.flowOrchestrator.getActiveFlowSummary();

    const seenTypes = new Set<string>();
    let hasDuplicateTypes = false;
    for (let i = 0; i < blocks.length; i++) {
      if (seenTypes.has(blocks[i].type)) {
        hasDuplicateTypes = true;
        break;
      }
      seenTypes.add(blocks[i].type);
    }

    return {
      snapshotId: createCorrelationId('snap'),
      correlationId: correlationId || createCorrelationId('turn'),
      createdAt: Date.now(),
      summary: this.getProjectedCurrentStateSummary(),
      flowSummary,
      hasDuplicateTypes,
      blocks: blocks.map((b) => ({
        blockId: b.id,
        blockType: b.type,
        title: b.title,
        state: b.state,
        summary: b.summary,
        itemCount: b.itemCount,
        references: b.references,
        createdAt: b.createdAt,
        updatedAt: b.updatedAt
      })),
      referencesNewestFirst
    };
  }

  /**
   * Returns visual state summary enriched with numbered item references.
   * Used by grounding injection to give the LLM precise item indices.
   */
  public getVisualStateSummaryWithReferences(): string {
    if (!this.initialized) return '';

    const allRefs = this.tracker.getAllReferences();
    const flowSummary = this.flowOrchestrator.getActiveFlowSummary();

    const parts: string[] = [];
    const seenTypes = new Set<string>();
    let hasDuplicateTypes = false;

    // allRefs is sorted newest-first by BlockTracker
    allRefs.forEach((block) => {
      const tracked = this.tracker.get(block.blockId);
      if (!tracked) return;
      const items = block.references.slice(0, 10).map((ref) => {
        let entry = `${ref.index}) ${ref.title}`;
        if (ref.detail) entry += ` (${ref.detail})`;
        return entry;
      });
      const suffix = block.references.length > 10
        ? ` ...and ${block.references.length - 10} more`
        : '';
      // Tag older blocks of the same type so the LLM prioritizes the newest
      const isRepeat = seenTypes.has(tracked.type);
      seenTypes.add(tracked.type);
      if (isRepeat) hasDuplicateTypes = true;
      const tag = isRepeat ? ' [EARLIER — not current results]' : '';
      parts.push(`${tracked.summary}${tag}. Items: ${items.join(', ')}${suffix}`);
    });

    // When multiple blocks of the same type exist, prepend a resolution hint
    if (hasDuplicateTypes) {
      parts.unshift('IMPORTANT: Results are listed newest-first. Positional references ("first doc", "item 2") refer to the MOST RECENT results, not earlier ones.');
    }

    // Add blocks without references
    const blockSummary = this.tracker.getActiveSummary();
    if (parts.length === 0 && blockSummary) {
      parts.push(blockSummary);
    }

    if (flowSummary) parts.push(`Active flows: ${flowSummary}`);

    // Append verbosity hint if applicable
    const verbosity = this.verbosityDirector.getVerbosityHint();
    if (verbosity) {
      parts.push(`[Verbosity: ${verbosity.level} — ${verbosity.hint}]`);
    }

    return parts.join(' | ');
  }

  private recordEvent(event: IHieEvent): void {
    if (event.turnId) {
      this.currentTurn = this.rememberTurnLineage({
        turnId: event.turnId,
        rootTurnId: event.rootTurnId || event.turnId,
        parentTurnId: event.parentTurnId
      });
    }
    this.recentEvents.push(event);
    if (this.recentEvents.length > HIE_EVENT_HISTORY_LIMIT) {
      this.recentEvents = this.recentEvents.slice(-HIE_EVENT_HISTORY_LIMIT);
    }
    this.derivedState = reduceHieDerivedState(this.derivedState, event);

    const envelope: IHieEventEnvelope = {
      envelopeId: createCorrelationId('hieenv'),
      eventId: event.eventId,
      correlationId: event.correlationId,
      turnId: event.turnId,
      rootTurnId: event.rootTurnId,
      parentTurnId: event.parentTurnId,
      createdAt: event.timestamp,
      eventName: event.eventName,
      source: event.source,
      surface: event.surface,
      exposureMode: event.exposurePolicy.mode,
      blockId: event.blockId,
      blockType: event.blockType,
      payload: event.payload
    };
    logService.debug('system', 'HIE event envelope', JSON.stringify(envelope));
  }

  private projectEventToPrompt(event: IHieEvent): void {
    if (event.exposurePolicy.mode === 'store-only') {
      return;
    }
    const isBlockInteractionEvent = event.eventName.startsWith('block.interaction.');
    if (isBlockInteractionEvent) {
      if (event.exposurePolicy.mode === 'response-triggering' && !this.config.interactionNotificationsEnabled) {
        return;
      }
      if (event.exposurePolicy.mode === 'silent-context' && !this.config.contextInjectionEnabled) {
        return;
      }
    } else if (!this.config.contextInjectionEnabled) {
      return;
    }

    const projected = this.contextProjector.projectEvent(event, this.derivedState, this.tracker);
    if (!projected) {
      return;
    }

    this.injector.injectProjectedContext(
      projected,
      event.exposurePolicy.mode === 'response-triggering',
      event.blockId ? [event.blockId] : [],
      event.turnId
    );
  }

  /**
   * Send a tool error message to the LLM so it can acknowledge and retry.
   * Triggering a response ensures the user gets immediate voice/text feedback.
   */
  public sendToolError(toolName: string, errorMessage: string): void {
    if (!this.initialized || !this.config.contextInjectionEnabled || !this.toolErrorInjectionEnabled) return;

    const message: IHiePromptMessage = {
      kind: 'tool-error',
      body: `${toolName} failed — ${errorMessage}. You may retry once or suggest an alternative.`
    };
    this.injector.injectProjectedContext(message, true, [], this.currentTurn?.turnId);

    logService.debug('system', `HIE: Tool error injected for ${toolName}`);
  }

  /**
   * Suppress expression revert timers during multi-step tool loops.
   * Prevents thinking→happy→idle→thinking flashing.
   */
  public suppressExpressionReverts(): void {
    if (!this.initialized) return;
    this.expressionDirector.suppressReverts();
  }

  /**
   * Release expression revert suppression after tool loop ends.
   */
  public releaseExpressionReverts(): void {
    if (!this.initialized) return;
    this.expressionDirector.releaseReverts();
  }

  /**
   * Get the current configuration.
   */
  public getConfig(): Readonly<IHIEConfig> {
    return this.config;
  }

  /** Briefly gaze toward the action panel when new data appears. */
  private triggerGazeToActionPanel(): void {
    this.setGaze('action-panel');
    if (this.gazeRevertTimer) clearTimeout(this.gazeRevertTimer);
    this.gazeRevertTimer = setTimeout(() => {
      this.setGaze('none');
      this.gazeRevertTimer = undefined;
    }, 3000);
  }

  /**
   * Check if the engine has been initialized.
   */
  public isInitialized(): boolean {
    return this.initialized;
  }

  private parseSourceBlockId(originTool?: string): string | undefined {
    if (!originTool) return undefined;
    if (originTool.startsWith('share-form:')) {
      return originTool.slice('share-form:'.length);
    }
    if (originTool.startsWith('block-recap:')) {
      return originTool.slice('block-recap:'.length);
    }
    return undefined;
  }

  private buildBlockEventPayload(
    block: IBlock,
    trackedSourceBlock?: Readonly<{ id: string; type: IBlock['type']; title: string; turnId?: string; rootTurnId?: string; parentTurnId?: string }>,
    currentTaskContext?: Readonly<IHieTaskContext>,
    blockTurnLineage?: Readonly<IHieTurnLineage>,
    sourceContext?: Readonly<IHieSourceContext>
  ): Record<string, unknown> {
    const existingArtifact = this.derivedState.artifacts[block.id];
    const formSubmissionTargetContext = block.type === 'form'
      ? ((block.data as { submissionTarget?: { targetContext?: unknown } }).submissionTarget?.targetContext as Record<string, unknown> | undefined)
      : undefined;
    const originSourceBlockId = this.parseSourceBlockId(block.originTool);
    const hasConflictingCurrentTurn = !!currentTaskContext && !!this.currentTurn && (
      (!!currentTaskContext.rootTurnId && currentTaskContext.rootTurnId !== this.currentTurn.rootTurnId)
      || (
        !currentTaskContext.rootTurnId
        && !!currentTaskContext.turnId
        && currentTaskContext.turnId !== this.currentTurn.turnId
      )
    );
    const canInheritTaskContext = !!currentTaskContext && (
      !!sourceContext
      || !!trackedSourceBlock
      || !!originSourceBlockId
      || !hasConflictingCurrentTurn
    );
    const effectiveTaskContext = canInheritTaskContext ? currentTaskContext : undefined;
    const blockTargetContext = mergeMcpTargetContexts(
      effectiveTaskContext?.targetContext,
      existingArtifact?.targetContext,
      sourceContext?.targetContext,
      deriveMcpTargetContextFromUnknown(block.data, 'hie-selection'),
      deriveMcpTargetContextFromUnknown(formSubmissionTargetContext, 'explicit-user')
    );
    const sourceArtifactId = existingArtifact?.sourceArtifactId
      || sourceContext?.sourceArtifactId
      || (effectiveTaskContext?.derivedBlockId && effectiveTaskContext.derivedBlockId !== block.id
        ? effectiveTaskContext.derivedBlockId
        : undefined);

    return {
      blockId: block.id,
      blockType: block.type,
      blockTitle: block.title,
      originTool: block.originTool,
      sourceBlockId: existingArtifact?.sourceBlockId
        || sourceContext?.sourceBlockId
        || effectiveTaskContext?.sourceBlockId
        || trackedSourceBlock?.id
        || originSourceBlockId,
      sourceBlockType: sourceContext?.sourceBlockType || effectiveTaskContext?.sourceBlockType || trackedSourceBlock?.type,
      sourceBlockTitle: sourceContext?.sourceBlockTitle || effectiveTaskContext?.sourceBlockTitle || trackedSourceBlock?.title,
      sourceArtifactId,
      sourceTaskKind: existingArtifact?.sourceTaskKind || sourceContext?.sourceTaskKind || effectiveTaskContext?.kind,
      sourceEventName: existingArtifact?.sourceEventName || sourceContext?.sourceEventName || effectiveTaskContext?.eventName,
      sourceCorrelationId: existingArtifact?.correlationId || sourceContext?.correlationId || effectiveTaskContext?.correlationId,
      sourceTurnId: existingArtifact?.sourceTurnId
        || sourceContext?.sourceTurnId
        || effectiveTaskContext?.turnId
        || trackedSourceBlock?.turnId
        || blockTurnLineage?.turnId,
      sourceRootTurnId: existingArtifact?.sourceRootTurnId
        || sourceContext?.sourceRootTurnId
        || effectiveTaskContext?.rootTurnId
        || trackedSourceBlock?.rootTurnId
        || blockTurnLineage?.rootTurnId,
      sourceParentTurnId: existingArtifact?.sourceParentTurnId
        || sourceContext?.sourceParentTurnId
        || effectiveTaskContext?.parentTurnId
        || trackedSourceBlock?.parentTurnId
        || blockTurnLineage?.parentTurnId,
      selectedItems: sourceContext?.selectedItems || effectiveTaskContext?.selectedItems,
      targetContext: blockTargetContext,
      linkedSourceBlockId: originSourceBlockId
    };
  }

  private resolveEventTurnLineage(
    event: Omit<IHieEvent, 'eventId' | 'timestamp'> & { eventId?: string; timestamp?: number }
  ): IHieTurnLineage | undefined {
    const payload = event.payload;
    const eventTurnLineage = this.buildTurnLineage(event.turnId, event.rootTurnId, event.parentTurnId);
    if (eventTurnLineage) {
      return eventTurnLineage;
    }

    const payloadTurnId = typeof payload.sourceTurnId === 'string' && payload.sourceTurnId.trim()
      ? payload.sourceTurnId.trim()
      : undefined;
    const payloadRootTurnId = typeof payload.sourceRootTurnId === 'string' && payload.sourceRootTurnId.trim()
      ? payload.sourceRootTurnId.trim()
      : undefined;
    const payloadParentTurnId = typeof payload.sourceParentTurnId === 'string' && payload.sourceParentTurnId.trim()
      ? payload.sourceParentTurnId.trim()
      : undefined;
    const payloadTurnLineage = this.buildTurnLineage(payloadTurnId, payloadRootTurnId, payloadParentTurnId);
    if (payloadTurnLineage) {
      return payloadTurnLineage;
    }

    const payloadSourceArtifactId = typeof payload.sourceArtifactId === 'string' && payload.sourceArtifactId.trim()
      ? payload.sourceArtifactId.trim()
      : undefined;
    if (payloadSourceArtifactId) {
      const sourceArtifact = this.derivedState.artifacts[payloadSourceArtifactId];
      const sourceArtifactLineage = this.buildTurnLineage(
        sourceArtifact?.sourceTurnId,
        sourceArtifact?.sourceRootTurnId,
        sourceArtifact?.sourceParentTurnId
      );
      if (sourceArtifactLineage) {
        return sourceArtifactLineage;
      }
    }

    const payloadSourceBlockId = typeof payload.sourceBlockId === 'string' && payload.sourceBlockId.trim()
      ? payload.sourceBlockId.trim()
      : undefined;
    if (payloadSourceBlockId) {
      const sourceBlock = this.tracker.get(payloadSourceBlockId);
      const sourceBlockLineage = this.getTrackedBlockTurnLineage(sourceBlock);
      if (sourceBlockLineage) {
        return sourceBlockLineage;
      }
    }

    if (event.blockId) {
      const tracked = this.tracker.get(event.blockId);
      const trackedLineage = this.getTrackedBlockTurnLineage(tracked);
      if (trackedLineage) {
        return trackedLineage;
      }
    }

    return this.getCurrentContextTurnLineage();
  }

  /**
   * Called on voice session disconnect — clears all state.
   */
  public reset(): void {
    this.tracker.clear();
    this.injector.reset();
    this.expressionDirector.reset();
    this.flowOrchestrator.reset();
    this.verbosityDirector.reset();
    this.turnsSinceGrounding = 0;
    this.recentEvents = [];
    this.derivedState = createInitialHieDerivedState();
    this.pendingInitializationActions = [];
    this.initialized = false;
    this.voicePathActive = false;
    this.currentTurn = undefined;
    this.turnLineageById.clear();
    this.toolErrorInjectionEnabled = true;
    this.onAsyncToolCompletion = undefined;
    this.setGaze = () => { /* no-op */ };
    if (this.gazeRevertTimer) { clearTimeout(this.gazeRevertTimer); this.gazeRevertTimer = undefined; }

    logService.info('system', 'HIE reset');
  }

  private runOrQueue(label: string, action: () => void): void {
    if (this.initialized) {
      action();
      return;
    }

    if (this.pendingInitializationActions.length >= HIE_PENDING_ACTION_LIMIT) {
      this.pendingInitializationActions.shift();
    }
    this.pendingInitializationActions.push({ label, action });
    logService.debug('system', `HIE: queued pre-init action (${label})`);
  }

  private flushPendingInitializationActions(): void {
    if (this.pendingInitializationActions.length === 0) {
      return;
    }

    const pending = [...this.pendingInitializationActions];
    this.pendingInitializationActions = [];
    pending.forEach((entry) => entry.action());
    logService.debug('system', `HIE: replayed ${pending.length} pre-init action${pending.length === 1 ? '' : 's'}`);
  }

  private resolveTurnStartDecision(
    options: IHieTurnStartOptions,
    previousTurn?: IHieTurnLineage
  ): { mode: Exclude<HieTurnStartMode, 'auto'>; reason: string } {
    const requestedMode = options.mode || 'inherit';
    if (requestedMode === 'inherit') {
      return { mode: 'inherit', reason: options.reason || (previousTurn ? 'explicit-inherit' : 'no-previous-thread') };
    }
    if (requestedMode === 'new-root') {
      return { mode: 'new-root', reason: options.reason || 'explicit-new-root' };
    }

    const normalizedText = (options.text || '').trim().toLowerCase();
    if (!previousTurn) {
      return { mode: 'new-root', reason: 'no-previous-thread' };
    }
    if (this.matchesExplicitThreadReset(normalizedText)) {
      return { mode: 'new-root', reason: 'explicit-reset-phrase' };
    }
    if (!this.derivedState.taskContext && this.tracker.getSize() === 0) {
      return { mode: 'new-root', reason: 'no-active-context' };
    }
    return { mode: 'inherit', reason: this.derivedState.taskContext ? 'active-task-context' : 'active-visual-context' };
  }

  private matchesExplicitThreadReset(text: string): boolean {
    if (!text) {
      return false;
    }

    const resetHints: ReadonlyArray<RegExp> = [
      /\bnew topic\b/i,
      /\bstart over\b/i,
      /\bforget that\b/i,
      /\bsomething else\b/i,
      /\bdifferent topic\b/i,
      /\bchange subject\b/i
    ];
    return resetHints.some((pattern) => pattern.test(text));
  }

  private emitThreadLifecycleEvent(
    previousTurn: IHieTurnLineage | undefined,
    currentTurn: IHieTurnLineage,
    resolvedMode: Exclude<HieTurnStartMode, 'auto'>,
    requestedMode: HieTurnStartMode,
    reason: string,
    text?: string
  ): void {
    const eventName = !previousTurn
      ? 'thread.started'
      : (resolvedMode === 'new-root' ? 'thread.reset' : 'thread.continued');
    const userTextPreview = text?.trim()
      ? text.trim().slice(0, 120)
      : undefined;

    this.emitEvent({
      eventName,
      source: 'hie',
      surface: 'unknown',
      correlationId: currentTurn.turnId,
      turnId: currentTurn.turnId,
      rootTurnId: currentTurn.rootTurnId,
      parentTurnId: currentTurn.parentTurnId,
      payload: {
        requestedMode,
        resolvedMode,
        reason,
        previousTurnId: previousTurn?.turnId,
        previousRootTurnId: previousTurn?.rootTurnId,
        currentTurnId: currentTurn.turnId,
        currentRootTurnId: currentTurn.rootTurnId,
        currentParentTurnId: currentTurn.parentTurnId,
        userTextPreview
      },
      exposurePolicy: { mode: 'store-only', relevance: 'contextual' }
    });
  }

  private getCurrentContextTurnLineage(): IHieTurnLineage | undefined {
    return this.currentTurn || this.getTaskContextTurnLineage(this.derivedState.taskContext);
  }

  private getTaskContextTurnLineage(taskContext?: Readonly<IHieTaskContext>): IHieTurnLineage | undefined {
    if (!taskContext?.turnId && !taskContext?.rootTurnId && !taskContext?.parentTurnId) {
      return undefined;
    }
    return this.buildTurnLineage(taskContext?.turnId, taskContext?.rootTurnId, taskContext?.parentTurnId);
  }

  private getTrackedBlockTurnLineage(
    tracked?: Readonly<{ turnId?: string; rootTurnId?: string; parentTurnId?: string }>
  ): IHieTurnLineage | undefined {
    if (!tracked?.turnId && !tracked?.rootTurnId && !tracked?.parentTurnId) {
      return undefined;
    }
    return this.buildTurnLineage(tracked?.turnId, tracked?.rootTurnId, tracked?.parentTurnId);
  }

  private buildTurnLineage(
    turnId?: string,
    rootTurnId?: string,
    parentTurnId?: string
  ): IHieTurnLineage | undefined {
    const resolvedTurnId = turnId?.trim();
    const resolvedRootTurnId = rootTurnId?.trim();
    const resolvedParentTurnId = parentTurnId?.trim();
    if (!resolvedTurnId && !resolvedRootTurnId && !resolvedParentTurnId) {
      return undefined;
    }

    const known = resolvedTurnId ? this.turnLineageById.get(resolvedTurnId) : undefined;
    const finalTurnId = resolvedTurnId || known?.turnId;
    const finalRootTurnId = resolvedRootTurnId || known?.rootTurnId || finalTurnId;
    if (!finalTurnId || !finalRootTurnId) {
      return undefined;
    }

    return {
      turnId: finalTurnId,
      rootTurnId: finalRootTurnId,
      parentTurnId: resolvedParentTurnId || known?.parentTurnId
    };
  }

  private rememberTurnLineage(lineage: IHieTurnLineage): IHieTurnLineage {
    const normalized: IHieTurnLineage = {
      turnId: lineage.turnId,
      rootTurnId: lineage.rootTurnId || lineage.turnId,
      parentTurnId: lineage.parentTurnId
    };
    this.turnLineageById.set(normalized.turnId, normalized);
    return normalized;
  }
}

// ─── Singleton ──────────────────────────────────────────────────

export const hybridInteractionEngine = new HybridInteractionEngine();
