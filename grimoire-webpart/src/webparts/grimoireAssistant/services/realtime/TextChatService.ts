/**
 * TextChatService
 * HTTP chat completions client for the text-only path.
 * When voice is NOT connected, text messages flow through this service
 * instead of the WebRTC data channel.
 *
 * Features:
 * - Streaming SSE response parsing
 * - Tool-call loop (max 10 iterations) with async tool support
 * - 429 rate-limit retry with backoff (Retry-After header or 5s default)
 * - Token optimization: tool result truncation and history summarization
 * - Conversation history management (capped at 30 messages)
 * - Same tool dispatch as voice path via shared handleFunctionCall()
 *
 * Follows the NanoService pattern for proxy HTTP calls.
 */

import type { IProxyConfig } from '../../store/useGrimoireStore';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import type {
  IBlock,
  IDocumentItem,
  IDocumentLibraryData,
  IFormFieldDefinition,
  ISearchResult,
  ISearchResultsData
} from '../../models/IBlock';
import type { PersonalityMode } from '../avatar/PersonalityEngine';
import type { VisageMode } from '../avatar/FaceTemplateData';
import type { IUserContext } from '../context/ContextService';
import { normalizeError } from '../utils/errorUtils';
import { getSystemPrompt } from './SystemPrompt';
import type { IPromptConfig } from './SystemPrompt';
import { wrapToolResult } from './PromptSafety';
import { getTools } from './ToolRegistry';
import { logService } from '../logging/LogService';
import {
  getToolCompletionAckFromOutput
} from '../tools/ToolCompletionAcknowledgment';
import {
  createAbortError,
  resolveImmediateCompletionAck,
  sanitizeAssistantText,
  shouldSuppressDisplayToolCall,
  sleepWithSignal,
  summarizeToolResult,
  truncateToolResult,
  withTimeout
} from './TextChatHelpers';
import {
  buildFirstTurnRoutingLogDetail,
  classifyFirstTurnRouting,
  classifyAssistantFirstTurnOutcome,
  getForcedFirstToolArgs,
  getForcedFirstToolName,
  getObservedFirstToolName,
  hasActionableSearchIntent,
  hasExplicitExternalWebHint,
  hasInternalEnterpriseContentHint,
  type IFirstTurnRoutingObservation
} from './IntentRoutingPolicy';
import { classifyExplicitPersonalOneDriveIntent } from './ExplicitPersonalOneDriveIntent';
import { hasShareableSessionContent } from '../sharing/SessionShareFormatter';
import { isHiePromptMessage } from '../hie/HiePromptProtocol';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import {
  deriveMcpTargetContextFromHie,
  deriveMcpTargetContextFromUnknown,
  mergeMcpTargetContexts
} from '../mcp/McpTargetContext';
import type { IMcpTargetContext } from '../mcp/McpTargetContext';
import {
  executeCompoundWorkflowPlan,
  planCompoundWorkflow,
  type ICompoundWorkflowPlan
} from './CompoundWorkflowExecutor';
import {
  looksLikeMailDiscussionReplyAllRequest,
  resolveMailDiscussionReplyAllPlan
} from './MailDiscussionReplyAllWorkflow';
import {
  buildComposeStaticArgsFromScope,
  resolveContextualComposeScope,
  type IResolvedContextualComposeScope
} from '../sharing/ContextualComposeScope';

// ─── Types ────────────────────────────────────────────────────

interface IChatMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content?: string;
  tool_calls?: IToolCall[];
  tool_call_id?: string;
}

interface IToolCall {
  id: string;
  type: 'function';
  function: {
    name: string;
    arguments: string;
  };
}

export interface ITextChatCallbacks {
  onToken: (chunk: string) => void;
  onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string | Promise<string>;
  onComplete: (fullText: string) => void;
  onError: (error: string) => void;
  onRateLimitRetry?: (info: ITextChatRateLimitInfo) => void;
  onRateLimitResolved?: (info: ITextChatRateLimitInfo) => void;
  onRateLimitExhausted?: (info: ITextChatRateLimitInfo) => void;
}

export type TextChatRateLimitHeaderSource =
  | 'retry-after-ms'
  | 'x-ms-retry-after-ms'
  | 'retry-after-seconds'
  | 'retry-after-date'
  | 'fallback-exponential';

export interface ITextChatRateLimitInfo {
  attempt: number;
  maxRetries: number;
  delayMs: number;
  headerSource: TextChatRateLimitHeaderSource;
  status: 'retrying' | 'resolved' | 'exhausted';
}

// ─── Service ──────────────────────────────────────────────────

const MAX_MESSAGES = 30;
const MAX_TOOL_ITERATIONS = 10;
const MAX_NO_PROGRESS_ITERATIONS = 3;
const TOOL_TIMEOUT_MS = 60000;
const MAX_429_RETRIES = 3;
const DEFAULT_RETRY_DELAY_MS = 5000;
const MAX_RETRY_DELAY_MS = 30000;

/** Tools that already render visual data blocks — used to suppress duplicate display calls */
const VISUAL_DATA_TOOLS: ReadonlySet<string> = new Set([
  'research_public_web',
  'search_emails', 'search_sharepoint', 'search_people', 'search_sites',
  'browse_document_library', 'show_file_details', 'show_site_info', 'show_list_items',
  'use_m365_capability', 'call_mcp_tool', 'get_recent_documents', 'get_trending_documents',
  'get_my_profile', 'show_permissions', 'show_activity_feed',
  'show_chart', 'show_selection_list', 'recall_notes', 'list_m365_servers'
]);
const NON_PROGRESS_TOOLS: ReadonlySet<string> = new Set([
  'set_expression'
]);
const SHAREPOINT_COLUMN_TYPE_OPTIONS: ReadonlyArray<{ key: string; text: string }> = [
  { key: 'Text', text: 'Text' },
  { key: 'Multiple lines of text', text: 'Multiple lines of text' },
  { key: 'Number', text: 'Number' },
  { key: 'Yes/No', text: 'Yes/No' },
  { key: 'Date and time', text: 'Date and time' },
  { key: 'Person or Group', text: 'Person or Group' },
  { key: 'Link or Picture', text: 'Link or Picture' },
  { key: 'Choice', text: 'Choice' }
];
/** Tool results older than this many messages get summarized to save tokens */
const TOOL_RESULT_SUMMARY_THRESHOLD = 10;
type AbortReason = 'cancel' | 'timeout';

interface IForcedToolOverride {
  toolName: string;
  toolArgs: Record<string, unknown>;
}

interface IResolvedToolExecution {
  toolName: string;
  args: Record<string, unknown>;
}

import {
  type ContextualVisibleAction,
  type ContextualMutationAction,
  type ContextualContainerAction,
  type IContextualContainerTarget,
  type IContextualVisibleItemMatch,
  type IContextualComposeOverrideDetails,
  CONTEXTUAL_COMPOSE_ACTION_HINTS,
  CONTEXTUAL_EMAIL_HINTS,
  CONTEXTUAL_TEAMS_HINTS,
  CONTEXTUAL_CHANNEL_HINTS,
  CONTEXTUAL_CHAT_HINTS,
  CONTEXTUAL_SHARE_TARGET_HINTS,
  CONTEXTUAL_SUMMARIZE_HINTS,
  CONTEXTUAL_PREVIEW_HINTS,
  CONTEXTUAL_PERMISSION_HINTS,
  CONTEXTUAL_RENAME_HINTS,
  CONTEXTUAL_COLUMN_CREATE_HINTS,
  CONTEXTUAL_COLUMN_NOUN_HINTS,
  CONTEXTUAL_SELECTED_REFERENCE_HINTS,
  CONTEXTUAL_RENAME_CLARIFICATION_HINTS,
  CONTEXTUAL_LIST_DISCOVERY_HINTS,
  CONTEXTUAL_LIST_CONTENT_HINTS,
  CONTEXTUAL_LIBRARY_HINTS,
  WHICH_ONE_PROMPT_HINTS,
  ORDINAL_INDEX_PATTERNS,
  matchesAny,
  normalizeVisibleContextText,
  normalizeVisibleItemTitleText,
  blockDataAsSiteInfo
} from './TextChatContextualOverrides';

export class TextChatService {
  private messages: IChatMessage[] = [];
  private proxyConfig: IProxyConfig;
  private abortController: AbortController | undefined;
  private abortReasons: WeakMap<AbortSignal, AbortReason> = new WeakMap();
  private toolsPayload: Array<{ type: 'function'; function: { name: string; description: string; parameters: unknown } }>;
  /** True while runToolLoop is active — prevents concurrent message array mutation */
  private processing: boolean = false;
  /** Context messages buffered while processing is true */
  private pendingContextMessages: Array<{
    text: string;
    triggerResponse: boolean;
    callbacks?: ITextChatCallbacks;
  }> = [];
  private lastContextFingerprint: string = '';
  private lastContextAt: number = 0;
  /** Monotonic id used to drop superseded sends before the first network request starts. */
  private sendSequence: number = 0;
  private buildCompoundWorkflowAcknowledgment(plan: ICompoundWorkflowPlan): string {
    const steps = plan.steps.map((step) => step.label.toLowerCase());
    if (steps.length === 0) {
      return 'I am starting a compound workflow.';
    }
    if (steps.length === 1) {
      return `Plan: ${steps[0]}.`;
    }
    if (steps.length === 2) {
      return `Plan: ${steps[0]}, then ${steps[1]}.`;
    }
    const last = steps[steps.length - 1];
    return `Plan: ${steps.slice(0, -1).join(', ')}, then ${last}.`;
  }

  constructor(
    proxyConfig: IProxyConfig,
    personality: PersonalityMode,
    userContext?: IUserContext,
    promptConfig?: IPromptConfig,
    visage: VisageMode = 'classic'
  ) {
    this.proxyConfig = proxyConfig;
    this.messages = [
      { role: 'system', content: getSystemPrompt(personality, visage, userContext, promptConfig) }
    ];
    this.toolsPayload = getTools({ avatarEnabled: promptConfig?.avatarEnabled }).map((t) => ({
      type: 'function' as const,
      function: {
        name: t.name,
        description: t.description,
        parameters: t.parameters
      }
    }));
  }

  /**
   * Send a text message and stream the response.
   * Handles tool-call loops automatically, awaiting async tool results.
   */
  public async send(text: string, callbacks: ITextChatCallbacks): Promise<void> {
    // Sanitize history from any previous failed send() before adding new message
    this.sanitizeHistory();
    const sendId = ++this.sendSequence;
    const compoundWorkflowPlan = resolveMailDiscussionReplyAllPlan(text) || await planCompoundWorkflow(text, this.proxyConfig);
    const forcedToolOverride = compoundWorkflowPlan ? undefined : this.getForcedToolOverride(text);
    const routingObservation = !compoundWorkflowPlan && !forcedToolOverride && this.shouldApplyFirstTurnRouting()
      ? await classifyFirstTurnRouting(text, this.proxyConfig)
      : undefined;
    if (sendId !== this.sendSequence) {
      logService.debug('llm', 'TextChatService: superseded before routing completed');
      return;
    }

    // Append user message
    this.messages.push({ role: 'user', content: text });
    this.trimHistory();

    if (compoundWorkflowPlan) {
      await this.runCompoundWorkflow(compoundWorkflowPlan, callbacks);
      return;
    }

    await this.runToolLoop(callbacks, routingObservation, forcedToolOverride);
  }

  /**
   * Inject a context message into the conversation history.
   * If triggerResponse is true, runs a full response cycle (tool loop + streaming).
   * If false, the message is appended silently — the LLM sees it on the next user message.
   *
   * When a tool loop is active, messages are buffered to prevent corrupting
   * the assistant(tool_calls) → tool message ordering required by the API.
   */
  public async injectContextMessage(
    text: string,
    triggerResponse: boolean,
    callbacks?: ITextChatCallbacks
  ): Promise<void> {
    const fingerprint = `${triggerResponse ? '1' : '0'}:${text}`;
    const now = Date.now();
    if (this.lastContextFingerprint === fingerprint && (now - this.lastContextAt) < 1200) {
      logService.debug('llm', 'TextChatService: duplicate context ignored');
      return;
    }
    this.lastContextFingerprint = fingerprint;
    this.lastContextAt = now;

    if (this.processing) {
      this.pendingContextMessages.push({ text, triggerResponse, callbacks });
      logService.debug('llm', 'TextChatService: buffered context message (processing active)');
      return;
    }

    // Sanitize before adding — prevents orphaned tool messages from crashing the API
    if (triggerResponse) {
      this.sanitizeHistory();
    }

    this.messages.push({ role: 'user', content: text });
    this.trimHistory();

    if (triggerResponse && callbacks) {
      await this.runToolLoop(callbacks);
    }
  }

  /**
   * Cancel any in-flight request.
   */
  public cancel(): void {
    if (this.abortController) {
      this.abortReasons.set(this.abortController.signal, 'cancel');
      this.abortController.abort();
    }
  }

  /**
   * Reset conversation history (keeps system prompt).
   */
  public reset(): void {
    const systemMsg = this.messages[0];
    this.messages = systemMsg ? [systemMsg] : [];
    this.pendingContextMessages = [];
    this.processing = false;
  }

  // ─── Private ──────────────────────────────────────────────────

  private ensureRequestStillActive(activeController: AbortController): void {
    if (this.abortController !== activeController || activeController.signal.aborted) {
      throw createAbortError();
    }
  }

  private async runToolLoop(
    callbacks: ITextChatCallbacks,
    routingObservation?: IFirstTurnRoutingObservation,
    forcedToolOverride?: IForcedToolOverride
  ): Promise<void> {
    const latestMessage = this.messages[this.messages.length - 1];
    const pendingMailDiscussionPlan = resolveMailDiscussionReplyAllPlan(undefined, {
      allowPending: true,
      latestUserMessageText: latestMessage?.role === 'user' && typeof latestMessage.content === 'string'
        ? latestMessage.content
        : undefined,
      requireHiePromptForPending: true
    });
    if (pendingMailDiscussionPlan) {
      await this.runCompoundWorkflow(pendingMailDiscussionPlan, callbacks, { announcePlan: false });
      return;
    }

    // Cancel any in-flight request
    if (this.abortController) {
      this.abortReasons.set(this.abortController.signal, 'cancel');
      this.abortController.abort();
    }
    this.abortController = new AbortController();
    const activeController = this.abortController;
    this.processing = true;

    // Track message count for rollback on failure
    const messageCountBefore = this.messages.length;

    // Layer 3B: track failed tool calls to prevent identical retries
    const failedCalls = new Set<string>();
    let forcedFirstToolName = forcedToolOverride?.toolName || getForcedFirstToolName(routingObservation);
    const forcedFirstToolArgs = forcedToolOverride?.toolArgs || getForcedFirstToolArgs(routingObservation);
    const routedFirstToolName = forcedToolOverride?.toolName || routingObservation?.expectedToolName;

    // Layer 4: suppress re-display attempts — when a visual data tool already pushed a block,
    // block show_markdown/show_info_card calls that would create duplicates with fabricated data
    let visualDataToolSucceeded = false;

    try {
      let iterations = 0;
      let needsResponse = true;
      let consecutiveNoProgressIterations = 0;

      while (needsResponse && iterations < MAX_TOOL_ITERATIONS) {
        iterations++;
        const result = await this.postCompletion(callbacks, forcedFirstToolName);
        this.ensureRequestStillActive(activeController);
        forcedFirstToolName = undefined;

        if (result.toolCalls && result.toolCalls.length > 0) {
          const localCompletionAcks: Array<{ toolName: string; text: string }> = [];
          const observedToolName = routingObservation
            ? getObservedFirstToolName(result.toolCalls.map((tc) => tc.function.name))
            : undefined;
          if (routingObservation && observedToolName) {
            this.logFirstTurnRoutingOutcome('tool_call', routingObservation, observedToolName);
            routingObservation = undefined;
          }

          // Append assistant message with tool_calls
          this.messages.push({
            role: 'assistant',
            content: result.text || undefined,
            tool_calls: result.toolCalls
          });

          const executedToolNames: string[] = [];
          // Execute each tool call and append results (sequentially to await async tools)
          for (const tc of result.toolCalls) {
            let args: Record<string, unknown> = {};
            try {
              args = JSON.parse(tc.function.arguments) as Record<string, unknown>;
            } catch {
              args = {};
            }

            if (forcedFirstToolArgs && routedFirstToolName && tc.function.name === routedFirstToolName) {
              args = { ...args, ...forcedFirstToolArgs };
            }

            const latestUserText = this.getLatestUserMessageText();
            const resolvedTool = this.resolveToolExecution(tc.function.name, args, latestUserText);
            const executedToolName = resolvedTool.toolName;
            const executedArgs = resolvedTool.args;
            const callKey = `${executedToolName}:${JSON.stringify(executedArgs)}`;
            executedToolNames.push(executedToolName);

            // Layer 3B: skip duplicate failed calls
            if (failedCalls.has(callKey)) {
              logService.warning('llm', `TextChatService: blocking duplicate failed call: ${executedToolName}`);
              this.messages.push({
                role: 'tool',
                tool_call_id: tc.id,
                content: wrapToolResult(
                  executedToolName,
                  JSON.stringify({
                    success: false,
                    duplicateCallBlocked: true,
                    error: `This exact call to "${executedToolName}" already failed with the same arguments. Do not retry. Explain the issue to the user or try a different approach.`
                  })
                )
              });
              continue;
            }

            // Layer 4: suppress re-display when a prior tool already rendered visual data
            if (shouldSuppressDisplayToolCall(executedToolName, visualDataToolSucceeded, latestUserText)) {
              logService.info('llm', `TextChatService: suppressed "${executedToolName}" — data already displayed by a prior tool`);
              this.messages.push({
                role: 'tool',
                tool_call_id: tc.id,
                content: wrapToolResult(
                  executedToolName,
                  JSON.stringify({
                    success: true,
                    alreadyDisplayed: true,
                    message: 'Data is already displayed to the user in the action panel.',
                    advice: 'Do not re-display it. Summarize or comment on the data instead.'
                  })
                )
              });
              continue;
            }

            let output: string;
            try {
              const rawOutput = callbacks.onFunctionCall(tc.id, executedToolName, executedArgs);
              const TIMEOUT_SENTINEL = '__TOOL_TIMEOUT__';
              const timeoutFallback = JSON.stringify({ success: false, error: `Tool "${executedToolName}" timed out after ${TOOL_TIMEOUT_MS / 1000}s` });
              output = await withTimeout(
                Promise.resolve(rawOutput),
                TOOL_TIMEOUT_MS,
                TIMEOUT_SENTINEL
              );
              if (output === TIMEOUT_SENTINEL) {
                logService.warning('llm', `TextChatService: tool "${executedToolName}" timed out after ${TOOL_TIMEOUT_MS / 1000}s (MCP call may still complete in background)`);
                output = timeoutFallback;
              }
            } catch (toolErr) {
              const normalizedToolErr = normalizeError(toolErr, 'Tool execution failed');
              output = JSON.stringify({ success: false, error: normalizedToolErr.message });
            }
            this.ensureRequestStillActive(activeController);

            // Track failures for retry prevention
            try {
              const parsed = JSON.parse(output);
              if (parsed && parsed.success === false) {
                failedCalls.add(callKey);
              }
            } catch { /* not JSON — treat as success */ }

            // Layer 4: track visual data tool success
            if (VISUAL_DATA_TOOLS.has(executedToolName) && !failedCalls.has(callKey)) {
              visualDataToolSucceeded = true;
            }

            this.messages.push({
              role: 'tool',
              tool_call_id: tc.id,
              content: wrapToolResult(executedToolName, truncateToolResult(output, executedToolName))
            });

            const completionAck = getToolCompletionAckFromOutput(
              executedToolName,
              output,
              useGrimoireStore.getState().conversationLanguage
            );
            if (completionAck) {
              localCompletionAcks.push({ toolName: executedToolName, text: completionAck });
            }
          }

          const immediateCompletionAck = resolveImmediateCompletionAck(
            executedToolNames,
            localCompletionAcks,
            this.getLatestUserMessageText()
          );

          const madeUserVisibleProgress = executedToolNames.some((toolName) => !NON_PROGRESS_TOOLS.has(toolName))
            || localCompletionAcks.length > 0;
          if (madeUserVisibleProgress) {
            consecutiveNoProgressIterations = 0;
          } else {
            consecutiveNoProgressIterations += 1;
          }

          if (immediateCompletionAck) {
            this.messages.push({ role: 'assistant', content: immediateCompletionAck });
            callbacks.onComplete(immediateCompletionAck);
            needsResponse = false;
            continue;
          }

          if (consecutiveNoProgressIterations >= MAX_NO_PROGRESS_ITERATIONS) {
            const fallbackMessage = this.buildNoProgressGuardrailMessage(this.getLatestUserMessageText());
            logService.warning('llm', `TextChatService: breaking no-progress tool loop after ${consecutiveNoProgressIterations} iterations`);
            this.messages.push({ role: 'assistant', content: fallbackMessage });
            callbacks.onComplete(fallbackMessage);
            needsResponse = false;
            continue;
          }
          // Loop again to get the model's next response after tool results
        } else {
          // No tool calls — final response
          const sanitizedText = sanitizeAssistantText(result.text || '');
          if (result.text && !sanitizedText) {
            logService.warning('llm', 'TextChatService: suppressed internal tool payload leaked as assistant text');
          }
          if (routingObservation) {
            const outcome = classifyAssistantFirstTurnOutcome(sanitizedText);
            this.logFirstTurnRoutingOutcome(outcome, routingObservation);
            routingObservation = undefined;
          }
          if (sanitizedText) {
            this.messages.push({ role: 'assistant', content: sanitizedText });
          }
          callbacks.onComplete(sanitizedText);
          needsResponse = false;
        }
      }

      if (iterations >= MAX_TOOL_ITERATIONS) {
        logService.warning('llm', `TextChatService: hit max tool iterations (${MAX_TOOL_ITERATIONS})`);
        const fallbackMessage = this.buildNoProgressGuardrailMessage(this.getLatestUserMessageText());
        this.messages.push({ role: 'assistant', content: fallbackMessage });
        callbacks.onComplete(fallbackMessage);
      }
    } catch (err) {
      const normalizedErr = normalizeError(err, 'Text chat request failed');
      const superseded = this.abortController !== activeController;
      const abortReason = this.abortReasons.get(activeController.signal);

      if (!superseded) {
        // Roll back messages to avoid corrupted history
        this.messages.length = messageCountBefore;
        logService.debug('llm', `TextChatService: rolled back to ${messageCountBefore} messages after error`);
      }

      if (normalizedErr.name === 'AbortError') {
        if (abortReason === 'cancel') {
          logService.debug('llm', superseded
            ? 'TextChatService: superseded request aborted'
            : 'TextChatService: request cancelled');
          return;
        }
        if (!superseded) {
          logService.debug('llm', 'TextChatService: request timed out');
          callbacks.onError('Request timed out');
        }
        return;
      }

      if (superseded) {
        logService.debug('llm', `TextChatService: ignored error from superseded request: ${normalizedErr.message}`);
        return;
      }

      callbacks.onError(normalizedErr.message);
    } finally {
      if (this.abortController === activeController) {
        this.abortController = undefined;
        this.processing = false;
        this.flushPendingContextMessages();
      }
    }
  }

  private async runCompoundWorkflow(
    plan: ICompoundWorkflowPlan,
    callbacks: ITextChatCallbacks,
    options?: { announcePlan?: boolean }
  ): Promise<void> {
    if (this.abortController) {
      this.abortReasons.set(this.abortController.signal, 'cancel');
      this.abortController.abort();
      this.abortController = undefined;
    }
    this.processing = true;

    try {
      if (options?.announcePlan !== false) {
        callbacks.onToken(this.buildCompoundWorkflowAcknowledgment(plan));
      }
      const assistantText = await executeCompoundWorkflowPlan(plan, {
        onFunctionCall: callbacks.onFunctionCall
      });
      this.messages.push({ role: 'assistant', content: assistantText });
      this.trimHistory();
      callbacks.onComplete(assistantText);
    } catch (err) {
      const normalizedErr = normalizeError(err, 'Compound workflow failed');
      callbacks.onError(normalizedErr.message);
    } finally {
      this.processing = false;
      this.flushPendingContextMessages();
    }
  }

  private flushPendingContextMessages(): void {
    if (this.pendingContextMessages.length === 0) return;

    const pending = this.pendingContextMessages.splice(0);

    // Consolidate: keep only the LAST silent message and the LAST triggered message.
    // Earlier messages are superseded by later state.
    let lastSilent: string | undefined;
    let lastTriggered: { text: string; callbacks: ITextChatCallbacks } | undefined;

    pending.forEach((msg) => {
      if (msg.triggerResponse && msg.callbacks) {
        lastTriggered = { text: msg.text, callbacks: msg.callbacks };
      } else {
        lastSilent = msg.text;
      }
    });

    if (lastSilent) {
      this.messages.push({ role: 'user', content: lastSilent });
    }
    if (lastTriggered) {
      this.messages.push({ role: 'user', content: lastTriggered.text });
    }
    this.trimHistory();

    const kept = (lastSilent ? 1 : 0) + (lastTriggered ? 1 : 0);
    logService.debug('llm', `TextChatService: flushed ${pending.length} buffered -> ${kept} consolidated message(s)`);
    if (lastTriggered) {
      this.runToolLoop(lastTriggered.callbacks).catch((err: unknown) => {
        const normalizedErr = normalizeError(err, 'Buffered context response failed');
        logService.error('llm', `TextChatService: buffered context response failed: ${normalizedErr.message}`);
      });
    }
  }

  private async postCompletion(
    callbacks: ITextChatCallbacks,
    forcedToolName?: string
  ): Promise<{ text: string; toolCalls: IToolCall[] }> {
    const { proxyUrl, proxyApiKey, deployment, apiVersion } = this.proxyConfig;
    const controller = this.abortController;
    if (!controller) {
      throw new Error('No active abort controller for text chat request');
    }
    const url = `${proxyUrl}/${this.proxyConfig.backend}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
    this.sanitizeHistory();
    if (controller.signal.aborted) {
      throw createAbortError();
    }

    const payload: Record<string, unknown> = {
      messages: this.messages,
      stream: true,
      tools: this.toolsPayload,
      tool_choice: forcedToolName
        ? {
            type: 'function',
            function: {
              name: forcedToolName
            }
          }
        : 'auto'
    };
    const body = JSON.stringify(payload);
    let retryCount = 0;
    let lastRateLimitInfo: ITextChatRateLimitInfo | undefined;

    for (let attempt = 0; attempt <= MAX_429_RETRIES; attempt++) {
      // Abort fetch if backend doesn't respond within TOOL_TIMEOUT_MS
      const fetchTimer = setTimeout(() => {
        this.abortReasons.set(controller.signal, 'timeout');
        controller.abort();
      }, TOOL_TIMEOUT_MS);

      let response: Response;
      try {
        response = await fetch(url, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-functions-key': proxyApiKey
          },
          body,
          signal: controller.signal
        });
      } catch (err) {
        clearTimeout(fetchTimer);
        throw err;
      }
      clearTimeout(fetchTimer);

      if (response.status === 429) {
        const retryMeta = resolveRateLimitDelay(response.headers, retryCount);
        if (retryCount < MAX_429_RETRIES) {
          retryCount += 1;
          lastRateLimitInfo = {
            attempt: retryCount,
            maxRetries: MAX_429_RETRIES,
            delayMs: retryMeta.delayMs,
            headerSource: retryMeta.headerSource,
            status: 'retrying'
          };
          callbacks.onRateLimitRetry?.(lastRateLimitInfo);
          logService.warning(
            'llm',
            `TextChatService: 429 rate-limited, retry ${retryCount}/${MAX_429_RETRIES} after ${retryMeta.delayMs}ms`,
            `source=${retryMeta.headerSource}`
          );
          await sleepWithSignal(controller.signal, retryMeta.delayMs);
          continue;
        }

        callbacks.onRateLimitExhausted?.({
          attempt: retryCount,
          maxRetries: MAX_429_RETRIES,
          delayMs: retryMeta.delayMs,
          headerSource: retryMeta.headerSource,
          status: 'exhausted'
        });
        throw new Error('Chat completion failed: rate limited after all retries');
      }

      if (lastRateLimitInfo) {
        callbacks.onRateLimitResolved?.({
          ...lastRateLimitInfo,
          status: 'resolved'
        });
        lastRateLimitInfo = undefined;
      }

      if (!response.ok) {
        const respBody = await response.text();
        throw new Error(`Chat completion failed (${response.status}): ${respBody.slice(0, 200)}`);
      }

      return this.parseSSEStream(response, callbacks);
    }

    throw new Error('Chat completion failed: rate limited after all retries');
  }

  private async parseSSEStream(
    response: Response,
    callbacks: ITextChatCallbacks
  ): Promise<{ text: string; toolCalls: IToolCall[] }> {
    const reader = response.body?.getReader();
    if (!reader) {
      throw new Error('No response body');
    }

    const decoder = new TextDecoder();
    let buffer = '';
    let fullText = '';
    const toolCallMap = new Map<number, IToolCall>();

    try {
      let done = false;
      while (!done) {
        const readResult = await reader.read();
        done = readResult.done;
        const value = readResult.value;
        if (done || !value) break;

        buffer += decoder.decode(value, { stream: true });

        // Split on newlines to process SSE events
        const lines = buffer.split('\n');
        // Keep the last (potentially incomplete) line in the buffer
        buffer = lines.pop() || '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || !trimmed.startsWith('data: ')) continue;

          const data = trimmed.slice(6); // Remove 'data: ' prefix
          if (data === '[DONE]') continue;

          let parsed: {
            choices?: Array<{
              delta?: {
                content?: string;
                tool_calls?: Array<{
                  index: number;
                  id?: string;
                  type?: string;
                  function?: {
                    name?: string;
                    arguments?: string;
                  };
                }>;
              };
              finish_reason?: string;
            }>;
          };

          try {
            parsed = JSON.parse(data);
          } catch {
            continue;
          }

          const choice = parsed.choices?.[0];
          if (!choice) continue;

          const delta = choice.delta;
          if (!delta) continue;

          // Content tokens
          if (delta.content) {
            fullText += delta.content;
            callbacks.onToken(delta.content);
          }

          // Tool call deltas
          if (delta.tool_calls) {
            for (const tc of delta.tool_calls) {
              const existing = toolCallMap.get(tc.index);
              if (!existing) {
                // First chunk for this tool call
                toolCallMap.set(tc.index, {
                  id: tc.id || '',
                  type: 'function',
                  function: {
                    name: tc.function?.name || '',
                    arguments: tc.function?.arguments || ''
                  }
                });
              } else {
                // Append arguments delta
                if (tc.function?.arguments) {
                  existing.function.arguments += tc.function.arguments;
                }
                if (tc.id) {
                  existing.id = tc.id;
                }
                if (tc.function?.name) {
                  existing.function.name = tc.function.name;
                }
              }
            }
          }
        }
      }
    } finally {
      reader.releaseLock();
    }

    // Collect tool calls in index order
    const toolCalls: IToolCall[] = [];
    const sortedKeys: number[] = [];
    toolCallMap.forEach((_v, k) => sortedKeys.push(k));
    sortedKeys.sort((a, b) => a - b);
    sortedKeys.forEach((k) => {
      const tc = toolCallMap.get(k);
      if (tc) toolCalls.push(tc);
    });

    return { text: fullText, toolCalls };
  }

  /**
   * Remove invalid assistant/tool sequences from history. This can happen if a
   * previous send() failed or was superseded mid-way through tool processing.
   */
  private sanitizeHistory(): void {
    if (this.messages.length <= 1) {
      return;
    }

    const sanitized: IChatMessage[] = [];
    let removedCount = 0;

    for (let i = 0; i < this.messages.length; i++) {
      const msg = this.messages[i];
      if (msg.role === 'assistant' && msg.tool_calls && msg.tool_calls.length > 0) {
        const expectedIds = new Set<string>();
        msg.tool_calls.forEach((tc) => expectedIds.add(tc.id));

        let j = i + 1;
        while (j < this.messages.length && this.messages[j].role === 'tool') {
          const toolCallId = this.messages[j].tool_call_id;
          if (toolCallId) {
            expectedIds.delete(toolCallId);
          }
          j++;
        }

        if (expectedIds.size > 0) {
          removedCount += (j - i);
          i = j - 1;
          continue;
        }

        sanitized.push(msg);
        for (let k = i + 1; k < j; k++) {
          sanitized.push(this.messages[k]);
        }
        i = j - 1;
        continue;
      }

      if (msg.role === 'tool') {
        removedCount += 1;
        continue;
      }

      sanitized.push(msg);
    }

    if (removedCount > 0) {
      this.messages = sanitized;
      logService.warning('llm', `TextChatService: sanitized ${removedCount} dangling messages from history`);
    }
  }

  private getContextualComposeOverride(text: string): IForcedToolOverride | undefined {
    const normalizedText = text.trim().toLowerCase();
    if (!normalizedText || isHiePromptMessage(normalizedText)) {
      return undefined;
    }

    const state = useGrimoireStore.getState();
    const hasVisibleShareableBlocks = state.blocks.some((block) => block.type !== 'form' && block.type !== 'confirmation-dialog');
    if (!hasVisibleShareableBlocks || !hasShareableSessionContent(state.blocks, state.transcript)) {
      return undefined;
    }

    const tokenCount = normalizedText.split(/\s+/).filter(Boolean).length;
    const hasComposeVerb = matchesAny(normalizedText, CONTEXTUAL_COMPOSE_ACTION_HINTS);
    const hasTargetOnlyHint = tokenCount <= 6 && (
      matchesAny(normalizedText, CONTEXTUAL_EMAIL_HINTS)
      || matchesAny(normalizedText, CONTEXTUAL_TEAMS_HINTS)
    );
    if (!hasComposeVerb && !hasTargetOnlyHint) {
      return undefined;
    }

    const isMailDiscussionRequest = looksLikeMailDiscussionReplyAllRequest(normalizedText);
    const looksContextual = isMailDiscussionRequest
      || tokenCount <= 10
      || matchesAny(normalizedText, CONTEXTUAL_SHARE_TARGET_HINTS)
      || matchesAny(normalizedText, CONTEXTUAL_SUMMARIZE_HINTS);
    if (!looksContextual) {
      return undefined;
    }

    const composeScope = resolveContextualComposeScope({
      text: normalizedText,
      blocks: state.blocks,
      activeBlockId: state.activeActionBlockId,
      selectedActionIndices: state.selectedActionIndices
    });
    const visibleItemMatch = composeScope
      ? undefined
      : this.resolveContextualVisibleItemMatch(normalizedText);

    if (isMailDiscussionRequest) {
      return undefined;
    }

    if (matchesAny(normalizedText, CONTEXTUAL_EMAIL_HINTS)) {
      const composeDetails = this.buildContextualComposeOverrideDetails(
        composeScope,
        visibleItemMatch,
        'Pre-filled with the currently visible content. Add recipients and send when ready.',
        'Pre-filled to send the selected item. Add recipients and send when ready.',
        'Pre-filled to send the selected items. Add recipients and send when ready.'
      );
      logService.info('llm', 'TextChatService: contextual compose override → email-compose');
      return {
        toolName: 'show_compose_form',
        toolArgs: {
          preset: 'email-compose',
          title: 'Share by Email',
          ...(composeDetails.staticArgsJson ? { static_args_json: composeDetails.staticArgsJson } : {})
        }
      };
    }

    if (matchesAny(normalizedText, CONTEXTUAL_TEAMS_HINTS) && matchesAny(normalizedText, CONTEXTUAL_CHANNEL_HINTS)) {
      const composeDetails = this.buildContextualComposeOverrideDetails(
        composeScope,
        visibleItemMatch,
        'Choose the destination channel, review the message, then send.',
        'Choose the destination channel, review the selected item, then send.',
        'Choose the destination channel, review the selected items, then send.'
      );
      logService.info('llm', 'TextChatService: contextual compose override → share-teams-channel');
      return {
        toolName: 'show_compose_form',
        toolArgs: {
          preset: 'share-teams-channel',
          title: 'Post to a Teams Channel',
          description: composeDetails.description,
          ...(composeDetails.staticArgsJson ? { static_args_json: composeDetails.staticArgsJson } : {})
        }
      };
    }

    if (matchesAny(normalizedText, CONTEXTUAL_CHAT_HINTS)) {
      const composeDetails = this.buildContextualComposeOverrideDetails(
        composeScope,
        visibleItemMatch,
        'Choose recipients, review the message, then send.',
        'Choose recipients, review the selected item, then send.',
        'Choose recipients, review the selected items, then send.'
      );
      logService.info('llm', 'TextChatService: contextual compose override → share-teams-chat');
      return {
        toolName: 'show_compose_form',
        toolArgs: {
          preset: 'share-teams-chat',
          title: 'Share to Teams Chat',
          description: composeDetails.description,
          ...(composeDetails.staticArgsJson ? { static_args_json: composeDetails.staticArgsJson } : {})
        }
      };
    }

    if (matchesAny(normalizedText, CONTEXTUAL_TEAMS_HINTS)) {
      const composeDetails = this.buildContextualComposeOverrideDetails(
        composeScope,
        visibleItemMatch,
        'Choose the destination channel, review the message, then send.',
        'Choose the destination channel, review the selected item, then send.',
        'Choose the destination channel, review the selected items, then send.'
      );
      logService.info('llm', 'TextChatService: contextual compose override → share-teams-channel');
      return {
        toolName: 'show_compose_form',
        toolArgs: {
          preset: 'share-teams-channel',
          title: 'Post to a Teams Channel',
          description: composeDetails.description,
          ...(composeDetails.staticArgsJson ? { static_args_json: composeDetails.staticArgsJson } : {})
        }
      };
    }

    return undefined;
  }

  private getContextualSharePointColumnOverride(text: string): IForcedToolOverride | undefined {
    const normalizedText = normalizeVisibleContextText(text);
    const loweredText = normalizedText.toLowerCase();
    if (!loweredText || isHiePromptMessage(loweredText)) {
      return undefined;
    }

    if (
      !matchesAny(loweredText, CONTEXTUAL_COLUMN_CREATE_HINTS)
      || !matchesAny(loweredText, CONTEXTUAL_COLUMN_NOUN_HINTS)
    ) {
      return undefined;
    }

    const targetContext = this.resolveContextualSharePointColumnTarget(loweredText);
    const hasListTarget = !!(targetContext?.listId || targetContext?.listUrl || targetContext?.listName);
    const hasSiteTarget = !!(targetContext?.siteId || targetContext?.siteUrl || targetContext?.siteName);
    if (!hasListTarget || !hasSiteTarget) {
      return undefined;
    }

    const inferredDisplayName = this.inferSharePointColumnDisplayName(normalizedText);
    const inferredType = this.inferSharePointColumnTypeLabel(loweredText);
    const prefill: Record<string, string> = {};
    if (inferredType) {
      prefill.columnType = inferredType;
    }
    if (inferredDisplayName) {
      prefill.displayName = inferredDisplayName;
      prefill.columnName = inferredDisplayName;
    }

    const staticArgs: Record<string, unknown> = {};
    const copyContextKey = (key: 'siteId' | 'siteUrl' | 'siteName' | 'listId' | 'listUrl' | 'listName'): void => {
      const value = targetContext?.[key];
      if (value) {
        staticArgs[key] = value;
      }
    };
    copyContextKey('siteId');
    copyContextKey('siteUrl');
    copyContextKey('siteName');
    copyContextKey('listId');
    copyContextKey('listUrl');
    copyContextKey('listName');

    const formFields: IFormFieldDefinition[] = [
      {
        key: 'columnType',
        label: 'Column type',
        type: 'dropdown',
        required: true,
        placeholder: 'Select a column type',
        options: [...SHAREPOINT_COLUMN_TYPE_OPTIONS]
      },
      {
        key: 'displayName',
        label: 'Display name',
        type: 'text',
        required: true,
        placeholder: 'User-facing column name'
      },
      {
        key: 'columnName',
        label: 'Internal name',
        type: 'text',
        required: false,
        placeholder: 'Optional; generated from display name if left blank'
      },
      {
        key: 'description',
        label: 'Description',
        type: 'textarea',
        required: false,
        placeholder: 'Optional description for this column',
        rows: 3
      },
      {
        key: 'choiceValues',
        label: 'Choices',
        type: 'textarea',
        required: true,
        placeholder: 'One option per line',
        rows: 4,
        visibleWhen: {
          fieldKey: 'columnType',
          equals: 'Choice'
        }
      },
      {
        key: 'required',
        label: 'Required',
        type: 'toggle',
        required: false,
        group: 'Settings',
        width: 'half'
      },
      {
        key: 'enforceUniqueValues',
        label: 'Unique values',
        type: 'toggle',
        required: false,
        group: 'Settings',
        width: 'half'
      },
      {
        key: 'hidden',
        label: 'Hidden',
        type: 'toggle',
        required: false,
        group: 'Settings',
        width: 'half'
      },
      {
        key: 'indexed',
        label: 'Indexed',
        type: 'toggle',
        required: false,
        group: 'Settings',
        width: 'half'
      },
      {
        key: 'readOnly',
        label: 'Read-only',
        type: 'toggle',
        required: false,
        group: 'Settings',
        width: 'half'
      }
    ];

    const listLabel = targetContext?.listName || targetContext?.listId || 'the selected list';
    const siteLabel = targetContext?.siteName || targetContext?.siteUrl;
    const description = siteLabel
      ? `Review the new column settings for list ${listLabel} in ${siteLabel}, then submit to create it.`
      : `Review the new column settings for list ${listLabel}, then submit to create it.`;

    logService.info('llm', `TextChatService: contextual SharePoint column override → createListColumn (${listLabel})`);
    return {
      toolName: 'show_compose_form',
      toolArgs: {
        preset: 'generic',
        title: targetContext?.listName ? `Add Column to ${targetContext.listName}` : 'Add SharePoint Column',
        description,
        ...(Object.keys(prefill).length > 0 ? { prefill_json: JSON.stringify(prefill) } : {}),
        custom_fields_json: JSON.stringify(formFields),
        custom_target_json: JSON.stringify({
          toolName: 'createListColumn',
          serverId: 'mcp_SharePointListsTools',
          staticArgs,
          fieldToParamMap: {
            columnType: 'columnType',
            displayName: 'displayName',
            columnName: 'columnName',
            description: 'description',
            choiceValues: 'choiceValues',
            required: 'required',
            enforceUniqueValues: 'enforceUniqueValues',
            hidden: 'hidden',
            indexed: 'indexed',
            readOnly: 'readOnly'
          },
          targetContext
        })
      }
    };
  }

  private getForcedToolOverride(text: string): IForcedToolOverride | undefined {
    const composeOverride = this.getContextualComposeOverride(text);
    if (composeOverride) {
      return composeOverride;
    }

    const sharePointColumnOverride = this.getContextualSharePointColumnOverride(text);
    if (sharePointColumnOverride) {
      return sharePointColumnOverride;
    }

    const explicitPersonalOneDriveOverride = this.getExplicitPersonalOneDriveOverride(text);
    if (explicitPersonalOneDriveOverride) {
      return explicitPersonalOneDriveOverride;
    }

    const mutationOverride = this.getContextualMutationOverride(text);
    if (mutationOverride) {
      return mutationOverride;
    }

    const containerOverride = this.getContextualContainerOverride(text);
    if (containerOverride) {
      return containerOverride;
    }

    return this.getContextualVisibleItemOverride(text);
  }

  private getExplicitPersonalOneDriveOverride(text: string): IForcedToolOverride | undefined {
    const normalizedText = text.trim();
    if (!normalizedText || isHiePromptMessage(normalizedText)) {
      return undefined;
    }

    const intent = classifyExplicitPersonalOneDriveIntent(normalizedText);
    if (!intent) {
      return undefined;
    }

    if (intent.kind === 'unsupported-filter') {
      const body = 'I can browse your personal OneDrive root and search it by file name through MCP, but the currently exposed ODSP tools do not support filtering your personal OneDrive by date, recency, or file type.';
      logService.info('llm', 'TextChatService: explicit personal OneDrive override → unsupported filter limitation');
      return {
        toolName: 'show_info_card',
        toolArgs: {
          heading: 'Personal OneDrive MCP Limitation',
          body,
          icon: 'Info'
        }
      };
    }

    if (intent.kind === 'search-by-name' && intent.searchQuery) {
      logService.info('llm', `TextChatService: explicit personal OneDrive override → findFileOrFolder (${intent.searchQuery})`);
      return {
        toolName: 'use_m365_capability',
        toolArgs: {
          tool_name: 'findFileOrFolder',
          arguments_json: JSON.stringify({
            searchQuery: intent.searchQuery,
            personalOneDrive: true
          }),
          server_hint: 'mcp_ODSPRemoteServer'
        }
      };
    }

    logService.info('llm', 'TextChatService: explicit personal OneDrive override → getFolderChildren');
    return {
      toolName: 'use_m365_capability',
      toolArgs: {
        tool_name: 'getFolderChildren',
        arguments_json: JSON.stringify({
          personalOneDrive: true
        }),
        server_hint: 'mcp_ODSPRemoteServer'
      }
    };
  }

  private getContextualMutationOverride(text: string): IForcedToolOverride | undefined {
    const normalizedText = normalizeVisibleContextText(text).toLowerCase();
    if (!normalizedText || isHiePromptMessage(normalizedText)) {
      return undefined;
    }

    const inferredAction = this.detectContextualMutationAction(normalizedText);
    if (!inferredAction) {
      return undefined;
    }

    const match = this.resolveContextualVisibleItemMatch(normalizedText);
    if (!match) {
      return undefined;
    }

    if (inferredAction === 'rename') {
      return this.buildContextualRenameOverride(match);
    }

    return undefined;
  }

  private getContextualContainerOverride(text: string): IForcedToolOverride | undefined {
    const normalizedText = normalizeVisibleContextText(text).toLowerCase();
    if (!normalizedText || isHiePromptMessage(normalizedText)) {
      return undefined;
    }

    const target = this.resolveContextualContainerTarget(normalizedText);
    const action = this.detectContextualContainerAction(normalizedText, target);
    if (!action) {
      return undefined;
    }

    if (action === 'list-lists') {
      const siteReference = target.siteName || this.extractExplicitSiteReference(normalizedText);
      if (!siteReference) {
        return undefined;
      }

      logService.info('llm', `TextChatService: contextual container override → listLists (${siteReference})`);
      return {
        toolName: 'use_m365_capability',
        toolArgs: {
          tool_name: 'listLists',
          arguments_json: JSON.stringify({ siteName: siteReference }),
          server_hint: 'mcp_SharePointListsTools'
        }
      };
    }

    if (!target.siteUrl) {
      return undefined;
    }

    if (action === 'show-list-items') {
      if (!target.listName) {
        return undefined;
      }

      logService.info('llm', `TextChatService: contextual container override → show_list_items (${target.listName})`);
      return {
        toolName: 'show_list_items',
        toolArgs: {
          site_url: target.siteUrl,
          list_name: target.listName
        }
      };
    }

    logService.info('llm', `TextChatService: contextual container override → browse_document_library (${target.libraryName || target.siteName || target.siteUrl})`);
    return {
      toolName: 'browse_document_library',
      toolArgs: {
        site_url: target.siteUrl,
        ...(target.libraryName ? { library_name: target.libraryName } : {})
      }
    };
  }

  private getContextualVisibleItemOverride(text: string): IForcedToolOverride | undefined {
    const normalizedText = normalizeVisibleContextText(text).toLowerCase();
    if (!normalizedText || isHiePromptMessage(normalizedText)) {
      return undefined;
    }

    if (
      matchesAny(normalizedText, CONTEXTUAL_COMPOSE_ACTION_HINTS)
      && (matchesAny(normalizedText, CONTEXTUAL_EMAIL_HINTS) || matchesAny(normalizedText, CONTEXTUAL_TEAMS_HINTS))
    ) {
      return undefined;
    }

    const inferredAction = this.detectContextualVisibleAction(normalizedText);
    if (!inferredAction) {
      return undefined;
    }

    const match = this.resolveContextualVisibleItemMatch(normalizedText);
    if (!match) {
      return undefined;
    }

    if (inferredAction === 'summarize') {
      logService.info('llm', `TextChatService: contextual visible-item override → summarize #${match.index} (${match.title})`);
      return {
        toolName: 'read_file_content',
        toolArgs: {
          file_url: match.url,
          mode: 'summarize'
        }
      };
    }

    if (inferredAction === 'permissions') {
      logService.info('llm', `TextChatService: contextual visible-item override → permissions #${match.index} (${match.title})`);
      return {
        toolName: 'show_permissions',
        toolArgs: {
          target_name: match.title,
          target_url: match.url
        }
      };
    }

    logService.info('llm', `TextChatService: contextual visible-item override → preview #${match.index} (${match.title})`);
    return {
      toolName: 'show_file_details',
      toolArgs: {
        file_url: match.url,
        file_name: match.title
      }
    };
  }

  private detectContextualMutationAction(text: string): ContextualMutationAction | undefined {
    if (matchesAny(text, CONTEXTUAL_RENAME_HINTS)) {
      return 'rename';
    }

    const lastAssistantText = this.getLatestAssistantMessageText().toLowerCase();
    if (
      matchesAny(text, CONTEXTUAL_SELECTED_REFERENCE_HINTS)
      && matchesAny(lastAssistantText, CONTEXTUAL_RENAME_CLARIFICATION_HINTS)
    ) {
      return 'rename';
    }

    return undefined;
  }

  private resolveContextualSharePointColumnTarget(text: string): IMcpTargetContext | undefined {
    const state = useGrimoireStore.getState();
    const latestSiteBlock = [...state.blocks]
      .reverse()
      .find((block) => block.type === 'site-info');
    const latestSiteData = latestSiteBlock?.type === 'site-info' ? blockDataAsSiteInfo(latestSiteBlock) : undefined;
    const hieTarget = deriveMcpTargetContextFromHie(
      hybridInteractionEngine.captureCurrentSourceContext(),
      hybridInteractionEngine.getCurrentTaskContext(),
      hybridInteractionEngine.getCurrentArtifacts()
    );
    const siteBlockTarget = deriveMcpTargetContextFromUnknown(latestSiteData, 'hie-selection');
    const containerTarget = this.resolveContextualContainerTarget(text);

    return mergeMcpTargetContexts(
      hieTarget,
      siteBlockTarget,
      deriveMcpTargetContextFromUnknown({
        siteUrl: containerTarget.siteUrl,
        siteName: containerTarget.siteName,
        listName: containerTarget.listName
      }, 'hie-selection')
    );
  }

  private inferSharePointColumnDisplayName(text: string): string | undefined {
    const match = text.match(/\b(?:named|called|name)\s+["'“”]?([^"'“”]+?)["'“”]?(?=\s+(?:of\s+type|type|as|in|on|to|for|with|and|at)\b|[.?!,;:]?$|$)/i);
    if (!match?.[1]) {
      return undefined;
    }

    const candidate = match[1].trim().replace(/[.?!,;:]+$/, '').trim();
    return candidate || undefined;
  }

  private inferSharePointColumnTypeLabel(text: string): string | undefined {
    const patterns: Array<{ label: string; pattern: RegExp }> = [
      { label: 'Multiple lines of text', pattern: /\b(?:multiple lines of text|multi line text|multiline|notes?|paragraph|long text)\b/i },
      { label: 'Yes/No', pattern: /\b(?:yes\s*\/\s*no|yes no|checkbox|boolean|bool)\b/i },
      { label: 'Date and time', pattern: /\b(?:date and time|datetime|date time)\b/i },
      { label: 'Person or Group', pattern: /\b(?:person or group|person|people|user|group)\b/i },
      { label: 'Link or Picture', pattern: /\b(?:link or picture|hyperlink or picture|hyperlink|link|url|picture|image)\b/i },
      { label: 'Choice', pattern: /\b(?:choice|choices|dropdown|option|options)\b/i },
      { label: 'Number', pattern: /\b(?:number|numeric|integer|decimal|currency)\b/i },
      { label: 'Text', pattern: /\b(?:text|string|single line(?: of text)?)\b/i }
    ];

    const match = patterns.find((candidate) => candidate.pattern.test(text));
    return match?.label;
  }

  private buildContextualRenameOverride(match: IContextualVisibleItemMatch): IForcedToolOverride {
    const staticArgs: Record<string, unknown> = {
      oldFileOrFolderName: match.title,
      fileOrFolderUrl: match.targetContext?.fileOrFolderUrl || match.url,
      fileOrFolderName: match.targetContext?.fileOrFolderName || match.title
    };
    if (match.targetContext?.fileOrFolderId) {
      staticArgs.fileOrFolderId = match.targetContext.fileOrFolderId;
    }
    if (match.targetContext?.documentLibraryId) {
      staticArgs.documentLibraryId = match.targetContext.documentLibraryId;
    }
    if (match.targetContext?.documentLibraryUrl) {
      staticArgs.documentLibraryUrl = match.targetContext.documentLibraryUrl;
    }
    if (match.targetContext?.documentLibraryName) {
      staticArgs.documentLibraryName = match.targetContext.documentLibraryName;
    }
    if (match.targetContext?.siteUrl) {
      staticArgs.siteUrl = match.targetContext.siteUrl;
    }
    if (match.targetContext?.siteName) {
      staticArgs.siteName = match.targetContext.siteName;
    }

    logService.info('llm', `TextChatService: contextual mutation override → item-rename #${match.index} (${match.title})`);
    return {
      toolName: 'show_compose_form',
      toolArgs: {
        preset: 'item-rename',
        title: `Rename ${match.title}`,
        description: 'Review the selected item and enter the new name.',
        prefill_json: JSON.stringify({
          newFileOrFolderName: match.title
        }),
        static_args_json: JSON.stringify(staticArgs)
      }
    };
  }

  private buildNoProgressGuardrailMessage(latestUserText: string): string {
    const normalizedLatestUserText = normalizeVisibleContextText(latestUserText).toLowerCase();
    const priorUserText = this.getLatestPriorUserMessageText().toLowerCase();
    const contextualText = normalizedLatestUserText || priorUserText;
    const selectedMatch = this.resolveContextualVisibleItemMatch(contextualText)
      || (priorUserText ? this.resolveContextualVisibleItemMatch(priorUserText) : undefined);

    if (this.detectContextualMutationAction(contextualText) === 'rename' && selectedMatch) {
      return `I found "${selectedMatch.title}" selected. Tell me the new name you want, and I'll open the rename form.`;
    }

    if (selectedMatch) {
      return `I found "${selectedMatch.title}" selected, but I got stuck resolving the action. Please tell me exactly what you want to do with it.`;
    }

    return 'I got stuck resolving that request from the current context. Please specify the item number or filename and what you want to do.';
  }

  private detectContextualVisibleAction(text: string): ContextualVisibleAction | undefined {
    if (matchesAny(text, CONTEXTUAL_SUMMARIZE_HINTS)) {
      return 'summarize';
    }

    if (matchesAny(text, CONTEXTUAL_PERMISSION_HINTS)) {
      return 'permissions';
    }

    if (matchesAny(text, CONTEXTUAL_PREVIEW_HINTS)) {
      return 'preview';
    }

    const implicitIndex = this.parseReferencedVisibleItemIndex(text);
    if (implicitIndex === undefined) {
      return undefined;
    }

    const lastAssistantText = this.getLatestAssistantMessageText().toLowerCase();
    if (!matchesAny(lastAssistantText, WHICH_ONE_PROMPT_HINTS)) {
      return undefined;
    }

    const contextualUserText = this.getLatestUserMessageText().toLowerCase()
      || this.getLatestPriorUserMessageText().toLowerCase();
    if (matchesAny(contextualUserText, CONTEXTUAL_SUMMARIZE_HINTS)) {
      return 'summarize';
    }

    if (matchesAny(contextualUserText, CONTEXTUAL_PERMISSION_HINTS)) {
      return 'permissions';
    }

    if (matchesAny(contextualUserText, CONTEXTUAL_PREVIEW_HINTS)) {
      return 'preview';
    }

    return undefined;
  }

  private detectContextualContainerAction(
    text: string,
    target: IContextualContainerTarget
  ): ContextualContainerAction | undefined {
    const hasExplicitListDiscovery = matchesAny(text, CONTEXTUAL_LIST_DISCOVERY_HINTS);
    if (hasExplicitListDiscovery && (target.siteName || this.extractExplicitSiteReference(text))) {
      return 'list-lists';
    }

    if (target.listName && matchesAny(text, CONTEXTUAL_LIST_CONTENT_HINTS)) {
      return 'show-list-items';
    }

    if (target.libraryName && matchesAny(text, CONTEXTUAL_LIBRARY_HINTS)) {
      return 'browse-library';
    }

    return undefined;
  }

  private resolveContextualContainerTarget(text: string): IContextualContainerTarget {
    const state = useGrimoireStore.getState();
    const latestSiteBlock = [...state.blocks]
      .reverse()
      .find((block) => block.type === 'site-info');
    const latestSiteData = latestSiteBlock?.type === 'site-info' ? blockDataAsSiteInfo(latestSiteBlock) : undefined;

    const hieTarget = deriveMcpTargetContextFromHie(
      undefined,
      hybridInteractionEngine.getCurrentTaskContext(),
      hybridInteractionEngine.getCurrentArtifacts()
    );
    const siteBlockTarget = deriveMcpTargetContextFromUnknown(latestSiteData, 'hie-selection');
    const targetContext = mergeMcpTargetContexts(hieTarget, siteBlockTarget);
    const currentPageSiteUrl = state.userContext?.currentSiteUrl || state.userContext?.currentWebUrl;
    const currentPageSiteName = state.userContext?.currentSiteTitle || state.userContext?.currentWebTitle;
    const explicitSiteReference = this.extractExplicitSiteReference(text);

    const libraryName = this.resolveNamedContainerMatch(
      text,
      latestSiteData?.libraries || [],
      targetContext?.documentLibraryName
    );
    const matchedListName = this.resolveNamedContainerMatch(
      text,
      latestSiteData?.lists || [],
      targetContext?.listName
    );
    const explicitListName = this.extractExplicitListReference(text);
    const listName = matchedListName
      || explicitListName
      || ((targetContext?.listId || targetContext?.listUrl) ? targetContext?.listName : undefined);

    return {
      siteUrl: targetContext?.siteUrl || latestSiteData?.siteUrl || (!explicitSiteReference ? currentPageSiteUrl : undefined),
      siteName: targetContext?.siteName || latestSiteData?.siteName || explicitSiteReference || currentPageSiteName,
      libraryName,
      listName
    };
  }

  private resolveNamedContainerMatch(
    text: string,
    candidates: string[],
    fallbackName?: string
  ): string | undefined {
    const normalizedText = normalizeVisibleItemTitleText(text);
    const orderedCandidates = [
      ...(fallbackName ? [fallbackName] : []),
      ...candidates
    ].filter((value, index, array) => !!value && array.indexOf(value) === index);

    for (let i = 0; i < orderedCandidates.length; i++) {
      const candidate = orderedCandidates[i];
      const normalizedCandidate = normalizeVisibleItemTitleText(candidate);
      if (!normalizedCandidate) {
        continue;
      }
      if (normalizedText.includes(normalizedCandidate)) {
        return candidate;
      }
    }

    return undefined;
  }

  private extractExplicitSiteReference(text: string): string | undefined {
    const siteMatch = text.match(/\bsite\s+([a-z0-9][a-z0-9\- ]*[a-z0-9])\b/i);
    if (siteMatch?.[1]) {
      return siteMatch[1].trim();
    }

    const listsInMatch = text.match(/\blists?\s+in\s+([a-z0-9][a-z0-9\- ]*[a-z0-9])\b/i);
    if (listsInMatch?.[1]) {
      return listsInMatch[1].trim();
    }

    return undefined;
  }

  private extractExplicitListReference(text: string): string | undefined {
    const patterns = [
      /\b(?:to|in|on|for|from)\s+(?:the\s+)?list\s+["'“”]?([^"'“”]+?)["'“”]?(?=\s+(?:in|on|to|for|with|and|at)\b|[.?!,;:]?$|$)/i,
      /\blist\s+["'“”]?([^"'“”]+?)["'“”]?(?=\s+(?:in|on|to|for|with|and|at)\b|[.?!,;:]?$|$)/i
    ];

    for (let i = 0; i < patterns.length; i++) {
      const match = text.match(patterns[i]);
      if (!match?.[1]) {
        continue;
      }

      const candidate = match[1].trim().replace(/[.?!,;:]+$/, '').trim();
      if (candidate) {
        return candidate;
      }
    }

    return undefined;
  }

  private getCurrentHieSelectedIndex(): number | undefined {
    const selectedItems = hybridInteractionEngine.captureCurrentSourceContext()?.selectedItems
      || hybridInteractionEngine.getCurrentTaskContext()?.selectedItems;
    if (!selectedItems || selectedItems.length !== 1) {
      return undefined;
    }

    const index = selectedItems[0]?.index;
    return typeof index === 'number' && index > 0 ? index : undefined;
  }

  private getPreferredContextualBlockIds(): string[] {
    const sourceContext = hybridInteractionEngine.captureCurrentSourceContext();
    const taskContext = hybridInteractionEngine.getCurrentTaskContext();
    const candidates = [
      sourceContext?.sourceBlockId,
      taskContext?.sourceBlockId,
      taskContext?.derivedBlockId
    ];

    return candidates.filter((value, index, values): value is string => (
      typeof value === 'string' && value.trim().length > 0 && values.indexOf(value) === index
    ));
  }

  private resolveContextualVisibleItemMatch(text: string): IContextualVisibleItemMatch | undefined {
    const state = useGrimoireStore.getState();
    const explicitIndex = this.parseReferencedVisibleItemIndex(text);
    const blocks = this.getContextualCandidateBlocks();
    const selectedIndex = explicitIndex !== undefined
      ? explicitIndex
      : (this.getCurrentHieSelectedIndex()
        || (state.selectedActionIndices.length === 1 ? state.selectedActionIndices[0] : undefined));
    if (selectedIndex !== undefined) {
      for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        const match = this.getVisibleItemMatchFromBlock(block, selectedIndex);
        if (match) {
          return match;
        }
      }
    }

    return this.resolveVisibleItemMatchByTitle(text, blocks);
  }

  private getContextualCandidateBlocks(): IBlock[] {
    const state = useGrimoireStore.getState();
    const byId = new Map(state.blocks.map((block) => [block.id, block]));
    const candidates: IBlock[] = [];

    const addCandidate = (blockId?: string): void => {
      if (!blockId) return;
      const candidate = byId.get(blockId);
      if (!candidate || !this.isContextualVisibleItemBlock(candidate)) return;
      if (candidates.some((existing) => existing.id === candidate.id)) return;
      candidates.push(candidate);
    };

    this.getPreferredContextualBlockIds().forEach((blockId) => addCandidate(blockId));
    addCandidate(state.activeActionBlockId);

    for (let i = state.blocks.length - 1; i >= 0; i--) {
      const block = state.blocks[i];
      if (!this.isContextualVisibleItemBlock(block)) continue;
      if (candidates.some((candidate) => candidate.id === block.id)) continue;
      candidates.push(block);
    }

    return candidates;
  }

  private isContextualVisibleItemBlock(block: IBlock): boolean {
    return block.type === 'search-results' || block.type === 'document-library';
  }

  private getVisibleItemMatchFromBlock(block: IBlock, index: number): IContextualVisibleItemMatch | undefined {
    const zeroBased = index - 1;
    if (zeroBased < 0) {
      return undefined;
    }

    if (block.type === 'search-results') {
      const items = (block.data as ISearchResultsData).results as ISearchResult[];
      const item = items[zeroBased];
      if (!item?.url) return undefined;
      return {
        block,
        index,
        title: item.title,
        url: item.url,
        targetContext: deriveMcpTargetContextFromUnknown(
          {
            url: item.url,
            fileOrFolderUrl: item.url,
            fileOrFolderName: item.title
          },
          'hie-selection'
        )
      };
    }

    if (block.type === 'document-library') {
      const items = (block.data as IDocumentLibraryData).items as IDocumentItem[];
      const item = items[zeroBased];
      if (!item?.url || item.type !== 'file') return undefined;
      const documentLibraryData = block.data as IDocumentLibraryData;
      return {
        block,
        index,
        title: item.name,
        url: item.url,
        targetContext: deriveMcpTargetContextFromUnknown(
          {
            siteName: documentLibraryData.siteName,
            libraryName: documentLibraryData.libraryName,
            fileOrFolderUrl: item.url,
            fileOrFolderId: item.fileOrFolderId,
            fileOrFolderName: item.name,
            documentLibraryId: item.documentLibraryId,
            documentLibraryName: documentLibraryData.libraryName
          },
          'hie-selection'
        )
      };
    }

    return undefined;
  }

  private resolveVisibleItemMatchByTitle(
    text: string,
    blocks: IBlock[]
  ): IContextualVisibleItemMatch | undefined {
    const normalizedText = normalizeVisibleItemTitleText(text);
    if (!normalizedText) {
      return undefined;
    }

    const compactText = normalizedText.replace(/\s+/g, '');
    const matches: Array<IContextualVisibleItemMatch & { score: number }> = [];

    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const blockItems = this.getVisibleItemMatchesFromBlock(block);
      for (let j = 0; j < blockItems.length; j++) {
        const item = blockItems[j];
        const normalizedTitle = normalizeVisibleItemTitleText(item.title);
        if (normalizedTitle.length < 4) {
          continue;
        }

        const compactTitle = normalizedTitle.replace(/\s+/g, '');
        const exactPhraseMatch = normalizedText.includes(normalizedTitle);
        const compactMatch = compactTitle.length >= 4 && compactText.includes(compactTitle);
        if (!exactPhraseMatch && !compactMatch) {
          continue;
        }

        matches.push({
          ...item,
          score: Math.max(normalizedTitle.length, compactTitle.length)
        });
      }
    }

    if (matches.length === 0) {
      return undefined;
    }

    matches.sort((left, right) => right.score - left.score);
    const bestMatch = matches[0];
    const competingMatch = matches.find((candidate) => candidate.url !== bestMatch.url && candidate.score === bestMatch.score);
    if (competingMatch) {
      return undefined;
    }

    return {
      block: bestMatch.block,
      index: bestMatch.index,
      title: bestMatch.title,
      url: bestMatch.url
    };
  }

  private getVisibleItemMatchesFromBlock(block: IBlock): IContextualVisibleItemMatch[] {
    if (block.type === 'search-results') {
      const items = (block.data as ISearchResultsData).results as ISearchResult[];
      return items
        .map((item, index) => item?.url ? {
          block,
          index: index + 1,
          title: item.title,
          url: item.url
        } : undefined)
        .filter((item): item is IContextualVisibleItemMatch => !!item);
    }

    if (block.type === 'document-library') {
      const items = (block.data as IDocumentLibraryData).items as IDocumentItem[];
      return items
        .map((item, index) => item?.url && item.type === 'file' ? {
          block,
          index: index + 1,
          title: item.name,
          url: item.url
        } : undefined)
        .filter((item): item is IContextualVisibleItemMatch => !!item);
    }

    return [];
  }

  private buildContextualComposeOverrideDetails(
    scope: IResolvedContextualComposeScope | undefined,
    match: IContextualVisibleItemMatch | undefined,
    defaultDescription: string,
    itemDescription: string,
    multiItemDescription: string
  ): IContextualComposeOverrideDetails {
    if (scope) {
      return {
        description: scope.resolved && (scope.selectedIndices?.length || 0) > 1
          ? multiItemDescription
          : (scope.resolved ? itemDescription : defaultDescription),
        staticArgsJson: JSON.stringify(buildComposeStaticArgsFromScope(scope))
      };
    }

    if (!match) {
      return { description: defaultDescription };
    }

    return {
      description: itemDescription,
      staticArgsJson: JSON.stringify({
        attachmentUris: [match.url],
        shareSelectionIndices: [match.index],
        shareBlockId: match.block.id,
        shareItemTitle: match.title,
        fileOrFolderUrl: match.url,
        fileOrFolderName: match.title
      })
    };
  }

  private parseReferencedVisibleItemIndex(text: string): number | undefined {
    const normalizedText = normalizeVisibleContextText(text).toLowerCase();
    if (!normalizedText) {
      return undefined;
    }

    if (/\blast\b/i.test(normalizedText)) {
      const blocks = this.getContextualCandidateBlocks();
      if (blocks.length === 0) return undefined;
      const firstBlock = blocks[0];
      if (firstBlock.type === 'search-results') {
        return (firstBlock.data as ISearchResultsData).results.length || undefined;
      }
      if (firstBlock.type === 'document-library') {
        return (firstBlock.data as IDocumentLibraryData).items.length || undefined;
      }
    }

    const explicitMatch = normalizedText.match(/\b(?:item|result|document|doc|file|number|#)\s*(\d{1,2})(?:st|nd|rd|th)?\b/i)
      || normalizedText.match(/\b(\d{1,2})(?:st|nd|rd|th)?\b/);
    if (explicitMatch) {
      const parsed = parseInt(explicitMatch[1], 10);
      if (Number.isFinite(parsed) && parsed > 0) {
        return parsed;
      }
    }

    for (let i = 0; i < ORDINAL_INDEX_PATTERNS.length; i++) {
      if (ORDINAL_INDEX_PATTERNS[i].pattern.test(normalizedText)) {
        return ORDINAL_INDEX_PATTERNS[i].index;
      }
    }

    return undefined;
  }

  private shouldApplyFirstTurnRouting(): boolean {
    for (let i = 1; i < this.messages.length; i++) {
      const message = this.messages[i];
      if (message.role === 'user') {
        const content = typeof message.content === 'string' ? message.content.trim() : '';
        if (!content || isHiePromptMessage(content)) {
          continue;
        }
        logService.debug('llm', 'TextChatService: skipped first-turn routing (prior user turn present)');
        return false;
      }

      if (message.role === 'assistant' || message.role === 'tool') {
        logService.debug('llm', 'TextChatService: skipped first-turn routing (conversation already active)');
        return false;
      }
    }

    return true;
  }

  private trimHistory(): void {
    // Keep system prompt + last MAX_MESSAGES
    if (this.messages.length > MAX_MESSAGES + 1) {
      const systemMsg = this.messages[0];
      let sliced = this.messages.slice(-(MAX_MESSAGES));

      // Ensure we don't start with orphaned tool messages (no preceding assistant with tool_calls).
      // Trim leading tool messages — they'd cause a 400 from the API.
      while (sliced.length > 0 && sliced[0].role === 'tool') {
        sliced = sliced.slice(1);
      }
      // Also trim a leading assistant message with tool_calls if its tool responses were cut
      if (sliced.length > 0 && sliced[0].role === 'assistant' && sliced[0].tool_calls && sliced[0].tool_calls.length > 0) {
        const expectedIds = new Set<string>();
        sliced[0].tool_calls.forEach((tc) => expectedIds.add(tc.id));
        for (let j = 1; j < sliced.length; j++) {
          if (sliced[j].role === 'tool' && sliced[j].tool_call_id) {
            expectedIds.delete(sliced[j].tool_call_id!);
          }
          if (sliced[j].role === 'assistant' || sliced[j].role === 'user') break;
        }
        if (expectedIds.size > 0) {
          // Missing some tool responses — remove the assistant + its tool messages
          let endIdx = 1;
          while (endIdx < sliced.length && sliced[endIdx].role === 'tool') endIdx++;
          sliced = sliced.slice(endIdx);
        }
      }

      this.messages = [systemMsg, ...sliced];
    }

    // Summarize old tool results deep in history to save tokens.
    // Only compact messages beyond the recent window (last TOOL_RESULT_SUMMARY_THRESHOLD messages).
    const recentStart = this.messages.length - TOOL_RESULT_SUMMARY_THRESHOLD;
    for (let i = 1; i < recentStart; i++) {
      const msg = this.messages[i];
      if (msg.role === 'tool' && msg.content && msg.content.length > 200) {
        msg.content = summarizeToolResult(msg.content);
      }
    }
  }

  private getLatestUserMessageText(): string {
    for (let i = this.messages.length - 1; i >= 0; i--) {
      const msg = this.messages[i];
      if (msg.role !== 'user') continue;
      if (typeof msg.content !== 'string') continue;
      const text = msg.content;
      if (isHiePromptMessage(text)) {
        continue;
      }
      return text;
    }
    return '';
  }

  private getLatestPriorUserMessageText(): string {
    let seenLatest = false;
    for (let i = this.messages.length - 1; i >= 0; i--) {
      const msg = this.messages[i];
      if (msg.role !== 'user' || typeof msg.content !== 'string') continue;
      if (isHiePromptMessage(msg.content)) continue;
      if (!seenLatest) {
        seenLatest = true;
        continue;
      }
      return msg.content;
    }
    return '';
  }

  private getLatestAssistantMessageText(): string {
    for (let i = this.messages.length - 1; i >= 0; i--) {
      const msg = this.messages[i];
      if (msg.role !== 'assistant') continue;
      if (typeof msg.content !== 'string') continue;
      const text = msg.content.trim();
      if (!text) continue;
      return text;
    }
    return '';
  }

  private resolveToolExecution(
    toolName: string,
    args: Record<string, unknown>,
    latestUserText: string
  ): IResolvedToolExecution {
    if (toolName !== 'research_public_web') {
      return { toolName, args };
    }

    const query = typeof args.query === 'string' ? args.query.trim() : '';
    const targetUrl = typeof args.target_url === 'string' ? args.target_url.trim() : '';
    const normalizedLatestUserText = latestUserText.trim();

    if (targetUrl) {
      return { toolName, args };
    }

    const hasExternalSignal = hasExplicitExternalWebHint(normalizedLatestUserText)
      || (query ? hasExplicitExternalWebHint(query) : false);
    if (hasExternalSignal) {
      return { toolName, args };
    }

    const enterpriseSearchLike = hasActionableSearchIntent(normalizedLatestUserText)
      || hasInternalEnterpriseContentHint(normalizedLatestUserText)
      || (query ? hasInternalEnterpriseContentHint(query) : false);
    if (!enterpriseSearchLike) {
      return { toolName, args };
    }

    const nextArgs: Record<string, unknown> = {};
    const fallbackQuery = query || normalizedLatestUserText;
    if (fallbackQuery) {
      nextArgs.query = fallbackQuery;
    }

    logService.info('llm', 'TextChatService: rerouted research_public_web → search_sharepoint for enterprise search phrasing');
    return {
      toolName: 'search_sharepoint',
      args: nextArgs
    };
  }

  private logFirstTurnRoutingOutcome(
    outcome: 'tool_call' | 'clarification' | 'answer_only',
    observation: IFirstTurnRoutingObservation,
    actualToolName?: string
  ): void {
    const detail = buildFirstTurnRoutingLogDetail('text', outcome, observation, actualToolName);
    logService.info('llm', `First-turn routing: ${outcome}`, detail);

    if (outcome === 'clarification' && observation.isGenericEnterpriseSearch) {
      logService.warning('llm', 'Generic enterprise search fell back to clarification', detail);
    }
  }
}

// ─── Helpers (kept here — depend on TextChatRateLimitHeaderSource) ──

function resolveRateLimitDelay(
  headers: Pick<Headers, 'get'>,
  retryIndex: number,
  nowMs: number = Date.now()
): { delayMs: number; headerSource: TextChatRateLimitHeaderSource } {
  const retryAfterMs = parsePositiveNumber(headers.get('retry-after-ms') || undefined);
  if (retryAfterMs !== undefined) {
    return {
      delayMs: Math.ceil(retryAfterMs),
      headerSource: 'retry-after-ms'
    };
  }

  const compatibilityRetryAfterMs = parsePositiveNumber(headers.get('x-ms-retry-after-ms') || undefined);
  if (compatibilityRetryAfterMs !== undefined) {
    return {
      delayMs: Math.ceil(compatibilityRetryAfterMs),
      headerSource: 'x-ms-retry-after-ms'
    };
  }

  const retryAfter = headers.get('Retry-After') || undefined;
  const retryAfterSeconds = parsePositiveNumber(retryAfter);
  if (retryAfterSeconds !== undefined) {
    return {
      delayMs: Math.ceil(retryAfterSeconds * 1000),
      headerSource: 'retry-after-seconds'
    };
  }

  const retryAfterDateMs = retryAfter ? Date.parse(retryAfter) : NaN;
  if (Number.isFinite(retryAfterDateMs)) {
    const delayMs = retryAfterDateMs - nowMs;
    if (delayMs > 0) {
      return {
        delayMs,
        headerSource: 'retry-after-date'
      };
    }
  }

  return {
    delayMs: Math.min(DEFAULT_RETRY_DELAY_MS * Math.pow(2, retryIndex), MAX_RETRY_DELAY_MS),
    headerSource: 'fallback-exponential'
  };
}

function parsePositiveNumber(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return undefined;
  }
  return parsed;
}

// ─── Re-exports (preserve public API surface) ──────────────────

export { looksLikeInternalToolPayload, sanitizeAssistantText, shouldSuppressDisplayToolCall, truncateToolResult } from './TextChatHelpers';
