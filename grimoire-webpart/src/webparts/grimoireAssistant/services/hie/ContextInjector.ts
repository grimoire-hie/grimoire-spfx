/**
 * ContextInjector — Summarizes block data and injects context messages
 * into the LLM conversation via RealtimeAudioService.sendContextMessage().
 *
 * Two message types:
 * - [Visual context: ...] — silent injection when data arrives (no response triggered)
 * - [User interaction: ...] — injected when user clicks a block element (triggers response)
 */

import { BlockTracker } from './BlockTracker';
import { CONTEXT_COMPRESSION_SYSTEM_PROMPT } from '../../config/promptCatalog';
import {
  DEFAULT_HIE_EXPOSURE_POLICY,
  IBlockInteraction,
  IContextMessage,
  IHIEConfig
} from './HIETypes';
import type { NanoService } from '../nano/NanoService';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { logService } from '../logging/LogService';
import { createCorrelationId, IContextEnvelope } from './HAEContracts';
import { HieInteractionFormatter } from './HieInteractionFormatter';
import {
  formatHiePromptMessage,
  type IHiePromptMessage
} from './HiePromptProtocol';

export type SendContextFn = (text: string, triggerResponse: boolean, turnId?: string) => void;

const TOOL_SUGGESTIONS: Partial<Record<string, string>> = {
  'search-results': 'Suggested follow-ups: show_file_details (preview a result), browse_document_library (see its folder)',
  'document-library': 'Suggested follow-ups: show_file_details (preview a file), show_permissions (check access if the MCP server supports permission lookup)',
  'user-card': 'Suggested follow-ups: search_sharepoint (find their documents), show_activity_feed (recent activity)',
  'site-info': 'Suggested follow-ups: browse_document_library (browse libraries), show_list_items (view lists)',
  'list-items': 'Suggested follow-ups: show_file_details (details on linked items), show_chart (visualize data)',
  'permissions-view': 'Suggested follow-ups: use_m365_capability (modify sharing)',
  'file-preview': 'Suggested follow-ups: browse_document_library (see sibling files)',
  'activity-feed': 'Suggested follow-ups: show_file_details (details on mentioned files)',
  'form': 'User is filling a form. Do NOT ask for the same info via chat — the form handles it.',
  'markdown': 'Suggested follow-ups: save_note (save key info), read_email_content (read full email if this shows email results)'
};

const MAX_HISTORY_ENTRIES = 100;
const TOOL_COMPLETION_DEBOUNCE_MS = 300;

export interface IToolCompletionInfo {
  toolName: string;
  blockId: string;
  itemCount: number;
  /** Defaults to true. When false, completion context is injected silently. */
  triggerResponse?: boolean;
}

export class ContextInjector {
  private config: IHIEConfig;
  private tracker: BlockTracker;
  private nanoService: NanoService | undefined;
  private sendFn: SendContextFn;
  private debounceTimer: ReturnType<typeof setTimeout> | undefined;
  private pendingBlockIds: Set<string> = new Set();
  private history: IContextMessage[] = [];
  private completionDebounceTimer: ReturnType<typeof setTimeout> | undefined;
  private pendingCompletions: IToolCompletionInfo[] = [];
  private lastContextSignature: string = '';
  private lastContextAt: number = 0;
  private interactionFormatter: HieInteractionFormatter;

  constructor(config: IHIEConfig, tracker: BlockTracker, sendFn: SendContextFn, nanoService?: NanoService) {
    this.config = config;
    this.tracker = tracker;
    this.sendFn = sendFn;
    this.nanoService = nanoService;
    this.interactionFormatter = new HieInteractionFormatter(tracker);
  }

  public scheduleInjection(blockId: string): void {
    if (!this.config.contextInjectionEnabled) return;

    this.pendingBlockIds.add(blockId);

    if (this.debounceTimer !== undefined) {
      clearTimeout(this.debounceTimer);
    }

    this.debounceTimer = setTimeout(() => {
      this.flushPendingContext();
    }, this.config.contextDebounceMs);
  }

  public injectInteraction(interaction: IBlockInteraction): void {
    const exposurePolicy = interaction.exposurePolicy || DEFAULT_HIE_EXPOSURE_POLICY;
    if (exposurePolicy.mode === 'store-only') return;
    if (exposurePolicy.mode === 'response-triggering' && !this.config.interactionNotificationsEnabled) return;
    if (exposurePolicy.mode === 'silent-context' && !this.config.contextInjectionEnabled) return;

    const promptMessage = this.interactionFormatter.format(interaction);
    if (!promptMessage) {
      logService.warning('system', `HIE: No interaction format for action "${interaction.action}" — click dropped`);
      return;
    }
    const text = formatHiePromptMessage(promptMessage);

    logService.info('system', `HIE: Sending interaction to LLM: ${text.substring(0, 120)}...`);
    const triggerResponse = exposurePolicy.mode === 'response-triggering';
    this.emitContext('interaction', text, triggerResponse, [interaction.blockId], true, interaction.turnId);

    this.pushHistory({
      contextType: 'interaction',
      text,
      blockIds: [interaction.blockId],
      sentAt: Date.now()
    });
  }

  public injectFlowUpdate(text: string, blockIds: string[]): void {
    if (!this.config.flowOrchestrationEnabled) return;

    const message = formatHiePromptMessage({ kind: 'flow', body: text });
    this.emitContext('flow', message, false, blockIds, false);

    this.pushHistory({
      contextType: 'flow',
      text: message,
      blockIds,
      sentAt: Date.now()
    });
  }

  public getHistory(): ReadonlyArray<IContextMessage> {
    return this.history;
  }

  public injectVisualStateReminder(summary: string): void {
    if (!this.config.contextInjectionEnabled || !summary) return;

    const message = formatHiePromptMessage({ kind: 'visual-current-state', body: summary });
    this.emitContext('visual', message, false, [], false);

    this.pushHistory({
      contextType: 'visual',
      text: message,
      blockIds: [],
      sentAt: Date.now()
    });
  }

  public injectRawContext(text: string, triggerResponse: boolean): void {
    this.injectProjectedContext({ kind: 'visual', body: text }, triggerResponse, []);
  }

  public injectProjectedContext(
    message: IHiePromptMessage,
    triggerResponse: boolean,
    blockIds: string[],
    turnId?: string
  ): void {
    const text = formatHiePromptMessage(message);
    const contextType = message.kind === 'interaction'
      ? 'interaction'
      : (message.kind === 'flow' ? 'flow' : 'visual');
    this.emitContext(contextType, text, triggerResponse, blockIds, false, turnId);

    this.pushHistory({
      contextType,
      text,
      blockIds,
      sentAt: Date.now()
    });
  }

  public injectToolCompletion(info: IToolCompletionInfo): void {
    this.pendingCompletions.push(info);

    if (this.completionDebounceTimer !== undefined) {
      clearTimeout(this.completionDebounceTimer);
    }

    this.completionDebounceTimer = setTimeout(() => {
      this.flushPendingCompletions();
    }, TOOL_COMPLETION_DEBOUNCE_MS);
  }

  public reset(): void {
    if (this.debounceTimer !== undefined) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = undefined;
    }
    if (this.completionDebounceTimer !== undefined) {
      clearTimeout(this.completionDebounceTimer);
      this.completionDebounceTimer = undefined;
    }
    this.pendingBlockIds.clear();
    this.pendingCompletions = [];
    this.history = [];
    this.lastContextSignature = '';
    this.lastContextAt = 0;
  }

  private pushHistory(entry: IContextMessage): void {
    this.history.push(entry);
    if (this.history.length > MAX_HISTORY_ENTRIES) {
      this.history = this.history.slice(-MAX_HISTORY_ENTRIES);
    }
  }

  private flushPendingCompletions(): void {
    this.completionDebounceTimer = undefined;
    if (this.pendingCompletions.length === 0) return;

    const parts: string[] = [];
    const blockIds: string[] = [];
    let triggerResponse = false;
    const completionTurnIds = new Set<string>();

    this.pendingCompletions.forEach((completion) => {
      const tracked = this.tracker.get(completion.blockId);
      if (tracked && tracked.summary) {
        parts.push(`${completion.toolName}: ${tracked.summary}`);
      } else {
        parts.push(`${completion.toolName}: ${completion.itemCount} item${completion.itemCount !== 1 ? 's' : ''} returned`);
      }
      if (completion.blockId) blockIds.push(completion.blockId);
      if (tracked?.turnId) {
        completionTurnIds.add(tracked.turnId);
      }
      if (completion.triggerResponse !== false) {
        triggerResponse = true;
      }
    });

    this.pendingCompletions = [];

    const text = triggerResponse
      ? formatHiePromptMessage({
          kind: 'tool-completion',
          body: `${parts.join('; ')}. The results are now visible in the action panel. Briefly acknowledge what was found.`
        })
      : formatHiePromptMessage({
          kind: 'tool-completion',
          body: parts.join('; ')
        });
    this.emitContext(
      'visual',
      text,
      triggerResponse,
      blockIds,
      false,
      completionTurnIds.size === 1 ? Array.from(completionTurnIds)[0] : undefined
    );

    this.pushHistory({
      contextType: 'visual',
      text,
      blockIds,
      sentAt: Date.now()
    });
  }

  private flushPendingContext(): void {
    this.debounceTimer = undefined;
    if (this.pendingBlockIds.size === 0) return;

    const parts: string[] = [];
    const injectedIds: string[] = [];

    this.pendingBlockIds.forEach((blockId) => {
      const tracked = this.tracker.get(blockId);
      if (tracked && tracked.summary && !tracked.contextInjected) {
        if (tracked.references.length > 0) {
          const items = tracked.references.slice(0, 10).map((reference) => {
            let entry = `${reference.index}) ${reference.title}`;
            if (reference.detail) entry += ` (${reference.detail})`;
            return entry;
          });
          const suffix = tracked.references.length > 10
            ? ` ...and ${tracked.references.length - 10} more`
            : '';
          parts.push(`${tracked.summary}. Items: ${items.join(', ')}${suffix}`);
        } else {
          parts.push(tracked.summary);
        }
        injectedIds.push(blockId);
        this.tracker.markContextInjected(blockId);
      }
    });

    this.pendingBlockIds.clear();

    if (parts.length === 0) return;

    const seenSuggestions = new Set<string>();
    injectedIds.forEach((id) => {
      const tracked = this.tracker.get(id);
      if (tracked) {
        const suggestion = tracked.originTool === 'research_public_web'
          ? 'Suggested follow-ups: research_public_web (refine the public web search), save_note (save key info)'
          : TOOL_SUGGESTIONS[tracked.type];
        if (suggestion) seenSuggestions.add(suggestion);
      }
    });

    if (seenSuggestions.size > 0) {
      seenSuggestions.forEach((suggestion) => parts.push(suggestion));
    }

    const rawText = parts.join('. ');

    this.compressContext(rawText, this.config.maxContextLength).then((text) => {
      const message = formatHiePromptMessage({ kind: 'visual', body: text });
      this.emitContext('visual', message, false, injectedIds, false);

      this.pushHistory({
        contextType: 'visual',
        text: message,
        blockIds: injectedIds,
        sentAt: Date.now()
      });
    }).catch(() => { /* Nano compression is best-effort */ });
  }

  private async compressContext(rawContext: string, maxChars: number): Promise<string> {
    if (rawContext.length <= maxChars) return rawContext;

    if (!this.nanoService) return rawContext.slice(0, maxChars - 3) + '...';

    const tuning = getRuntimeTuningConfig().nano;
    const response = await this.nanoService.classify(
      CONTEXT_COMPRESSION_SYSTEM_PROMPT,
      `Max ${maxChars} chars. Compress:\n${rawContext}`,
      tuning.contextCompressionTimeoutMs
    );

    return response || rawContext.slice(0, maxChars - 3) + '...';
  }

  private emitContext(
    contextType: 'visual' | 'interaction' | 'flow',
    text: string,
    triggerResponse: boolean,
    blockIds: string[],
    allowDuplicate: boolean,
    turnId?: string
  ): void {
    const now = Date.now();
    const signature = `${contextType}|${triggerResponse}|${text}`;
    const duplicateWindowMs = 1200;

    if (!allowDuplicate && this.lastContextSignature === signature && (now - this.lastContextAt) < duplicateWindowMs) {
      logService.debug('system', `HIE: duplicate context skipped (${contextType})`);
      return;
    }

    this.lastContextSignature = signature;
    this.lastContextAt = now;

    this.sendFn(text, triggerResponse, turnId);

    const envelope: IContextEnvelope = {
      envelopeId: createCorrelationId('ctxenv'),
      correlationId: createCorrelationId('ctx'),
      createdAt: now,
      contextType,
      triggerResponse,
      blockIds,
      text,
      source: 'hie',
      turnId
    };
    logService.debug('system', 'HIE context envelope', JSON.stringify(envelope));
  }
}
