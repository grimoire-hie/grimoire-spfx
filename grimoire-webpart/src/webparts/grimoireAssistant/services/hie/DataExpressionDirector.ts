/**
 * DataExpressionDirector — Centralized data-aware expression triggers.
 *
 * Replaces scattered hardcoded setExpression() calls in tool handlers
 * with priority-based rules. Coordinates with SentimentAnalyzer:
 *   - Data expressions: priority 8–15
 *   - Sentiment expressions: priority 5–10
 *   - LLM's set_expression tool: priority 20 (always wins)
 */

import { Expression } from '../avatar/ExpressionEngine';
import { IHIEConfig, IExpressionRule } from './HIETypes';

export interface IExpressionTriggerSnapshot {
  triggerId: string;
  expression: Expression;
  firedAt: number;
}

// ─── Expression Rules ───────────────────────────────────────────

const DATA_EXPRESSION_RULES: IExpressionRule[] = [
  { id: 'search-many', trigger: 'search-many', expression: 'happy', revertMs: 2000, priority: 10 },
  { id: 'search-zero', trigger: 'search-zero', expression: 'confused', revertMs: 3000, priority: 12 },
  { id: 'error', trigger: 'error', expression: 'confused', revertMs: 3000, priority: 15 },
  { id: 'user-click', trigger: 'user-click', expression: 'thinking', revertMs: 2000, priority: 8 },
  { id: 'user-confirm', trigger: 'user-confirm', expression: 'happy', revertMs: 1500, priority: 10 },
  { id: 'user-cancel', trigger: 'user-cancel', expression: 'idle', revertMs: 0, priority: 5 },
  { id: 'flow-awaiting', trigger: 'flow-awaiting', expression: 'listening', revertMs: 0, priority: 6 },
  { id: 'tool-success', trigger: 'tool-success', expression: 'happy', revertMs: 2000, priority: 9 },
  { id: 'browse-loaded', trigger: 'browse-loaded', expression: 'happy', revertMs: 1500, priority: 8 },
  { id: 'form-displayed', trigger: 'form-displayed', expression: 'listening', revertMs: 0, priority: 8 },
  { id: 'form-submitted', trigger: 'form-submitted', expression: 'happy', revertMs: 2000, priority: 10 },
  { id: 'form-error', trigger: 'form-error', expression: 'confused', revertMs: 3000, priority: 12 },
  { id: 'profile-loaded', trigger: 'profile-loaded', expression: 'happy', revertMs: 1500, priority: 8 },
  { id: 'insights-empty', trigger: 'insights-empty', expression: 'confused', revertMs: 3000, priority: 10 },
  { id: 'note-saved', trigger: 'note-saved', expression: 'happy', revertMs: 1500, priority: 9 },
  { id: 'user-look', trigger: 'user-look', expression: 'thinking', revertMs: 2000, priority: 8 },
  { id: 'user-summarize', trigger: 'user-summarize', expression: 'thinking', revertMs: 3000, priority: 9 },
  { id: 'user-chat-about', trigger: 'user-chat-about', expression: 'happy', revertMs: 1500, priority: 8 }
];

export type SetExpressionFn = (expression: Expression) => void;

export class DataExpressionDirector {
  private config: IHIEConfig;
  private setExpression: SetExpressionFn;
  private currentPriority: number = 0;
  private revertTimer: ReturnType<typeof setTimeout> | undefined;
  /** When true, expression changes still apply but revert timers are deferred */
  private revertsSuppressed: boolean = false;
  /** Pending revert to apply when suppression ends */
  private pendingRevertMs: number = 0;
  /** Last expression trigger for inspector surfacing */
  private lastTrigger: IExpressionTriggerSnapshot | undefined;

  constructor(config: IHIEConfig, setExpression: SetExpressionFn) {
    this.config = config;
    this.setExpression = setExpression;
  }

  /**
   * Trigger expression based on a data event.
   * Only applies if priority >= current active priority.
   */
  public trigger(triggerId: string): void {
    if (!this.config.dataExpressionsEnabled) return;

    const rule = this.findRule(triggerId);
    if (!rule) return;

    if (rule.priority < this.currentPriority) return;

    this.applyRule(rule);
  }

  /**
   * Trigger expression for tool completion based on result data.
   */
  public onToolComplete(toolName: string, success: boolean, itemCount: number): void {
    if (!this.config.dataExpressionsEnabled) return;

    if (!success) {
      this.trigger('error');
      return;
    }

    // Search-specific rules
    if (toolName.startsWith('search_')) {
      if (itemCount > 3) {
        this.trigger('search-many');
      } else if (itemCount === 0) {
        this.trigger('search-zero');
      } else {
        this.trigger('tool-success');
      }
      return;
    }

    // Browse-specific
    if (toolName === 'browse_document_library' || toolName === 'show_list_items') {
      this.trigger('browse-loaded');
      return;
    }

    // Error tool
    if (toolName === 'show_error') {
      this.trigger('error');
      return;
    }

    // Form tool
    if (toolName === 'show_compose_form') {
      this.trigger('form-displayed');
      return;
    }

    // Generic success for other tools that produce results
    if (itemCount > 0) {
      this.trigger('tool-success');
    }
  }

  /**
   * Trigger expression for user interaction events.
   */
  public onInteraction(action: string): void {
    if (!this.config.dataExpressionsEnabled) return;

    if (action === 'look') {
      this.trigger('user-look');
    } else if (action === 'summarize') {
      this.trigger('user-summarize');
    } else if (action === 'chat-about') {
      this.trigger('user-chat-about');
    } else if (action === 'confirm' || action === 'submit-form') {
      this.trigger(action === 'submit-form' ? 'form-submitted' : 'user-confirm');
    } else if (action === 'cancel' || action === 'cancel-form') {
      this.trigger('user-cancel');
    } else if (action.startsWith('click-') || action === 'navigate') {
      this.trigger('user-click');
    }
  }

  /**
   * Notify that the LLM explicitly set an expression via tool call.
   * This always wins (priority 20) and clears any active data expression.
   */
  public onLlmExpression(): void {
    this.clearRevert();
    this.currentPriority = 20;
  }

  /**
   * Suppress revert timers during multi-step tool loops to prevent
   * expression flashing (e.g. thinking→happy→idle→thinking oscillation).
   * Expressions still apply immediately, but reverts are deferred until release.
   */
  public suppressReverts(): void {
    this.revertsSuppressed = true;
    this.clearRevert();
    this.pendingRevertMs = 0;
  }

  /**
   * Release revert suppression. If the last expression had a revert,
   * start the revert timer now.
   */
  public releaseReverts(): void {
    this.revertsSuppressed = false;
    if (this.pendingRevertMs > 0) {
      this.revertTimer = setTimeout(() => {
        this.revertTimer = undefined;
        this.currentPriority = 0;
        this.setExpression('idle');
      }, this.pendingRevertMs);
      this.pendingRevertMs = 0;
    }
  }

  public getLastTrigger(): Readonly<IExpressionTriggerSnapshot> | undefined {
    return this.lastTrigger;
  }

  public reset(): void {
    this.clearRevert();
    this.currentPriority = 0;
    this.revertsSuppressed = false;
    this.pendingRevertMs = 0;
    this.lastTrigger = undefined;
  }

  // ─── Private ──────────────────────────────────────────────────

  private findRule(triggerId: string): IExpressionRule | undefined {
    for (let i = 0; i < DATA_EXPRESSION_RULES.length; i++) {
      if (DATA_EXPRESSION_RULES[i].trigger === triggerId) {
        return DATA_EXPRESSION_RULES[i];
      }
    }
    return undefined;
  }

  private applyRule(rule: IExpressionRule): void {
    this.clearRevert();
    this.currentPriority = rule.priority;
    this.setExpression(rule.expression);
    this.lastTrigger = { triggerId: rule.id, expression: rule.expression, firedAt: Date.now() };

    if (rule.revertMs > 0) {
      if (this.revertsSuppressed) {
        // Defer revert until suppression is released
        this.pendingRevertMs = rule.revertMs;
      } else {
        this.revertTimer = setTimeout(() => {
          this.revertTimer = undefined;
          this.currentPriority = 0;
          this.setExpression('idle');
        }, rule.revertMs);
      }
    }
  }

  private clearRevert(): void {
    if (this.revertTimer !== undefined) {
      clearTimeout(this.revertTimer);
      this.revertTimer = undefined;
    }
  }
}
