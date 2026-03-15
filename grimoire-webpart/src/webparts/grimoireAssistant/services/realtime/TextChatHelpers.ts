/**
 * TextChatHelpers
 * Utility functions extracted from TextChatService to keep the main service file focused
 * on orchestration and streaming logic.
 */

import { getObservedFirstToolName } from './IntentRoutingPolicy';
import { shouldUseImmediateCompletionAck } from '../tools/ToolCompletionAcknowledgment';
import { isExplicitSelectionListRequest } from '../tools/ToolAcknowledgment';
import { wrapToolResult, unwrapToolResult } from './PromptSafety';
import { TOOL_CONTEXT_LENGTH_LIMITS } from '../../config/assistantLengthLimits';

// ─── Constants ──────────────────────────────────────────────────

/** Max characters kept per tool result message — longer results are truncated */
const MAX_TOOL_RESULT_CHARS = TOOL_CONTEXT_LENGTH_LIMITS.toolResultMaxChars;
/** Higher content cap for read_* tools before passing context back to the model */
const MAX_READ_TOOL_CONTENT_CHARS = TOOL_CONTEXT_LENGTH_LIMITS.readToolContentMaxChars;

/** Tools that already render visual data blocks — used to suppress duplicate display calls */
const DISPLAY_TOOLS: ReadonlySet<string> = new Set(['show_markdown', 'show_info_card', 'show_selection_list']);
const READ_CONTENT_TOOLS: ReadonlySet<string> = new Set([
  'read_file_content',
  'read_email_content',
  'read_teams_messages'
]);

// ─── Functions ──────────────────────────────────────────────────

export function looksLikeInternalToolPayload(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;

  const hasRoute = /"to"\s*:\s*"functions\.[^"]+"/i.test(trimmed);
  const hasArgs = /"args"\s*:\s*\{[\s\S]*\}/i.test(trimmed);
  const expressionOnly = /^\s*\{\s*"expression"\s*:\s*"[^"]+"\s*\}\s*$/i.test(trimmed);
  const looksJsonLike = trimmed.startsWith('{') || trimmed.startsWith('[');

  return expressionOnly || (looksJsonLike && hasRoute && hasArgs);
}

export function sanitizeAssistantText(text: string): string {
  const trimmed = text.trim();
  if (!trimmed) return '';
  if (looksLikeInternalToolPayload(trimmed)) return '';
  return text;
}

export function resolveImmediateCompletionAck(
  toolNames: ReadonlyArray<string>,
  completionAcks: ReadonlyArray<{ toolName: string; text: string }>,
  latestUserText: string
): string | undefined {
  if (completionAcks.length !== 1) {
    return undefined;
  }

  const meaningfulToolNames = toolNames.filter((toolName) => getObservedFirstToolName([toolName]) !== undefined);
  if (meaningfulToolNames.length !== 1) {
    return undefined;
  }

  const toolName = meaningfulToolNames[0];
  if (completionAcks[0].toolName !== toolName) {
    return undefined;
  }

  if (!shouldUseImmediateCompletionAck(toolName, latestUserText)) {
    return undefined;
  }

  return completionAcks[0].text;
}

export function createAbortError(): Error {
  if (typeof DOMException !== 'undefined') {
    return new DOMException('The operation was aborted.', 'AbortError');
  }
  const error = new Error('The operation was aborted.');
  error.name = 'AbortError';
  return error;
}

export function sleepWithSignal(signal: AbortSignal, ms: number): Promise<void> {
  if (signal.aborted) {
    return Promise.reject(createAbortError());
  }

  return new Promise((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | undefined = undefined;
    function onAbort(): void {
      if (timer !== undefined) {
        clearTimeout(timer);
      }
      signal.removeEventListener('abort', onAbort);
      reject(createAbortError());
    }

    timer = setTimeout(() => {
      signal.removeEventListener('abort', onAbort);
      resolve();
    }, ms);

    signal.addEventListener('abort', onAbort, { once: true });
  });
}

export function shouldSuppressDisplayToolCall(
  toolName: string,
  visualDataToolSucceeded: boolean,
  latestUserText: string
): boolean {
  if (!visualDataToolSucceeded || !DISPLAY_TOOLS.has(toolName)) return false;
  if (toolName === 'show_selection_list' && isExplicitSelectionListRequest(latestUserText)) {
    return false;
  }
  return true;
}

/** Cap tool result size before adding to history */
export function truncateToolResult(output: string, toolName?: string): string {
  if (toolName && READ_CONTENT_TOOLS.has(toolName)) {
    return truncateReadContentToolResult(output);
  }
  if (output.length <= MAX_TOOL_RESULT_CHARS) return output;
  return output.slice(0, MAX_TOOL_RESULT_CHARS) + '... [truncated]';
}

/**
 * For read_* tools, preserve valid JSON and truncate only the `content` field.
 * This avoids cutting JSON mid-string and keeps richer context for summarization.
 */
export function truncateReadContentToolResult(output: string): string {
  let parsed: unknown;
  try {
    parsed = JSON.parse(output);
  } catch {
    // Preserve raw output if not parseable; avoid creating malformed JSON fragments.
    return output;
  }

  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return output;
  }

  const payload = parsed as Record<string, unknown>;
  const content = payload.content;
  if (typeof content !== 'string') {
    return output;
  }

  if (content.length <= MAX_READ_TOOL_CONTENT_CHARS) {
    return output;
  }

  payload.content = content.slice(0, MAX_READ_TOOL_CONTENT_CHARS);
  payload.truncated = true;
  payload.content_truncated_for_context = true;
  return JSON.stringify(payload);
}

/** Summarize an old tool result to just success/error + key identifiers */
export function summarizeToolResult(content: string): string {
  const wrapped = unwrapToolResult(content);
  if (wrapped) {
    return wrapToolResult(wrapped.tool, summarizeToolResult(wrapped.content));
  }

  try {
    const parsed = JSON.parse(content);
    if (parsed && typeof parsed === 'object') {
      const success = parsed.success !== undefined ? parsed.success : true;
      const error = parsed.error ? `: ${String(parsed.error).slice(0, 80)}` : '';
      // Preserve a short fingerprint so the model can still reference it
      const keys = Object.keys(parsed).slice(0, 3).join(', ');
      return JSON.stringify({ _summary: true, success, keys, error: error || undefined });
    }
  } catch { /* not JSON */ }
  // Non-JSON: keep first 150 chars
  return content.slice(0, 150) + (content.length > 150 ? '...' : '');
}

export function withTimeout<T>(promise: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([
    promise.then((v) => { clearTimeout(timer); return v; }),
    new Promise<T>((resolve) => { timer = setTimeout(() => resolve(fallback), ms); })
  ]);
}
