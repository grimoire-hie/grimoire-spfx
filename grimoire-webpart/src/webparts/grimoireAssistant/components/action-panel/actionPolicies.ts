/**
 * actionPolicies
 * Central guardrails for header actions (Focus, Summarize, Chat).
 */

import type { IBlock } from '../../models/IBlock';
import type { ISelectionCandidate } from './selectionHelpers';

export type HeaderAction = 'focus' | 'summarize' | 'chat';

export interface IActionPolicyResult {
  enabled: boolean;
  reason?: string;
  eligibleCount: number;
  selectionCount: number;
  maxItems?: number;
}

export interface IFocusGuardrailSummary {
  singleSelectionRequiredFor: string[];
  notes: string[];
  emailSelectionCount: number;
  eventSelectionCount: number;
}

const SUMMARIZABLE_KINDS: ReadonlySet<string> = new Set([
  'document',
  'file',
  'email',
  'calendar-event',
  'message',
  'info'
]);

const SUMMARIZE_MAX_ITEMS = 5;
const CHAT_MAX_ITEMS = 5;

function getEligibleForAction(action: HeaderAction, candidates: ISelectionCandidate[]): ISelectionCandidate[] {
  if (action === 'summarize') {
    return candidates.filter((c) => SUMMARIZABLE_KINDS.has(c.kind));
  }
  return candidates;
}

export function evaluateHeaderAction(
  action: HeaderAction,
  activeBlock: IBlock | undefined,
  selectedCandidates: ISelectionCandidate[]
): IActionPolicyResult {
  if (!activeBlock) {
    return {
      enabled: false,
      reason: 'No current result block.',
      eligibleCount: 0,
      selectionCount: 0
    };
  }

  if (selectedCandidates.length === 0) {
    return {
      enabled: false,
      reason: 'Select one or more items first.',
      eligibleCount: 0,
      selectionCount: 0
    };
  }

  const eligible = getEligibleForAction(action, selectedCandidates);
  if (eligible.length === 0) {
    return {
      enabled: false,
      reason: action === 'summarize'
        ? 'Summarize supports document, email, calendar/message, and info items.'
        : 'No eligible selection.',
      eligibleCount: 0,
      selectionCount: selectedCandidates.length
    };
  }

  if (action === 'summarize') {
    return {
      enabled: true,
      eligibleCount: eligible.length,
      selectionCount: selectedCandidates.length,
      maxItems: SUMMARIZE_MAX_ITEMS
    };
  }

  if (action === 'chat') {
    return {
      enabled: true,
      eligibleCount: eligible.length,
      selectionCount: selectedCandidates.length,
      maxItems: CHAT_MAX_ITEMS
    };
  }

  return {
    enabled: true,
    eligibleCount: eligible.length,
    selectionCount: selectedCandidates.length
  };
}

export function getEligibleCandidatesForAction(
  action: HeaderAction,
  selectedCandidates: ISelectionCandidate[]
): ISelectionCandidate[] {
  if (action === 'summarize') {
    return selectedCandidates.filter((c) => SUMMARIZABLE_KINDS.has(c.kind));
  }
  return selectedCandidates.slice();
}

export function buildFocusGuardrailSummary(selectedCandidates: ISelectionCandidate[]): IFocusGuardrailSummary {
  const emailSelectionCount = selectedCandidates.filter((c) => c.kind === 'email').length;
  const eventSelectionCount = selectedCandidates.filter((c) => c.kind === 'calendar-event').length;
  const notes: string[] = [];

  if (emailSelectionCount > 1) {
    notes.push('Email reply/forward/delete must target exactly one focused email.');
  }
  if (eventSelectionCount > 1) {
    notes.push('Calendar update/cancel must target exactly one focused event.');
  }
  if (selectedCandidates.length > 1) {
    notes.push('Batch focus supports broad tasks like summarize/compare, not single-target mutations.');
  }

  return {
    singleSelectionRequiredFor: [
      'email-reply',
      'email-reply-all-thread',
      'email-forward',
      'email-delete',
      'event-update',
      'event-cancel'
    ],
    notes,
    emailSelectionCount,
    eventSelectionCount
  };
}
