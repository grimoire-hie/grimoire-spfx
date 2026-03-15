import { evaluateHeaderAction, getEligibleCandidatesForAction, buildFocusGuardrailSummary } from './actionPolicies';
import type { IBlock } from '../../models/IBlock';
import type { ISelectionCandidate } from './selectionHelpers';

function makeBlock(type: IBlock['type']): IBlock {
  return {
    id: 'b1',
    type,
    title: 'Block',
    timestamp: new Date(),
    dismissible: true,
    data: { kind: type } as never
  };
}

function makeCandidate(overrides?: Partial<ISelectionCandidate>): ISelectionCandidate {
  return {
    index: 1,
    title: 'Item 1',
    kind: 'document',
    payload: {},
    ...overrides
  };
}

describe('actionPolicies', () => {
  it('disables actions when no active block is present', () => {
    const policy = evaluateHeaderAction('focus', undefined, [makeCandidate()]);
    expect(policy.enabled).toBe(false);
    expect(policy.reason).toContain('No current result block');
  });

  it('allows multi-selection chat and exposes max cap', () => {
    const policy = evaluateHeaderAction('chat', makeBlock('search-results'), [
      makeCandidate({ index: 1 }),
      makeCandidate({ index: 2, title: 'Item 2' })
    ]);
    expect(policy.enabled).toBe(true);
    expect(policy.maxItems).toBe(5);
  });

  it('filters summarize candidates by kind', () => {
    const selected = [
      makeCandidate({ kind: 'document', index: 1 }),
      makeCandidate({ kind: 'list-item', index: 2 })
    ];
    const policy = evaluateHeaderAction('summarize', makeBlock('list-items'), selected);
    expect(policy.enabled).toBe(true);
    const eligible = getEligibleCandidatesForAction('summarize', selected);
    expect(eligible).toHaveLength(1);
    expect(eligible[0].kind).toBe('document');
  });

  it('adds guardrail notes for multi-email focus', () => {
    const summary = buildFocusGuardrailSummary([
      makeCandidate({ kind: 'email', index: 1 }),
      makeCandidate({ kind: 'email', index: 2, title: 'Item 2' })
    ]);
    expect(summary.emailSelectionCount).toBe(2);
    expect(summary.notes.join(' ')).toContain('exactly one focused email');
  });
});
