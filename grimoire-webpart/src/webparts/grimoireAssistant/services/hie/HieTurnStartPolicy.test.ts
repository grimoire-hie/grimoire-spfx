import { resolveIngressTurnStartPolicy } from './HieTurnStartPolicy';

describe('HieTurnStartPolicy', () => {
  it('starts a new root when there is no active context', () => {
    expect(resolveIngressTurnStartPolicy('search for invoices', {
      hasTaskContext: false,
      hasVisibleBlocks: false
    })).toEqual({
      mode: 'new-root',
      reason: 'no-active-context'
    });
  });

  it('inherits the thread for contextual follow-up actions', () => {
    expect(resolveIngressTurnStartPolicy('summarize document 4', {
      hasTaskContext: true,
      hasVisibleBlocks: true
    })).toEqual({
      mode: 'inherit',
      reason: 'contextual-follow-up'
    });
  });

  it('inherits the thread for title-led follow-up actions against visible references', () => {
    expect(resolveIngressTurnStartPolicy('summarize teamsfx', {
      hasTaskContext: true,
      hasVisibleBlocks: true,
      visibleReferenceTitles: ['SPFx', 'TeamsFx']
    })).toEqual({
      mode: 'inherit',
      reason: 'visible-title-follow-up'
    });
  });

  it('starts a new root for obvious top-level requests even with active context', () => {
    expect(resolveIngressTurnStartPolicy('search for invoices', {
      hasTaskContext: true,
      hasVisibleBlocks: true
    })).toEqual({
      mode: 'new-root',
      reason: 'top-level-request'
    });
  });

  it('defers ambiguous follow-ups to HIE auto mode', () => {
    expect(resolveIngressTurnStartPolicy('can you help me with invoices?', {
      hasTaskContext: true,
      hasVisibleBlocks: true
    })).toEqual({
      mode: 'auto',
      reason: 'defer-to-hie-auto'
    });
  });

  it('still starts a new root for top-level search even if the query matches visible titles', () => {
    expect(resolveIngressTurnStartPolicy('search teamsfx', {
      hasTaskContext: true,
      hasVisibleBlocks: true,
      visibleReferenceTitles: ['TeamsFx']
    })).toEqual({
      mode: 'new-root',
      reason: 'top-level-request'
    });
  });
});
