import { createCorrelationId } from './HAEContracts';
import { HybridInteractionEngine } from './HybridInteractionEngine';
import { createBlock } from '../../models/IBlock';

describe('HAEContracts', () => {
  it('creates correlation IDs with prefix', () => {
    const id = createCorrelationId('turn');
    expect(id.startsWith('turn-')).toBe(true);
  });

  it('builds visual snapshots with duplicate type detection', () => {
    const engine = new HybridInteractionEngine();
    engine.initialize(() => { /* no-op */ });

    const b1 = createBlock('search-results', 'Search A', {
      kind: 'search-results',
      query: 'alpha',
      results: [{ title: 'Doc A', summary: 'A', url: 'https://a' }],
      totalCount: 1,
      source: 'test'
    });
    const b2 = createBlock('search-results', 'Search B', {
      kind: 'search-results',
      query: 'beta',
      results: [{ title: 'Doc B', summary: 'B', url: 'https://b' }],
      totalCount: 1,
      source: 'test'
    });

    engine.onBlockCreated(b1);
    engine.onBlockCreated(b2);

    const snapshot = engine.getVisualStateSnapshot('turn-test');
    expect(snapshot.correlationId).toBe('turn-test');
    expect(snapshot.blocks.length).toBe(2);
    expect(snapshot.hasDuplicateTypes).toBe(true);
  });
});
