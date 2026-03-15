import { getPreferencePersistenceDecision } from './PreferencePersistenceGate';

describe('PreferencePersistenceGate', () => {
  it('seeds the baseline without persisting on the first observed payload', () => {
    expect(getPreferencePersistenceDecision(undefined, 'payload-a')).toEqual({
      nextBaseline: 'payload-a',
      shouldPersist: false
    });
  });

  it('does not persist when the payload still matches the baseline', () => {
    expect(getPreferencePersistenceDecision('payload-a', 'payload-a')).toEqual({
      nextBaseline: 'payload-a',
      shouldPersist: false
    });
  });

  it('requests persistence when the payload changed after the baseline was established', () => {
    expect(getPreferencePersistenceDecision('payload-a', 'payload-b')).toEqual({
      nextBaseline: 'payload-a',
      shouldPersist: true
    });
  });
});
