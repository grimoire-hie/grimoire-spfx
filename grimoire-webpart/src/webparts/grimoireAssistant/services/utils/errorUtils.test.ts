import { normalizeError } from './errorUtils';

describe('normalizeError', () => {
  it('normalizes Error instances', () => {
    const normalized = normalizeError(new Error('boom'), 'fallback');
    expect(normalized).toEqual({ name: 'Error', message: 'boom' });
  });

  it('normalizes string throws', () => {
    const normalized = normalizeError('bad', 'fallback');
    expect(normalized).toEqual({ name: 'Error', message: 'bad' });
  });

  it('normalizes object throws with fallback message', () => {
    const normalized = normalizeError({ name: 'CustomFailure' }, 'fallback');
    expect(normalized).toEqual({ name: 'CustomFailure', message: 'fallback' });
  });
});
