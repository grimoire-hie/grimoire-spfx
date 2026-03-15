import { completeOutcome, errorOutcome } from './ToolRuntimeOutcomeHelpers';

describe('ToolRuntimeOutcomeHelpers', () => {
  it('builds complete outcome', () => {
    expect(completeOutcome('ok')).toEqual({ output: 'ok', phase: 'complete' });
  });

  it('builds error outcome', () => {
    expect(errorOutcome('bad')).toEqual({ output: 'bad', phase: 'error' });
  });
});
