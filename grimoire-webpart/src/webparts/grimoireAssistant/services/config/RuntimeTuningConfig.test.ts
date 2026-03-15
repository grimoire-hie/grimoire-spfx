import { getRuntimeTuningConfig } from './RuntimeTuningConfig';

describe('RuntimeTuningConfig', () => {
  afterEach(() => {
    delete window.__GRIMOIRE_RUNTIME_TUNING__;
  });

  it('provides default Nano tuning for the compound workflow planner', () => {
    const tuning = getRuntimeTuningConfig().nano;

    expect(tuning.compoundWorkflowPlannerTimeoutMs).toBe(4500);
    expect(tuning.compoundWorkflowPlannerMaxTokens).toBe(120);
    expect(tuning.compoundWorkflowPlannerConfidenceThreshold).toBe(0.78);
    expect(tuning.blockRecapMaxTokens).toBe(640);
    expect(tuning.blockRecapRetryHeadroomTokens).toBe(160);
    expect(tuning.blockRecapRetryMinTokens).toBe(800);
  });

  it('applies window overrides for the compound workflow planner tuning', () => {
    window.__GRIMOIRE_RUNTIME_TUNING__ = {
      nano: {
        compoundWorkflowPlannerTimeoutMs: 5100,
        compoundWorkflowPlannerMaxTokens: 96,
        compoundWorkflowPlannerConfidenceThreshold: 0.9,
        blockRecapRetryHeadroomTokens: 120,
        blockRecapRetryMinTokens: 720
      }
    };

    const tuning = getRuntimeTuningConfig().nano;
    expect(tuning.compoundWorkflowPlannerTimeoutMs).toBe(5100);
    expect(tuning.compoundWorkflowPlannerMaxTokens).toBe(96);
    expect(tuning.compoundWorkflowPlannerConfidenceThreshold).toBe(0.9);
    expect(tuning.blockRecapRetryHeadroomTokens).toBe(120);
    expect(tuning.blockRecapRetryMinTokens).toBe(720);
  });
});
