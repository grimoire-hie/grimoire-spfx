import { computeTimeDomainAmplitude, resolveSpeechGateThreshold } from './SpeechMouthAnalyzer';

describe('SpeechMouthAnalyzer helpers', () => {
  it('returns near-zero amplitude for a silent centered waveform', () => {
    const data = new Uint8Array([128, 128, 128, 128, 128, 128]);

    expect(computeTimeDomainAmplitude(data)).toBeCloseTo(0, 5);
  });

  it('returns a higher amplitude for a more energetic waveform', () => {
    const quiet = new Uint8Array([126, 128, 130, 128, 127, 129]);
    const loud = new Uint8Array([90, 166, 76, 180, 88, 168]);

    expect(computeTimeDomainAmplitude(loud)).toBeGreaterThan(computeTimeDomainAmplitude(quiet));
  });

  it('keeps the speech gate threshold bounded and above the silence floor', () => {
    const lowNoiseThreshold = resolveSpeechGateThreshold(0.004, 0.015);
    const highNoiseThreshold = resolveSpeechGateThreshold(0.03, 0.015);

    expect(lowNoiseThreshold).toBeGreaterThanOrEqual(0.015);
    expect(highNoiseThreshold).toBeGreaterThan(lowNoiseThreshold);
    expect(highNoiseThreshold).toBeLessThanOrEqual(0.04);
  });
});
