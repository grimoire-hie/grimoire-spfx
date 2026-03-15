import {
  AvatarMotionProfile,
  BLINK_INTERVAL_MAX_MS,
  BLINK_INTERVAL_MIN_MS
} from './AvatarMotionProfile';

describe('AvatarMotionProfile', () => {
  it('produces deterministic ambient and scalar samples for the same seed', () => {
    const a = new AvatarMotionProfile('seed:user@example.com|classic');
    const b = new AvatarMotionProfile('seed:user@example.com|classic');

    const sampleA = a.sampleAmbient(4321);
    const sampleB = b.sampleAmbient(4321);

    expect(sampleA).toEqual(sampleB);
    expect(a.getRootFloatAmplitude()).toBeCloseTo(b.getRootFloatAmplitude(), 8);
    expect(a.getRootPulseAmplitude()).toBeCloseTo(b.getRootPulseAmplitude(), 8);
    expect(a.getEyeJitterAmplitude()).toBeCloseTo(b.getEyeJitterAmplitude(), 8);
  });

  it('starts first blink within configured interval bounds', () => {
    const profile = new AvatarMotionProfile('blink-bounds');
    let blinkStart = -1;

    for (let t = 0; t <= 9000; t += 16) {
      const blink = profile.sampleBlink(t);
      if (blink > 0) {
        blinkStart = t;
        break;
      }
    }

    expect(blinkStart).toBeGreaterThanOrEqual(BLINK_INTERVAL_MIN_MS);
    expect(blinkStart).toBeLessThanOrEqual(BLINK_INTERVAL_MAX_MS + 32);
  });

  it('generates a blink waveform that closes and reopens', () => {
    const profile = new AvatarMotionProfile('blink-waveform');
    let blinkStart = -1;

    for (let t = 0; t <= 9000; t += 16) {
      if (profile.sampleBlink(t) > 0) {
        blinkStart = t;
        break;
      }
    }

    expect(blinkStart).toBeGreaterThan(0);

    let peak = 0;
    let reopenedAt = -1;
    for (let t = blinkStart; t <= blinkStart + 1000; t += 16) {
      const value = profile.sampleBlink(t);
      peak = Math.max(peak, value);
      if (t > blinkStart + 150 && value === 0) {
        reopenedAt = t;
        break;
      }
    }

    expect(peak).toBeGreaterThan(0.95);
    expect(reopenedAt).toBeGreaterThan(blinkStart);
  });
});

