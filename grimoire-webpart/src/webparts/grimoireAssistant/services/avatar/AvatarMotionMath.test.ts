import { composeEyeScale, computeMouthTransform } from './AvatarMotionMath';

describe('AvatarMotionMath', () => {
  it('keeps a higher eye-scale floor while speaking', () => {
    const speaking = composeEyeScale(1, 1, true);
    const silent = composeEyeScale(1, 1, false);

    expect(speaking).toBeGreaterThanOrEqual(0.14);
    expect(silent).toBeGreaterThanOrEqual(0.08);
    expect(speaking).toBeGreaterThan(silent);
  });

  it('computes bounded mouth transform for v2 articulation', () => {
    const transform = computeMouthTransform({
      openness: 1,
      width: 1,
      round: 1,
      mouthLift: -3,
      mouthWidthBoost: 0.2,
      mouthOpenBoost: 0.5,
      factor: 1
    }, true);

    expect(transform.scaleX).toBeGreaterThanOrEqual(0.52);
    expect(transform.scaleX).toBeLessThanOrEqual(1.75);
    expect(transform.scaleY).toBeGreaterThanOrEqual(0.48);
    expect(transform.scaleY).toBeLessThanOrEqual(2.05);
    expect(transform.jawOffsetY).toBeGreaterThan(0);
  });

  it('keeps legacy and v2 transforms intentionally different', () => {
    const legacy = computeMouthTransform({
      openness: 0.7,
      width: 0.5,
      round: 0.8,
      mouthLift: 0,
      mouthWidthBoost: 0,
      mouthOpenBoost: 0,
      factor: 1
    }, false);

    const modern = computeMouthTransform({
      openness: 0.7,
      width: 0.5,
      round: 0.8,
      mouthLift: 0,
      mouthWidthBoost: 0,
      mouthOpenBoost: 0,
      factor: 1
    }, true);

    expect(modern.scaleX).toBeLessThan(legacy.scaleX);
    expect(modern.jawOffsetY).toBeGreaterThan(0);
  });
});

