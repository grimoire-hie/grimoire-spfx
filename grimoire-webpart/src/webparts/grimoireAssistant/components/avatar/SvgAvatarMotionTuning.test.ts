import { refineBrowMotionForVisage, resolveEyeMotionForVisage } from './SvgAvatarMotionTuning';

function magnitude(x: number, y: number): number {
  return Math.hypot(x, y);
}

describe('SvgAvatarMotionTuning', () => {
  it('makes action-panel gaze noticeably stronger than idle drift', () => {
    const idle = resolveEyeMotionForVisage('classic', 1.4, 0.8, 1, 1, 0.6);
    const focused = resolveEyeMotionForVisage('classic', 8.2, 1.2, 1, 1, 0.6);
    const idleAverage = (
      magnitude(idle.leftEyeOffsetX, idle.leftEyeOffsetY)
      + magnitude(idle.rightEyeOffsetX, idle.rightEyeOffsetY)
    ) / 2;
    const focusedAverage = (
      magnitude(focused.leftEyeOffsetX, focused.leftEyeOffsetY)
      + magnitude(focused.rightEyeOffsetX, focused.rightEyeOffsetY)
    ) / 2;

    expect(focusedAverage).toBeGreaterThan(idleAverage * 2.5);
  });

  it('keeps blink closure visibly stronger while staying bounded', () => {
    const blink = resolveEyeMotionForVisage('anonyMousse', 0, 0, 0.08, 1, 0);
    const open = resolveEyeMotionForVisage('anonyMousse', 0, 0, 1.2, 1, 0);

    expect(blink.eyeScale).toBeGreaterThanOrEqual(0.1);
    expect(blink.eyeScale).toBeLessThanOrEqual(0.2);
    expect(open.eyeScale).toBeGreaterThan(1);
    expect(open.eyeScale).toBeLessThanOrEqual(1.08);
  });

  it('uses separate left and right eye offsets instead of locking both eyes together', () => {
    const resolved = resolveEyeMotionForVisage('robot', 7.5, 1.1, 1, 1, 1.2);

    expect(resolved.leftEyeOffsetX).not.toBeCloseTo(resolved.rightEyeOffsetX, 6);
    expect(resolved.leftEyeOffsetY).not.toBeCloseTo(resolved.rightEyeOffsetY, 6);
  });

  it('increases brow motion for expressive states over idle', () => {
    const idle = refineBrowMotionForVisage('blackCat', 0, 0);
    const expressive = refineBrowMotionForVisage('blackCat', -6, 4);

    expect(Math.abs(expressive.browOffset)).toBeGreaterThan(Math.abs(idle.browOffset));
    expect(Math.abs(expressive.browRotation)).toBeGreaterThan(Math.abs(idle.browRotation));
  });

  it('keeps hybrid visages inside tighter crop-safe eye bounds than native svg visages', () => {
    const native = resolveEyeMotionForVisage('anonyMousse', 20, 10, 1, 1, 0.2);
    const hybrid = resolveEyeMotionForVisage('cat', 20, 10, 1, 1, 0.2);

    expect(Math.max(Math.abs(hybrid.leftEyeOffsetX), Math.abs(hybrid.rightEyeOffsetX)))
      .toBeLessThan(Math.max(Math.abs(native.leftEyeOffsetX), Math.abs(native.rightEyeOffsetX)));
    expect(Math.max(Math.abs(hybrid.leftEyeOffsetY), Math.abs(hybrid.rightEyeOffsetY)))
      .toBeLessThan(Math.max(Math.abs(native.leftEyeOffsetY), Math.abs(native.rightEyeOffsetY)));
    expect(Math.max(Math.abs(hybrid.leftEyeOffsetX), Math.abs(hybrid.rightEyeOffsetX))).toBeLessThanOrEqual(4.2);
    expect(Math.max(Math.abs(hybrid.leftEyeOffsetY), Math.abs(hybrid.rightEyeOffsetY))).toBeLessThanOrEqual(3);
  });
});
