import {
  AVATAR_ACTION_CUE_DURATIONS_MS,
  evaluateAvatarActionCue
} from './AvatarActionCue';

describe('AvatarActionCue', () => {
  it('focus cue pushes eyes right during the cue window', () => {
    const frame = evaluateAvatarActionCue('focus', 120, false);
    expect(frame.eyeOffsetXAdd).toBeGreaterThan(0);
    expect(frame.finished).toBe(false);
  });

  it('summarize cue compresses vertically and widens slightly', () => {
    const frame = evaluateAvatarActionCue('summarize', 120, false);
    expect(frame.rootScaleY).toBeLessThan(1);
    expect(frame.rootScaleX).toBeGreaterThan(1);
    expect(frame.finished).toBe(false);
  });

  it('chat cue generates mouth motion when no real speech is active', () => {
    const frame = evaluateAvatarActionCue('chat', 260, false);
    expect(frame.mouthOverride).toBeDefined();
    expect(frame.mouthOverride?.openness).toBeGreaterThan(0);
    expect(frame.finished).toBe(false);
  });

  it('chat cue suppresses mouth override when real speech is active', () => {
    const frame = evaluateAvatarActionCue('chat', 260, true);
    expect(frame.mouthOverride).toBeUndefined();
    expect(frame.finished).toBe(false);
  });

  it('all cue types report finished at or beyond configured duration', () => {
    (['focus', 'summarize', 'chat'] as Array<'focus' | 'summarize' | 'chat'>).forEach((type) => {
      const frame = evaluateAvatarActionCue(type, AVATAR_ACTION_CUE_DURATIONS_MS[type], false);
      expect(frame.finished).toBe(true);
    });
  });
});
