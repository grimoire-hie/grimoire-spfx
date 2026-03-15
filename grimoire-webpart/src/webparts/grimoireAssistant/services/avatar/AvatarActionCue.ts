/**
 * AvatarActionCue
 * Deterministic micro-animation cues for header actions (Focus/Summarize/Chat).
 */

export type AvatarActionCueType = 'focus' | 'summarize' | 'chat';

export interface IAvatarActionCue {
  id: number;
  type: AvatarActionCueType;
  at: number;
}

export interface IAvatarCueMouthParams {
  openness: number;
  width: number;
  round: number;
}

export interface IAvatarActionCueFrame {
  eyeOffsetXAdd: number;
  rootScaleX: number;
  rootScaleY: number;
  rootYOffsetAdd: number;
  mouthOverride?: IAvatarCueMouthParams;
  finished: boolean;
}

export const AVATAR_ACTION_CUE_DURATIONS_MS: Record<AvatarActionCueType, number> = {
  focus: 900,
  summarize: 780,
  chat: 1200
};

const NEUTRAL_FRAME: IAvatarActionCueFrame = {
  eyeOffsetXAdd: 0,
  rootScaleX: 1,
  rootScaleY: 1,
  rootYOffsetAdd: 0,
  finished: false
};

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function clamp01(value: number): number {
  return clamp(value, 0, 1);
}

function cloneNeutral(finished: boolean): IAvatarActionCueFrame {
  return {
    eyeOffsetXAdd: NEUTRAL_FRAME.eyeOffsetXAdd,
    rootScaleX: NEUTRAL_FRAME.rootScaleX,
    rootScaleY: NEUTRAL_FRAME.rootScaleY,
    rootYOffsetAdd: NEUTRAL_FRAME.rootYOffsetAdd,
    finished
  };
}

/**
 * Evaluate current frame modifiers for a given cue.
 * Output values are intentionally subtle and in avatar-local units.
 */
export function evaluateAvatarActionCue(
  type: AvatarActionCueType,
  elapsedMs: number,
  realSpeechActive: boolean
): IAvatarActionCueFrame {
  const duration = AVATAR_ACTION_CUE_DURATIONS_MS[type];
  const safeElapsed = Math.max(0, elapsedMs);
  if (safeElapsed >= duration) return cloneNeutral(true);

  const progress = clamp01(safeElapsed / duration);

  if (type === 'focus') {
    // Two rightward glances over 900ms.
    const glance = Math.max(0, Math.sin(progress * Math.PI * 4));
    return {
      eyeOffsetXAdd: glance * 6.4,
      rootScaleX: 1,
      rootScaleY: 1,
      rootYOffsetAdd: 0,
      finished: false
    };
  }

  if (type === 'summarize') {
    // Two short squash/rebound beats.
    const wave = Math.sin(progress * Math.PI * 4);
    const compress = Math.max(0, wave);
    const rebound = Math.max(0, -wave);
    return {
      eyeOffsetXAdd: 0,
      rootScaleX: 1 + (compress * 0.085) - (rebound * 0.025),
      rootScaleY: 1 - (compress * 0.16) + (rebound * 0.05),
      rootYOffsetAdd: (compress * 2.6) - (rebound * 0.85),
      finished: false
    };
  }

  // chat
  if (realSpeechActive) {
    return cloneNeutral(false);
  }

  const talkA = 0.5 + (Math.sin((safeElapsed * 0.03) + 0.8) * 0.5);
  const talkB = 0.5 + (Math.sin((safeElapsed * 0.017) + 1.2) * 0.5);
  const envelope = 1 - (progress * 0.18);

  return {
    eyeOffsetXAdd: 0,
    rootScaleX: 1,
    rootScaleY: 1,
    rootYOffsetAdd: 0,
    mouthOverride: {
      openness: clamp01((0.14 + (talkA * 0.42)) * envelope),
      width: clamp01(0.42 + (talkB * 0.26)),
      round: clamp01(0.08 + ((1 - talkB) * 0.35))
    },
    finished: false
  };
}
