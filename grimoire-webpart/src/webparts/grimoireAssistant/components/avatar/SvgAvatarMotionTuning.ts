import type { VisageMode } from '../../services/avatar/FaceTemplateData';

export interface IResolvedEyeMotion {
  leftEyeOffsetX: number;
  rightEyeOffsetX: number;
  leftEyeOffsetY: number;
  rightEyeOffsetY: number;
  eyeScale: number;
}

interface IEyeMotionTuning {
  horizontalScale: number;
  verticalScale: number;
  blinkCollapse: number;
  minEyeScale: number;
  openBoost: number;
  maxEyeScale: number;
  gazeLead: number;
  idleSpreadX: number;
  idleSpreadY: number;
  maxOffsetX: number;
  maxOffsetY: number;
}

interface IBrowMotionTuning {
  offsetScale: number;
  rotationScale: number;
  maxOffset: number;
  maxRotation: number;
}

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

const DEFAULT_EYE_TUNING: IEyeMotionTuning = {
  horizontalScale: 1.18,
  verticalScale: 1.1,
  blinkCollapse: 0.92,
  minEyeScale: 0.12,
  openBoost: 0.28,
  maxEyeScale: 1.08,
  gazeLead: 0.22,
  idleSpreadX: 0.08,
  idleSpreadY: 0.05,
  maxOffsetX: 9.5,
  maxOffsetY: 4.8
};

const EYE_TUNING: Record<VisageMode, IEyeMotionTuning> = {
  classic: {
    horizontalScale: 1.08,
    verticalScale: 1.06,
    blinkCollapse: 0.92,
    minEyeScale: 0.14,
    openBoost: 0.24,
    maxEyeScale: 1.06,
    gazeLead: 0.16,
    idleSpreadX: 0.05,
    idleSpreadY: 0.03,
    maxOffsetX: 7.8,
    maxOffsetY: 4.2
  },
  anonyMousse: {
    horizontalScale: 1.22,
    verticalScale: 1.08,
    blinkCollapse: 0.96,
    minEyeScale: 0.1,
    openBoost: 0.28,
    maxEyeScale: 1.08,
    gazeLead: 0.24,
    idleSpreadX: 0.08,
    idleSpreadY: 0.05,
    maxOffsetX: 9.2,
    maxOffsetY: 4.8
  },
  robot: {
    horizontalScale: 1.12,
    verticalScale: 1.02,
    blinkCollapse: 0.94,
    minEyeScale: 0.14,
    openBoost: 0.22,
    maxEyeScale: 1.05,
    gazeLead: 0.18,
    idleSpreadX: 0.05,
    idleSpreadY: 0.03,
    maxOffsetX: 7.4,
    maxOffsetY: 3.8
  },
  cat: {
    horizontalScale: 0.44,
    verticalScale: 0.38,
    blinkCollapse: 0.92,
    minEyeScale: 0.18,
    openBoost: 0.18,
    maxEyeScale: 1.04,
    gazeLead: 0.09,
    idleSpreadX: 0.028,
    idleSpreadY: 0.02,
    maxOffsetX: 4.2,
    maxOffsetY: 3
  },
  blackCat: {
    horizontalScale: 0.54,
    verticalScale: 0.42,
    blinkCollapse: 0.94,
    minEyeScale: 0.17,
    openBoost: 0.18,
    maxEyeScale: 1.04,
    gazeLead: 0.1,
    idleSpreadX: 0.03,
    idleSpreadY: 0.02,
    maxOffsetX: 4.8,
    maxOffsetY: 3.2
  },
  squirrel: {
    horizontalScale: 0.46,
    verticalScale: 0.4,
    blinkCollapse: 0.9,
    minEyeScale: 0.2,
    openBoost: 0.18,
    maxEyeScale: 1.04,
    gazeLead: 0.09,
    idleSpreadX: 0.03,
    idleSpreadY: 0.02,
    maxOffsetX: 4.4,
    maxOffsetY: 3.1
  },
  particleVisage: {
    horizontalScale: 0.52,
    verticalScale: 0.32,
    blinkCollapse: 0.94,
    minEyeScale: 0.16,
    openBoost: 0.18,
    maxEyeScale: 1.05,
    gazeLead: 0.11,
    idleSpreadX: 0.03,
    idleSpreadY: 0.02,
    maxOffsetX: 5,
    maxOffsetY: 2.6
  }
};

const DEFAULT_BROW_TUNING: IBrowMotionTuning = {
  offsetScale: 1.08,
  rotationScale: 1.08,
  maxOffset: 12,
  maxRotation: 8
};

const BROW_TUNING: Record<VisageMode, IBrowMotionTuning> = {
  classic: {
    offsetScale: 1.14,
    rotationScale: 1.1,
    maxOffset: 11,
    maxRotation: 7
  },
  anonyMousse: {
    offsetScale: 1.2,
    rotationScale: 1.16,
    maxOffset: 12,
    maxRotation: 8
  },
  robot: {
    offsetScale: 1.1,
    rotationScale: 1.06,
    maxOffset: 9,
    maxRotation: 6
  },
  cat: {
    offsetScale: 0.78,
    rotationScale: 0.72,
    maxOffset: 7,
    maxRotation: 4.8
  },
  blackCat: {
    offsetScale: 0.82,
    rotationScale: 0.76,
    maxOffset: 7.5,
    maxRotation: 5.2
  },
  squirrel: {
    offsetScale: 0.86,
    rotationScale: 0.8,
    maxOffset: 7.8,
    maxRotation: 5.4
  },
  particleVisage: {
    offsetScale: 0.7,
    rotationScale: 0.62,
    maxOffset: 6,
    maxRotation: 4.5
  }
};

export function resolveEyeMotionForVisage(
  visage: VisageMode,
  eyeOffsetX: number,
  eyeOffsetY: number,
  eyeScale: number,
  factor: number,
  phase: number = 0
): IResolvedEyeMotion {
  const tuning = EYE_TUNING[visage] || DEFAULT_EYE_TUNING;
  const maxOffsetX = tuning.maxOffsetX * factor;
  const maxOffsetY = tuning.maxOffsetY * factor;
  const baseX = clamp(eyeOffsetX * tuning.horizontalScale, -maxOffsetX, maxOffsetX);
  const baseY = clamp(eyeOffsetY * tuning.verticalScale, -maxOffsetY, maxOffsetY);
  const resolvedScale = eyeScale < 1
    ? clamp(1 - ((1 - eyeScale) * tuning.blinkCollapse), tuning.minEyeScale, 1)
    : clamp(1 + ((eyeScale - 1) * tuning.openBoost), 1, tuning.maxEyeScale);
  const gazeStrength = factor > 0
    ? clamp(baseX / Math.max(factor * 7.8, 0.001), -1, 1)
    : 0;
  const pairedShiftX = (
    (gazeStrength * tuning.gazeLead)
    + (Math.sin((phase * 0.83) + 0.7) * tuning.idleSpreadX)
  ) * factor;
  const pairedShiftY = Math.cos((phase * 0.69) + 0.4) * tuning.idleSpreadY * factor;

  return {
    leftEyeOffsetX: clamp(baseX - pairedShiftX, -maxOffsetX, maxOffsetX),
    rightEyeOffsetX: clamp(baseX + pairedShiftX, -maxOffsetX, maxOffsetX),
    leftEyeOffsetY: clamp(baseY - pairedShiftY, -maxOffsetY, maxOffsetY),
    rightEyeOffsetY: clamp(baseY + pairedShiftY, -maxOffsetY, maxOffsetY),
    eyeScale: resolvedScale
  };
}

export function refineBrowMotionForVisage(
  visage: VisageMode,
  browOffset: number,
  browRotation: number
): { browOffset: number; browRotation: number } {
  const tuning = BROW_TUNING[visage] || DEFAULT_BROW_TUNING;
  return {
    browOffset: clamp(browOffset * tuning.offsetScale, -tuning.maxOffset, tuning.maxOffset),
    browRotation: clamp(browRotation * tuning.rotationScale, -tuning.maxRotation, tuning.maxRotation)
  };
}
