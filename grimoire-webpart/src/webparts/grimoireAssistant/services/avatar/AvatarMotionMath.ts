interface IMouthTransformParams {
  openness: number;
  width: number;
  round: number;
  mouthLift: number;
  mouthWidthBoost: number;
  mouthOpenBoost: number;
  factor: number;
}

export interface IMouthTransform {
  scaleX: number;
  scaleY: number;
  liftY: number;
  jawOffsetY: number;
}

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function clamp01(value: number): number {
  return clamp(value, 0, 1);
}

export function composeEyeScale(
  expressionEyeScale: number,
  blinkAmount: number,
  speakingActive: boolean
): number {
  const blink = clamp01(blinkAmount);
  const scaled = expressionEyeScale * (1 - (0.92 * blink));
  return speakingActive
    ? Math.max(0.14, scaled)
    : Math.max(0.08, scaled);
}

export function computeMouthTransform(
  params: IMouthTransformParams,
  useSpeechMouthV2: boolean
): IMouthTransform {
  const combinedOpen = clamp01(params.openness + params.mouthOpenBoost);

  if (!useSpeechMouthV2) {
    const scaleX = clamp(
      1 + ((params.width - 0.5) * 0.8) + params.mouthWidthBoost,
      0.55,
      1.7
    );
    const scaleY = clamp(0.92 + (combinedOpen * 1.45), 0.5, 2.0);
    return {
      scaleX,
      scaleY,
      liftY: params.mouthLift * params.factor,
      jawOffsetY: 0
    };
  }

  const openCurve = Math.pow(combinedOpen, 0.88);
  const scaleX = clamp(
    1 + ((params.width - 0.5) * 0.95) + params.mouthWidthBoost - (params.round * 0.22),
    0.52,
    1.75
  );
  const scaleY = clamp(
    0.86 + (openCurve * 1.52) + (params.round * 0.08),
    0.48,
    2.05
  );

  return {
    scaleX,
    scaleY,
    liftY: (params.mouthLift * params.factor) - (openCurve * 0.35 * params.factor),
    jawOffsetY: openCurve * 0.62 * params.factor
  };
}
