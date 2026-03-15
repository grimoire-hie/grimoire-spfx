export interface IAvatarMotionV2Config {
  blinkV2: boolean;
  speechMouthV2: boolean;
  ambientV2: boolean;
  ambientAuraV1: boolean;
  surroundingsV2: boolean;
  twinklesV1: boolean;
  parallaxV1: boolean;
  expressionTuningV2: boolean;
}

/**
 * Internal rollout flags for avatar motion tuning.
 * Keep all enabled by default; set individual fields to false for fast rollback.
 */
export const AVATAR_MOTION_V2: IAvatarMotionV2Config = {
  blinkV2: true,
  speechMouthV2: true,
  ambientV2: true,
  ambientAuraV1: true,
  surroundingsV2: false,
  twinklesV1: false,
  parallaxV1: false,
  expressionTuningV2: true
};
