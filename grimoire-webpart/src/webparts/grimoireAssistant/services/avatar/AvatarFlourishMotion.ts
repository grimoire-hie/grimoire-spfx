import {
  AVATAR_ACTION_CUE_DURATIONS_MS,
  type AvatarActionCueType
} from './AvatarActionCue';
import type { Expression } from './ExpressionEngine';
import type {
  ISurroundingsBounds,
  ISurroundingsViewport,
  SurroundingsColorRole
} from './SurroundingsMotion';

const TWO_PI = Math.PI * 2;
const TWINKLE_RECENT_CUE_DECAY_MS = 920;
const VIEWPORT_PADDING = 12;

export const MOBILE_TWINKLE_NODE_COUNT = 3;
export const DESKTOP_TWINKLE_NODE_COUNT = 5;

export interface ITwinkleAnchor {
  x: number;
  y: number;
  swayX: number;
  swayY: number;
  cueAffinity: number;
  colorRole: SurroundingsColorRole;
}

export interface ITwinkleNode {
  id: number;
  anchorX: number;
  anchorY: number;
  swayX: number;
  swayY: number;
  cueAffinity: number;
  colorRole: SurroundingsColorRole;
  baseScale: number;
  driftSpeed: number;
  shimmerSpeed: number;
  pulseSpeed: number;
  sparkleSharpness: number;
  phaseA: number;
  phaseB: number;
  phaseC: number;
}

export interface ITwinkleCueState {
  activeType?: AvatarActionCueType;
  activeElapsedMs?: number;
  recentType?: AvatarActionCueType;
  recentElapsedMs?: number;
}

export interface ITwinkleIntensity {
  base: number;
  cue: number;
  total: number;
}

export interface ITwinkleFrame {
  x: number;
  y: number;
  scale: number;
  opacity: number;
}

export interface IClassicParallaxFrame {
  haloX: number;
  haloY: number;
  haloScale: number;
  pagesX: number;
  pagesY: number;
  pagesScale: number;
  coversX: number;
  coversY: number;
  linesX: number;
  linesY: number;
  stackX: number;
  stackY: number;
}

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function clamp01(value: number): number {
  return clamp(value, 0, 1);
}

function lerp(min: number, max: number, t: number): number {
  return min + ((max - min) * t);
}

function hash32(input: string): number {
  let hash = 2166136261;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function mulberry32(seed: number): () => number {
  let state = seed | 0;
  return (): number => {
    state = (state + 0x6d2b79f5) | 0;
    let t = Math.imul(state ^ (state >>> 15), 1 | state);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function cuePeak(type: AvatarActionCueType): number {
  switch (type) {
    case 'summarize':
      return 0.34;
    case 'focus':
      return 0.28;
    case 'chat':
    default:
      return 0.22;
  }
}

function activeCueLift(type: AvatarActionCueType | undefined, elapsedMs: number | undefined): number {
  if (!type || elapsedMs === undefined) return 0;
  const duration = AVATAR_ACTION_CUE_DURATIONS_MS[type];
  const progress = clamp01(elapsedMs / Math.max(1, duration));
  const window = Math.sin(progress * Math.PI);
  return cuePeak(type) * (0.28 + (window * 0.72));
}

function recentCueLift(type: AvatarActionCueType | undefined, elapsedMs: number | undefined): number {
  if (!type || elapsedMs === undefined || elapsedMs <= 0) return 0;
  const decay = clamp01(1 - (elapsedMs / TWINKLE_RECENT_CUE_DECAY_MS));
  return cuePeak(type) * 0.7 * decay * decay;
}

export function resolveTwinkleNodeCount(isCompactViewport: boolean): number {
  return isCompactViewport ? MOBILE_TWINKLE_NODE_COUNT : DESKTOP_TWINKLE_NODE_COUNT;
}

export function buildFallbackTwinkleAnchors(
  bounds: ISurroundingsBounds,
  viewport: ISurroundingsViewport
): ITwinkleAnchor[] {
  const centerX = bounds.x + (bounds.width * 0.5);
  const centerY = bounds.y + (bounds.height * 0.5);
  const marginX = Math.max(22, bounds.width * 0.12);
  const marginY = Math.max(18, bounds.height * 0.12);

  const points: ITwinkleAnchor[] = [
    {
      x: centerX,
      y: bounds.y - marginY,
      swayX: 2.8,
      swayY: 3.8,
      cueAffinity: 1,
      colorRole: 'feature'
    },
    {
      x: bounds.x - marginX,
      y: centerY - (bounds.height * 0.18),
      swayX: 3.2,
      swayY: 2.6,
      cueAffinity: 0.8,
      colorRole: 'secondary'
    },
    {
      x: bounds.x + bounds.width + marginX,
      y: centerY - (bounds.height * 0.12),
      swayX: 3.2,
      swayY: 2.6,
      cueAffinity: 0.8,
      colorRole: 'secondary'
    },
    {
      x: bounds.x + (bounds.width * 0.1),
      y: bounds.y + bounds.height + (marginY * 0.35),
      swayX: 2.6,
      swayY: 2.4,
      cueAffinity: 0.55,
      colorRole: 'primary'
    },
    {
      x: bounds.x + (bounds.width * 0.9),
      y: bounds.y + bounds.height + (marginY * 0.35),
      swayX: 2.6,
      swayY: 2.4,
      cueAffinity: 0.55,
      colorRole: 'primary'
    },
    {
      x: centerX,
      y: bounds.y + bounds.height + (marginY * 0.7),
      swayX: 2.2,
      swayY: 2,
      cueAffinity: 0.45,
      colorRole: 'feature'
    }
  ];

  return points.map((point) => ({
    ...point,
    x: clamp(point.x, VIEWPORT_PADDING, viewport.width - VIEWPORT_PADDING),
    y: clamp(point.y, VIEWPORT_PADDING, viewport.height - VIEWPORT_PADDING)
  }));
}

export function createTwinkleNodes(
  seedKey: string,
  anchors: ITwinkleAnchor[],
  isCompactViewport: boolean
): ITwinkleNode[] {
  if (anchors.length === 0) return [];

  const count = Math.min(resolveTwinkleNodeCount(isCompactViewport), anchors.length);
  const rnd = mulberry32(hash32(`${seedKey}|twinkles`));
  const selectedAnchors = anchors.slice(0, count);

  return selectedAnchors.map((anchor, id) => ({
    id,
    anchorX: anchor.x,
    anchorY: anchor.y,
    swayX: anchor.swayX * lerp(0.9, 1.22, rnd()),
    swayY: anchor.swayY * lerp(0.9, 1.24, rnd()),
    cueAffinity: clamp01(anchor.cueAffinity * lerp(0.92, 1.08, rnd())),
    colorRole: anchor.colorRole,
    baseScale: lerp(0.92, 1.24, rnd()),
    driftSpeed: lerp(0.18, 0.42, rnd()),
    shimmerSpeed: lerp(0.42, 0.9, rnd()),
    pulseSpeed: lerp(0.86, 1.48, rnd()),
    sparkleSharpness: lerp(7.2, 10.4, rnd()),
    phaseA: rnd() * TWO_PI,
    phaseB: rnd() * TWO_PI,
    phaseC: rnd() * TWO_PI
  }));
}

export function resolveTwinkleIntensity(
  expression: Expression,
  cueState: ITwinkleCueState,
  realSpeechActive: boolean
): ITwinkleIntensity {
  let base = 0.12;

  switch (expression) {
    case 'idle':
      base = 0.18;
      break;
    case 'listening':
      base = 0.2;
      break;
    case 'thinking':
      base = 0.22;
      break;
    case 'speaking':
      base = realSpeechActive ? 0.06 : 0.08;
      break;
    case 'happy':
      base = 0.16;
      break;
    case 'confused':
      base = 0.14;
      break;
    case 'surprised':
      base = 0.15;
      break;
    default:
      base = 0.12;
      break;
  }

  const cue = Math.max(
    activeCueLift(cueState.activeType, cueState.activeElapsedMs),
    recentCueLift(cueState.recentType, cueState.recentElapsedMs)
  );

  const unclamped = clamp01(base + cue);
  const total = expression === 'speaking' && realSpeechActive
    ? Math.min(0.2, unclamped)
    : unclamped;

  return {
    base,
    cue,
    total
  };
}

export function sampleTwinkleFrame(
  node: ITwinkleNode,
  nowMs: number,
  intensity: number
): ITwinkleFrame {
  const tSeconds = Math.max(0, nowMs) * 0.001;
  const driftX = Math.sin((tSeconds * node.driftSpeed) + node.phaseA);
  const driftY = Math.cos((tSeconds * (node.driftSpeed * 0.84)) + node.phaseB);
  const shimmer = 0.5 + (0.5 * Math.sin((tSeconds * node.shimmerSpeed) + node.phaseC));
  const pulse = 0.5 + (0.5 * Math.sin((tSeconds * node.pulseSpeed) + node.phaseB));
  const sparkle = Math.pow(shimmer, node.sparkleSharpness);
  const glow = intensity * node.cueAffinity;

  return {
    x: node.anchorX + (driftX * node.swayX * (0.42 + (glow * 0.58))),
    y: node.anchorY + (driftY * node.swayY * (0.4 + (glow * 0.6))),
    scale: node.baseScale * (0.78 + (sparkle * 1.05) + (glow * 0.28) + (pulse * 0.08)),
    opacity: clamp01(
      (sparkle * (0.04 + (glow * 0.2)))
      + (pulse * glow * 0.04)
    )
  };
}

export function sampleClassicParallaxFrame(
  nowMs: number,
  intensity: number
): IClassicParallaxFrame {
  const tSeconds = Math.max(0, nowMs) * 0.001;
  const depth = 0.35 + (clamp01(intensity) * 0.65);
  const slow = Math.sin(tSeconds * 0.28);
  const medium = Math.cos((tSeconds * 0.33) + 0.9);
  const shimmer = Math.sin((tSeconds * 0.21) + 1.7);
  const stack = Math.cos((tSeconds * 0.24) + 2.2);

  return {
    haloX: slow * 0.72 * depth,
    haloY: medium * 1.2 * depth,
    haloScale: 1 + (Math.sin((tSeconds * 0.22) + 0.5) * 0.008 * depth),
    pagesX: -slow * 0.46 * depth,
    pagesY: medium * 0.52 * depth,
    pagesScale: 1 + (Math.sin((tSeconds * 0.27) + 2.3) * 0.0036 * depth),
    coversX: slow * 0.18 * depth,
    coversY: -medium * 0.16 * depth,
    linesX: shimmer * 0.92 * depth,
    linesY: Math.cos((tSeconds * 0.31) + 1.2) * 0.48 * depth,
    stackX: -shimmer * 0.22 * depth,
    stackY: stack * 0.58 * depth
  };
}
