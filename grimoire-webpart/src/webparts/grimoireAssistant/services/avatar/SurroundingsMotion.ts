import {
  AVATAR_ACTION_CUE_DURATIONS_MS,
  type AvatarActionCueType
} from './AvatarActionCue';
import type { Expression } from './ExpressionEngine';

const TWO_PI = Math.PI * 2;
const VIEWPORT_PADDING = 8;
const RECENT_CUE_DECAY_MS = 320;

export const MOBILE_SURROUNDINGS_NODE_COUNT = 12;
export const DESKTOP_SURROUNDINGS_NODE_COUNT = 18;
export const SURROUNDINGS_MOBILE_BREAKPOINT_PX = 768;

export interface ISurroundingsViewport {
  width: number;
  height: number;
}

export interface ISurroundingsBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ISurroundingsPlacement {
  centerX: number;
  centerY: number;
  coreRadiusX: number;
  coreRadiusY: number;
  emissionRadiusX: number;
  emissionRadiusY: number;
  maxRadiusX: number;
  maxRadiusY: number;
  viewportPadding: number;
}

export type SurroundingsNodeKind = 'orbiter' | 'detached';
export type SurroundingsColorRole = 'primary' | 'secondary' | 'feature';

export interface ISurroundingsNode {
  id: number;
  kind: SurroundingsNodeKind;
  colorRole: SurroundingsColorRole;
  radius: number;
  anchorX: number;
  anchorY: number;
  outwardX: number;
  outwardY: number;
  tangentX: number;
  tangentY: number;
  tangentAmplitude: number;
  radialAmplitude: number;
  peelDistance: number;
  orbitSpeed: number;
  peelSpeed: number;
  pulseSpeed: number;
  phaseA: number;
  phaseB: number;
  phaseC: number;
  baseOpacity: number;
  opacityJitter: number;
}

export interface ISurroundingsCueState {
  activeType?: AvatarActionCueType;
  activeElapsedMs?: number;
  recentType?: AvatarActionCueType;
  recentElapsedMs?: number;
}

export interface ISurroundingsIntensity {
  base: number;
  cue: number;
  total: number;
}

export interface ISurroundingsNodeFrame {
  x: number;
  y: number;
  scale: number;
  opacity: number;
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
      return 0.22;
    case 'focus':
      return 0.18;
    case 'chat':
    default:
      return 0.16;
  }
}

function activeCueLift(type: AvatarActionCueType | undefined, elapsedMs: number | undefined): number {
  if (!type || elapsedMs === undefined) return 0;
  const duration = AVATAR_ACTION_CUE_DURATIONS_MS[type];
  const progress = clamp01(elapsedMs / Math.max(1, duration));
  const wave = Math.sin(progress * Math.PI);
  return cuePeak(type) * (0.45 + (wave * 0.55));
}

function recentCueLift(type: AvatarActionCueType | undefined, elapsedMs: number | undefined): number {
  if (!type || elapsedMs === undefined || elapsedMs <= 0) return 0;
  const decay = clamp01(1 - (elapsedMs / RECENT_CUE_DECAY_MS));
  return cuePeak(type) * 0.55 * decay * decay;
}

export function isCompactSurroundingsViewport(widthPx: number): boolean {
  return widthPx < SURROUNDINGS_MOBILE_BREAKPOINT_PX;
}

export function resolveSurroundingsNodeCount(isCompactViewport: boolean): number {
  return isCompactViewport ? MOBILE_SURROUNDINGS_NODE_COUNT : DESKTOP_SURROUNDINGS_NODE_COUNT;
}

export function resolveSurroundingsPlacement(
  bounds: ISurroundingsBounds,
  viewport: ISurroundingsViewport
): ISurroundingsPlacement {
  const halfWidth = Math.max(1, bounds.width * 0.5);
  const halfHeight = Math.max(1, bounds.height * 0.5);
  const centerX = bounds.x + halfWidth;
  const centerY = bounds.y + halfHeight;

  const availableInsetX = Math.max(
    12,
    Math.min(centerX, viewport.width - centerX) - halfWidth
  );
  const availableInsetY = Math.max(
    12,
    Math.min(centerY, viewport.height - centerY) - halfHeight
  );

  const maxRadiusX = Math.max(
    1,
    Math.min(centerX - VIEWPORT_PADDING, viewport.width - centerX - VIEWPORT_PADDING)
  );
  const maxRadiusY = Math.max(
    1,
    Math.min(centerY - VIEWPORT_PADDING, viewport.height - centerY - VIEWPORT_PADDING)
  );

  const desiredCoreRadiusX = halfWidth + clamp(availableInsetX * 0.34, 10, 22);
  const desiredCoreRadiusY = halfHeight + clamp(availableInsetY * 0.34, 8, 20);
  const coreRadiusX = Math.min(desiredCoreRadiusX, Math.max(1, maxRadiusX - 6));
  const coreRadiusY = Math.min(desiredCoreRadiusY, Math.max(1, maxRadiusY - 6));

  const emissionRadiusX = Math.min(
    maxRadiusX,
    coreRadiusX + clamp(availableInsetX * 0.18, 6, 14)
  );
  const emissionRadiusY = Math.min(
    maxRadiusY,
    coreRadiusY + clamp(availableInsetY * 0.18, 5, 12)
  );

  return {
    centerX,
    centerY,
    coreRadiusX,
    coreRadiusY,
    emissionRadiusX,
    emissionRadiusY,
    maxRadiusX,
    maxRadiusY,
    viewportPadding: VIEWPORT_PADDING
  };
}

export function createSurroundingsNodes(
  seedKey: string,
  bounds: ISurroundingsBounds,
  viewport: ISurroundingsViewport,
  isCompactViewport: boolean
): ISurroundingsNode[] {
  const count = resolveSurroundingsNodeCount(isCompactViewport);
  const placement = resolveSurroundingsPlacement(bounds, viewport);
  const rnd = mulberry32(hash32(`${seedKey}|surroundings`));
  const orbiterCount = Math.max(1, Math.round(count * 0.7));
  const nodes: ISurroundingsNode[] = [];

  for (let i = 0; i < count; i++) {
    const kind: SurroundingsNodeKind = i < orbiterCount ? 'orbiter' : 'detached';
    const angle = rnd() * TWO_PI;
    const outwardX = Math.cos(angle);
    const outwardY = Math.sin(angle);
    const tangentX = -outwardY;
    const tangentY = outwardX;

    const radiusBias = kind === 'orbiter'
      ? rnd() * 0.45
      : 0.38 + (rnd() * 0.62);

    const ringRadiusX = lerp(placement.emissionRadiusX, placement.maxRadiusX, radiusBias);
    const ringRadiusY = lerp(placement.emissionRadiusY, placement.maxRadiusY, radiusBias);
    const anchorX = placement.centerX + (outwardX * ringRadiusX);
    const anchorY = placement.centerY + (outwardY * ringRadiusY);

    const colorRoll = rnd();
    const colorRole: SurroundingsColorRole = colorRoll < 0.58
      ? 'primary'
      : colorRoll < 0.88
        ? 'secondary'
        : 'feature';

    nodes.push({
      id: i,
      kind,
      colorRole,
      radius: kind === 'orbiter'
        ? lerp(2.8, 5.2, rnd())
        : lerp(2.1, 4.1, rnd()),
      anchorX: clamp(anchorX, VIEWPORT_PADDING, viewport.width - VIEWPORT_PADDING),
      anchorY: clamp(anchorY, VIEWPORT_PADDING, viewport.height - VIEWPORT_PADDING),
      outwardX,
      outwardY,
      tangentX,
      tangentY,
      tangentAmplitude: kind === 'orbiter'
        ? lerp(5.2, 10.5, rnd())
        : lerp(3.4, 7.8, rnd()),
      radialAmplitude: kind === 'orbiter'
        ? lerp(2.1, 5.1, rnd())
        : lerp(1.4, 3.4, rnd()),
      peelDistance: kind === 'orbiter'
        ? lerp(4.5, 11, rnd())
        : lerp(14, 30, rnd()),
      orbitSpeed: lerp(0.52, 1.05, rnd()),
      peelSpeed: lerp(0.7, 1.5, rnd()),
      pulseSpeed: lerp(0.95, 2.1, rnd()),
      phaseA: rnd() * TWO_PI,
      phaseB: rnd() * TWO_PI,
      phaseC: rnd() * TWO_PI,
      baseOpacity: kind === 'orbiter'
        ? lerp(0.26, 0.42, rnd())
        : lerp(0.22, 0.36, rnd()),
      opacityJitter: kind === 'orbiter'
        ? lerp(0.12, 0.2, rnd())
        : lerp(0.14, 0.22, rnd())
    });
  }

  return nodes;
}

export function resolveSurroundingsIntensity(
  expression: Expression,
  cueState: ISurroundingsCueState,
  realSpeechActive: boolean
): ISurroundingsIntensity {
  let base = 0.3;

  switch (expression) {
    case 'idle':
      base = 0.52;
      break;
    case 'listening':
      base = 0.58;
      break;
    case 'thinking':
      base = 0.68;
      break;
    case 'speaking':
      base = realSpeechActive ? 0.2 : 0.28;
      break;
    case 'happy':
      base = 0.38;
      break;
    case 'confused':
      base = 0.4;
      break;
    case 'surprised':
      base = 0.36;
      break;
    default:
      base = 0.3;
      break;
  }

  const cue = Math.max(
    activeCueLift(cueState.activeType, cueState.activeElapsedMs),
    recentCueLift(cueState.recentType, cueState.recentElapsedMs)
  );

  const unclamped = clamp01(base + cue);
  const total = expression === 'speaking' && realSpeechActive
    ? Math.min(0.18, unclamped)
    : unclamped;

  return {
    base,
    cue,
    total
  };
}

export function sampleSurroundingsNodeFrame(
  node: ISurroundingsNode,
  nowMs: number,
  intensity: number,
  viewport: ISurroundingsViewport
): ISurroundingsNodeFrame {
  const tSeconds = Math.max(0, nowMs) * 0.001;
  const orbitWave = Math.sin((tSeconds * node.orbitSpeed) + node.phaseA);
  const radialWave = Math.cos((tSeconds * (node.orbitSpeed * 0.72)) + node.phaseB);
  const peelWave = 0.5 + (0.5 * Math.sin((tSeconds * node.peelSpeed) + node.phaseC));
  const pulseWave = 0.5 + (0.5 * Math.sin((tSeconds * node.pulseSpeed) + node.phaseB));

  const peelDistance = node.kind === 'detached'
    ? node.peelDistance * Math.pow(peelWave, 1.6)
    : node.peelDistance * Math.max(0, peelWave - 0.34) * 0.45;

  const x = clamp(
    node.anchorX
      + (node.tangentX * orbitWave * node.tangentAmplitude)
      + (node.outwardX * radialWave * node.radialAmplitude)
      + (node.outwardX * peelDistance),
    VIEWPORT_PADDING,
    viewport.width - VIEWPORT_PADDING
  );

  const y = clamp(
    node.anchorY
      + (node.tangentY * orbitWave * node.tangentAmplitude)
      + (node.outwardY * radialWave * node.radialAmplitude)
      + (node.outwardY * peelDistance),
    VIEWPORT_PADDING,
    viewport.height - VIEWPORT_PADDING
  );

  return {
    x,
    y,
    scale: node.kind === 'detached'
      ? 0.98 + (peelWave * 0.34) + (intensity * 0.16)
      : 1 + (Math.max(0, radialWave) * 0.16) + (intensity * 0.12),
    opacity: clamp01(
      (node.baseOpacity * 0.42)
      + (intensity * (node.baseOpacity + (pulseWave * node.opacityJitter)))
    )
  };
}
