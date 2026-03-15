export interface IAmbientSample {
  x: number;
  y: number;
  scale: number;
  opacityOffset: number;
}

export const BLINK_INTERVAL_MIN_MS = 2400;
export const BLINK_INTERVAL_MAX_MS = 7200;
export const BLINK_CLOSE_MS = 85;
export const BLINK_HOLD_MS = 45;
export const BLINK_OPEN_MS = 125;
export const DOUBLE_BLINK_GAP_MIN_MS = 90;
export const DOUBLE_BLINK_GAP_MAX_MS = 210;
export const DOUBLE_BLINK_PROBABILITY = 0.11;

const TWO_PI = Math.PI * 2;

interface IBlinkCycle {
  startMs: number;
  closeMs: number;
  holdMs: number;
  openMs: number;
}

function clamp01(value: number): number {
  if (value < 0) return 0;
  if (value > 1) return 1;
  return value;
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

export class AvatarMotionProfile {
  private readonly rnd: () => number;

  private readonly ambientPhaseA: number;
  private readonly ambientPhaseB: number;
  private readonly ambientPhaseC: number;
  private readonly ambientPhaseD: number;

  private readonly ambientAmpX: number;
  private readonly ambientAmpY: number;
  private readonly ambientScaleAmp: number;
  private readonly ambientOpacityAmp: number;

  private readonly rootFloatAmplitude: number;
  private readonly rootPulseAmplitude: number;
  private readonly eyeJitterAmplitude: number;

  private nextBlinkAtMs: number = 0;
  private activeBlink: IBlinkCycle | undefined = undefined;
  private pendingDoubleBlink: boolean = false;

  public constructor(seedKey: string) {
    const seed = hash32(seedKey || 'avatar-motion');
    this.rnd = mulberry32(seed);

    this.ambientPhaseA = this.rnd() * TWO_PI;
    this.ambientPhaseB = this.rnd() * TWO_PI;
    this.ambientPhaseC = this.rnd() * TWO_PI;
    this.ambientPhaseD = this.rnd() * TWO_PI;

    this.ambientAmpX = 0.75 + (this.rnd() * 0.55);
    this.ambientAmpY = 0.55 + (this.rnd() * 0.45);
    this.ambientScaleAmp = 0.004 + (this.rnd() * 0.004);
    this.ambientOpacityAmp = 0.055 + (this.rnd() * 0.03);

    this.rootFloatAmplitude = 1.4 + (this.rnd() * 0.9);
    this.rootPulseAmplitude = 0.004 + (this.rnd() * 0.004);
    this.eyeJitterAmplitude = 0.85 + (this.rnd() * 0.5);

    this.scheduleNextBlink(0);
  }

  public sampleBlink(nowMs: number): number {
    const currentNowMs = Math.max(0, nowMs);
    if (!this.activeBlink && currentNowMs >= this.nextBlinkAtMs) {
      this.startBlink(currentNowMs);
    }

    const cycle = this.activeBlink;
    if (!cycle) return 0;

    const elapsed = currentNowMs - cycle.startMs;
    if (elapsed < 0) return 0;

    const closeEnd = cycle.closeMs;
    const holdEnd = closeEnd + cycle.holdMs;
    const openEnd = holdEnd + cycle.openMs;

    if (elapsed <= closeEnd) {
      return clamp01(elapsed / Math.max(1, cycle.closeMs));
    }
    if (elapsed <= holdEnd) {
      return 1;
    }
    if (elapsed <= openEnd) {
      const t = (elapsed - holdEnd) / Math.max(1, cycle.openMs);
      return clamp01(1 - t);
    }

    this.activeBlink = undefined;
    if (this.pendingDoubleBlink) {
      this.pendingDoubleBlink = false;
      const quickGap = lerp(DOUBLE_BLINK_GAP_MIN_MS, DOUBLE_BLINK_GAP_MAX_MS, this.rnd());
      this.nextBlinkAtMs = currentNowMs + quickGap;
    } else {
      this.scheduleNextBlink(currentNowMs);
    }
    return 0;
  }

  public sampleAmbient(nowMs: number): IAmbientSample {
    const tSeconds = Math.max(0, nowMs) * 0.001;
    const x = Math.sin((tSeconds * 0.65) + this.ambientPhaseA) * this.ambientAmpX;
    const y = Math.cos((tSeconds * 0.51) + this.ambientPhaseB) * this.ambientAmpY;
    const scale = 1 + (Math.sin((tSeconds * 0.37) + this.ambientPhaseC) * this.ambientScaleAmp);
    const opacityOffset = Math.sin((tSeconds * 0.93) + this.ambientPhaseD) * this.ambientOpacityAmp;
    return { x, y, scale, opacityOffset };
  }

  public getRootFloatAmplitude(): number {
    return this.rootFloatAmplitude;
  }

  public getRootPulseAmplitude(): number {
    return this.rootPulseAmplitude;
  }

  public getEyeJitterAmplitude(): number {
    return this.eyeJitterAmplitude;
  }

  private startBlink(nowMs: number): void {
    this.activeBlink = {
      startMs: nowMs,
      closeMs: BLINK_CLOSE_MS,
      holdMs: BLINK_HOLD_MS,
      openMs: BLINK_OPEN_MS
    };
    this.pendingDoubleBlink = this.rnd() < DOUBLE_BLINK_PROBABILITY;
  }

  private scheduleNextBlink(nowMs: number): void {
    const interval = lerp(BLINK_INTERVAL_MIN_MS, BLINK_INTERVAL_MAX_MS, this.rnd());
    this.nextBlinkAtMs = nowMs + interval;
  }
}

