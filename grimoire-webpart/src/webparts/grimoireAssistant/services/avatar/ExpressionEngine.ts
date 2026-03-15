/**
 * ExpressionEngine
 * State machine for facial expressions with lerped transitions.
 * Each expression defines region offsets/scales applied on top of the base face.
 */

import { FaceRegion } from './ParticleSystem';

// ─── Types ─────────────────────────────────────────────────────

export type Expression =
  | 'idle'
  | 'listening'
  | 'thinking'
  | 'speaking'
  | 'surprised'
  | 'happy'
  | 'confused';

export interface IRegionModifier {
  region: FaceRegion;
  /** Horizontal offset (normalized, relative to canvas) */
  dx: number;
  /** Vertical offset (normalized) */
  dy: number;
  /** Scale multiplier (1.0 = no change) */
  scale: number;
}

export interface IExpressionDef {
  name: Expression;
  /** Region modifiers relative to base face */
  modifiers: IRegionModifier[];
  /** Transition duration in ms */
  transitionMs: number;
}

// ─── Expression Definitions ────────────────────────────────────

const EXPRESSIONS: Record<Expression, IExpressionDef> = {
  idle: {
    name: 'idle',
    modifiers: [
      // Warm micro-smile: upper lip lifts, lower drops slightly
      { region: 'mouth_upper', dx: 0, dy: -0.005, scale: 1.0 },
      { region: 'mouth_lower', dx: 0, dy: 0.002, scale: 1.0 }
    ],
    transitionMs: 300
  },

  listening: {
    name: 'listening',
    modifiers: [
      // Eyebrows raise — open, attentive face
      { region: 'left_eyebrow', dx: 0, dy: -0.010, scale: 1.0 },
      { region: 'right_eyebrow', dx: 0, dy: -0.010, scale: 1.0 },
      // Eyes open wider — engaged
      { region: 'left_eye', dx: 0, dy: -0.004, scale: 1.0 },
      { region: 'right_eye', dx: 0, dy: -0.004, scale: 1.0 },
      // Pupils steady, forward-facing
      { region: 'left_pupil', dx: 0, dy: 0, scale: 1.0 },
      { region: 'right_pupil', dx: 0, dy: 0, scale: 1.0 },
      // Gentle warmth — light smile
      { region: 'mouth_upper', dx: 0, dy: -0.005, scale: 1.0 }
    ],
    transitionMs: 250
  },

  thinking: {
    name: 'thinking',
    modifiers: [
      // Brows lift dramatically — "that's where the brain is"
      { region: 'left_eyebrow', dx: 0, dy: -0.020, scale: 1.0 },
      { region: 'right_eyebrow', dx: 0, dy: -0.020, scale: 1.0 },
      // Eyes drift upward — looking up to think
      { region: 'left_eye', dx: 0, dy: -0.012, scale: 1.0 },
      { region: 'right_eye', dx: 0, dy: -0.012, scale: 1.0 },
      // Pupils glance up-left (classic thinking direction)
      { region: 'left_pupil', dx: -0.004, dy: -0.008, scale: 1.0 },
      { region: 'right_pupil', dx: -0.004, dy: -0.008, scale: 1.0 },
      // Mouth compresses — concentration
      { region: 'mouth_upper', dx: 0, dy: 0.003, scale: 1.0 },
      { region: 'mouth_lower', dx: 0, dy: -0.002, scale: 1.0 },
      // Ambient particles drift up — "brain activity" cloud
      { region: 'ambient', dx: 0, dy: -0.006, scale: 1.0 }
    ],
    transitionMs: 300
  },

  speaking: {
    name: 'speaking',
    modifiers: [
      // Mouth is handled by SpeechMouthAnalyzer — minimal mods here
      // Light brow engagement
      { region: 'left_eyebrow', dx: 0, dy: -0.005, scale: 1.0 },
      { region: 'right_eyebrow', dx: 0, dy: -0.005, scale: 1.0 },
      // Eyes slightly animated
      { region: 'left_eye', dx: 0, dy: -0.002, scale: 1.0 },
      { region: 'right_eye', dx: 0, dy: -0.002, scale: 1.0 }
    ],
    transitionMs: 200
  },

  surprised: {
    name: 'surprised',
    modifiers: [
      // Eyebrows shoot up — highest of any expression
      { region: 'left_eyebrow', dx: 0, dy: -0.025, scale: 1.0 },
      { region: 'right_eyebrow', dx: 0, dy: -0.025, scale: 1.0 },
      // Eyes widen dramatically
      { region: 'left_eye', dx: 0, dy: -0.010, scale: 1.0 },
      { region: 'right_eye', dx: 0, dy: -0.010, scale: 1.0 },
      // Pupils forward
      { region: 'left_pupil', dx: 0, dy: 0, scale: 1.0 },
      { region: 'right_pupil', dx: 0, dy: 0, scale: 1.0 },
      // Mouth drops open
      { region: 'mouth_upper', dx: 0, dy: -0.004, scale: 1.0 },
      { region: 'mouth_lower', dx: 0, dy: 0.014, scale: 1.0 }
    ],
    transitionMs: 120
  },

  happy: {
    name: 'happy',
    modifiers: [
      // Eyes squint downward — Duchenne smile (eye crinkles)
      { region: 'left_eye', dx: 0, dy: 0.006, scale: 1.0 },
      { region: 'right_eye', dx: 0, dy: 0.006, scale: 1.0 },
      // Eyebrows lift gently
      { region: 'left_eyebrow', dx: 0, dy: -0.008, scale: 1.0 },
      { region: 'right_eyebrow', dx: 0, dy: -0.008, scale: 1.0 },
      // Mouth beams — strong upward curve
      { region: 'mouth_upper', dx: 0, dy: -0.018, scale: 1.0 },
      { region: 'mouth_lower', dx: 0, dy: 0.008, scale: 1.0 },
      // Cheeks lift
      { region: 'chin', dx: 0, dy: -0.004, scale: 1.0 }
    ],
    transitionMs: 300
  },

  confused: {
    name: 'confused',
    modifiers: [
      // Strong asymmetric brows — left raises, right drops
      { region: 'left_eyebrow', dx: -0.004, dy: -0.014, scale: 1.0 },
      { region: 'right_eyebrow', dx: 0.003, dy: 0.005, scale: 1.0 },
      // Eyes tilted — left higher, right lower
      { region: 'left_eye', dx: -0.003, dy: -0.005, scale: 1.0 },
      { region: 'right_eye', dx: -0.003, dy: 0.004, scale: 1.0 },
      // Pupils shift off-center
      { region: 'left_pupil', dx: 0.003, dy: -0.002, scale: 1.0 },
      { region: 'right_pupil', dx: 0.003, dy: 0.002, scale: 1.0 },
      // Mouth off-center — slight grimace
      { region: 'mouth_upper', dx: 0.005, dy: 0, scale: 1.0 },
      { region: 'mouth_lower', dx: 0.005, dy: 0.003, scale: 1.0 }
    ],
    transitionMs: 350
  }
};

// ─── Engine ────────────────────────────────────────────────────

interface ILerpedModifier {
  region: FaceRegion;
  currentDx: number;
  currentDy: number;
  currentScale: number;
  targetDx: number;
  targetDy: number;
  targetScale: number;
}

export class ExpressionEngine {
  private current: Expression = 'idle';
  private lerpedModifiers: Map<FaceRegion, ILerpedModifier> = new Map();
  private transitionProgress: number = 1; // 0-1, 1 = complete
  private transitionSpeed: number = 0;

  // Gaze direction (-1..1 for each axis)
  private gazeX: number = 0;
  private gazeY: number = 0;
  private targetGazeX: number = 0;
  private targetGazeY: number = 0;
  private gazeRevertTimer: ReturnType<typeof setTimeout> | undefined;

  /** Reusable array for getModifiers() output (avoids per-frame allocation) */
  private modifiersResult: Array<{ region: FaceRegion; dx: number; dy: number; scale: number }> = [];

  /**
   * Set target expression. Transition begins immediately.
   */
  public setExpression(expr: Expression): void {
    if (expr === this.current && this.transitionProgress >= 1) return;

    this.current = expr;
    const def = EXPRESSIONS[expr];
    this.transitionProgress = 0;
    this.transitionSpeed = 1000 / def.transitionMs; // progress per second

    // Set targets for all regions mentioned in the expression
    // First, set all existing targets to zero (return to neutral)
    this.lerpedModifiers.forEach((mod) => {
      mod.targetDx = 0;
      mod.targetDy = 0;
      mod.targetScale = 1;
    });

    // Now set targets for the new expression
    for (const m of def.modifiers) {
      let lerped = this.lerpedModifiers.get(m.region);
      if (!lerped) {
        lerped = {
          region: m.region,
          currentDx: 0,
          currentDy: 0,
          currentScale: 1,
          targetDx: 0,
          targetDy: 0,
          targetScale: 1
        };
        this.lerpedModifiers.set(m.region, lerped);
      }
      lerped.targetDx = m.dx;
      lerped.targetDy = m.dy;
      lerped.targetScale = m.scale;
    }
  }

  /**
   * Set gaze direction. The face/eyes will shift toward (x,y).
   * @param x - horizontal: -1 (left) to 1 (right)
   * @param y - vertical: -1 (up) to 1 (down)
   * @param revertMs - auto-revert to center after this many ms (0 = no auto-revert)
   */
  public setGazeDirection(x: number, y: number, revertMs: number = 3000): void {
    this.targetGazeX = Math.max(-1, Math.min(1, x));
    this.targetGazeY = Math.max(-1, Math.min(1, y));
    if (this.gazeRevertTimer) clearTimeout(this.gazeRevertTimer);
    if (revertMs > 0) {
      this.gazeRevertTimer = setTimeout(() => this.clearGaze(), revertMs);
    }
  }

  /** Clear gaze — revert to center. */
  public clearGaze(): void {
    this.targetGazeX = 0;
    this.targetGazeY = 0;
    if (this.gazeRevertTimer) {
      clearTimeout(this.gazeRevertTimer);
      this.gazeRevertTimer = undefined;
    }
  }

  /**
   * Update lerped values. Call every frame.
   * @param dtSeconds - delta time in seconds
   */
  public update(dtSeconds: number): void {
    // Lerp gaze toward target
    this.gazeX += (this.targetGazeX - this.gazeX) * 0.1;
    this.gazeY += (this.targetGazeY - this.gazeY) * 0.1;

    if (this.transitionProgress >= 1) return;

    this.transitionProgress = Math.min(1, this.transitionProgress + this.transitionSpeed * dtSeconds);

    // Ease-out cubic for smooth landing
    const t = 1 - Math.pow(1 - this.transitionProgress, 3);

    this.lerpedModifiers.forEach((mod) => {
      mod.currentDx += (mod.targetDx - mod.currentDx) * t * 0.15;
      mod.currentDy += (mod.targetDy - mod.currentDy) * t * 0.15;
      mod.currentScale += (mod.targetScale - mod.currentScale) * t * 0.15;
    });
  }

  /**
   * Get current interpolated modifiers to apply to particle system.
   * Returns modifiers scaled to pixel offsets for a given canvas size.
   */
  public getModifiers(canvasWidth: number, canvasHeight: number): Array<{
    region: FaceRegion;
    dx: number;
    dy: number;
    scale: number;
  }> {
    // Reuse array — reset length to avoid per-frame allocation
    let idx = 0;
    const result = this.modifiersResult;
    const hasGaze = Math.abs(this.gazeX) > 0.001 || Math.abs(this.gazeY) > 0.001;

    this.lerpedModifiers.forEach((mod) => {
      // Compute gaze offset for this region
      let gazeDx = 0;
      let gazeDy = 0;
      if (hasGaze) {
        const r = mod.region;
        if (r === 'left_pupil' || r === 'right_pupil') {
          gazeDx = this.gazeX * 0.016;
          gazeDy = this.gazeY * 0.012;
        } else if (r === 'left_eye' || r === 'right_eye') {
          gazeDx = this.gazeX * 0.010;
          gazeDy = this.gazeY * 0.008;
        } else {
          gazeDx = this.gazeX * 0.005;
          gazeDy = this.gazeY * 0.004;
        }
      }

      const totalDx = mod.currentDx + gazeDx;
      const totalDy = mod.currentDy + gazeDy;

      // Only include if non-trivial
      if (
        Math.abs(totalDx) > 0.0001 ||
        Math.abs(totalDy) > 0.0001 ||
        Math.abs(mod.currentScale - 1) > 0.001
      ) {
        if (idx < result.length) {
          result[idx].region = mod.region;
          result[idx].dx = totalDx * canvasWidth;
          result[idx].dy = totalDy * canvasHeight;
          result[idx].scale = mod.currentScale;
        } else {
          result.push({
            region: mod.region,
            dx: totalDx * canvasWidth,
            dy: totalDy * canvasHeight,
            scale: mod.currentScale
          });
        }
        idx++;
      }
    });

    // If gaze is active but a region has no lerpedModifier, we still need
    // to emit gaze offsets for pupil/eye regions that may not have expression modifiers
    if (hasGaze) {
      const gazeRegions: Array<{ region: FaceRegion; factor: number; factorY: number }> = [
        { region: 'left_pupil', factor: 0.016, factorY: 0.012 },
        { region: 'right_pupil', factor: 0.016, factorY: 0.012 },
        { region: 'left_eye', factor: 0.010, factorY: 0.008 },
        { region: 'right_eye', factor: 0.010, factorY: 0.008 }
      ];
      for (const gr of gazeRegions) {
        if (!this.lerpedModifiers.has(gr.region)) {
          if (idx < result.length) {
            result[idx].region = gr.region;
            result[idx].dx = this.gazeX * gr.factor * canvasWidth;
            result[idx].dy = this.gazeY * gr.factorY * canvasHeight;
            result[idx].scale = 1;
          } else {
            result.push({
              region: gr.region,
              dx: this.gazeX * gr.factor * canvasWidth,
              dy: this.gazeY * gr.factorY * canvasHeight,
              scale: 1
            });
          }
          idx++;
        }
      }
    }

    result.length = idx;
    return result;
  }

  public getCurrentExpression(): Expression {
    return this.current;
  }

  public isTransitioning(): boolean {
    return this.transitionProgress < 1;
  }
}
