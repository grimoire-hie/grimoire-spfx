/**
 * ParticleSystem
 * Canvas 2D particle physics engine for the Grimm avatar.
 * Manages the avatar particle field with spring dynamics, Perlin noise drift,
 * and glow effects for a ghostly floating-face aesthetic.
 */

import type { IIdleAnimationState } from './IdleAnimationController';

// ─── Perlin Noise (minimal implementation) ─────────────────────

const PERM = new Uint8Array(512);
const GRAD = [
  [1, 1], [-1, 1], [1, -1], [-1, -1],
  [1, 0], [-1, 0], [0, 1], [0, -1]
];

// Initialize permutation table
(function initPerlin(): void {
  const p = new Uint8Array(256);
  for (let i = 0; i < 256; i++) p[i] = i;
  for (let i = 255; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [p[i], p[j]] = [p[j], p[i]];
  }
  for (let i = 0; i < 512; i++) PERM[i] = p[i & 255];
})();

function fade(t: number): number {
  return t * t * t * (t * (t * 6 - 15) + 10);
}

function lerp(a: number, b: number, t: number): number {
  return a + t * (b - a);
}

function dot2(g: number[], x: number, y: number): number {
  return g[0] * x + g[1] * y;
}

export function perlin2(x: number, y: number): number {
  const xi = Math.floor(x) & 255;
  const yi = Math.floor(y) & 255;
  const xf = x - Math.floor(x);
  const yf = y - Math.floor(y);

  const u = fade(xf);
  const v = fade(yf);

  const aa = PERM[PERM[xi] + yi] & 7;
  const ab = PERM[PERM[xi] + yi + 1] & 7;
  const ba = PERM[PERM[xi + 1] + yi] & 7;
  const bb = PERM[PERM[xi + 1] + yi + 1] & 7;

  const x1 = lerp(dot2(GRAD[aa], xf, yf), dot2(GRAD[ba], xf - 1, yf), u);
  const x2 = lerp(dot2(GRAD[ab], xf, yf - 1), dot2(GRAD[bb], xf - 1, yf - 1), u);

  return lerp(x1, x2, v);
}

// ─── Types ─────────────────────────────────────────────────────

export type FaceRegion =
  | 'head_contour'
  | 'left_eye'
  | 'right_eye'
  | 'left_pupil'
  | 'right_pupil'
  | 'nose'
  | 'mouth_upper'
  | 'mouth_lower'
  | 'left_ear'
  | 'right_ear'
  | 'left_eyebrow'
  | 'right_eyebrow'
  | 'chin'
  | 'ambient';

export interface IParticle {
  /** Current position */
  x: number;
  y: number;
  /** Target position (from face template) */
  tx: number;
  ty: number;
  /** Velocity */
  vx: number;
  vy: number;
  /** Base size */
  size: number;
  /** Current opacity (0-1) */
  alpha: number;
  /** Target opacity */
  targetAlpha: number;
  /** Region this particle belongs to */
  region: FaceRegion;
  /** Normalized template coordinate X (0-1) */
  nx: number;
  /** Normalized template coordinate Y (0-1) */
  ny: number;
  /** Pseudo depth scalar (0-1), used for voxel shading/parallax */
  nz: number;
  /** Unique noise offset for drift */
  noiseOffsetX: number;
  noiseOffsetY: number;
  /** Color (hex string) */
  color: string;
  /** Trail position ring buffer */
  trail: Array<{ x: number; y: number }>;
  /** Write index into trail ring buffer */
  trailWrite: number;
  /** Number of valid trail entries (up to trailLength) */
  trailCount: number;
}

/** Trail rendering style */
export type TrailStyle = 'none' | 'afterglow' | 'sparkle' | 'sharp' | 'ember';

export interface IParticleConfig {
  /** Spring constant for formation (0-1) */
  springStrength: number;
  /** Damping factor (0-1) */
  damping: number;
  /** Perlin noise amplitude in pixels */
  noiseAmplitude: number;
  /** Perlin noise speed */
  noiseSpeed: number;
  /** Chaos factor - randomness in movement (0-1) */
  chaosFactor: number;
  /** Particle primitive shape */
  particleShape: 'square' | 'circle' | 'triangle';
  /** Min rendered voxel size in pixels */
  voxelSizeMin: number;
  /** Max rendered voxel size in pixels */
  voxelSizeMax: number;
  /** Strength of depth-based shading modulation (0-1) */
  voxelDepthShading: number;
  /** Sparse edge breakup amount (0-1) */
  voxelEdgeBreakup: number;
  /** Subtle dynamic camera-like sway on X axis */
  swayAmplitudeX: number;
  /** Subtle dynamic camera-like sway on Y axis */
  swayAmplitudeY: number;
  /** Sway animation speed */
  swaySpeed: number;
  /** Glow blur radius */
  glowRadius: number;
  /** Primary particle color */
  primaryColor: string;
  /** Secondary particle color (for variety) */
  secondaryColor: string;
  /** Glow color */
  glowColor: string;
  /** Background gradient start */
  bgGradientStart: string;
  /** Background gradient end */
  bgGradientEnd: string;
  /** Trail style (default: 'none') */
  trailStyle: TrailStyle;
  /** Trail length — number of past positions to keep (0 = no trail) */
  trailLength: number;
  /** Trail base opacity (0-1, fades linearly along trail) */
  trailOpacity: number;
  /** Enable mesh lines between nearby particles */
  meshEnabled: boolean;
  /** Max distance (in pixels) for mesh connections */
  meshMaxDistance: number;
  /** Mesh line base opacity (0-1, fades with distance) */
  meshOpacity: number;
  /** Mesh line width */
  meshLineWidth: number;
  /** Mesh style policy */
  meshStyle: 'subtle' | 'off';
}

export const DEFAULT_PARTICLE_CONFIG: IParticleConfig = {
  springStrength: 0.08,
  damping: 0.92,
  noiseAmplitude: 3,
  noiseSpeed: 0.003,
  chaosFactor: 0,
  particleShape: 'circle',
  voxelSizeMin: 0.78,
  voxelSizeMax: 2.7,
  voxelDepthShading: 0.78,
  voxelEdgeBreakup: 0.018,
  swayAmplitudeX: 4.6,
  swayAmplitudeY: 2.8,
  swaySpeed: 1.08,
  glowRadius: 4,
  primaryColor: '#a8d8ff',
  secondaryColor: '#e0f0ff',
  glowColor: 'rgba(120, 180, 255, 0.6)',
  bgGradientStart: '#0a0a1a',
  bgGradientEnd: '#0d1b2a',
  trailStyle: 'none',
  trailLength: 0,
  trailOpacity: 0,
  meshEnabled: true,
  meshMaxDistance: 48,
  meshOpacity: 0.08,
  meshLineWidth: 0.45,
  meshStyle: 'subtle'
};

export interface IFaceTemplate {
  region: FaceRegion;
  /** Normalized coordinates (0-1), mapped to canvas at render time */
  points: Array<{ x: number; y: number }>;
}

// ─── Thought Particle (ephemeral, for thinking expression) ────

interface IThoughtParticle {
  x: number; y: number;
  vx: number; vy: number;
  size: number;
  alpha: number;
  color: string;
  age: number;       // ms elapsed
  lifetime: number;  // total ms before removal
  noiseOffset: number;
}

export interface IThoughtParticleSnapshot {
  x: number;
  y: number;
  size: number;
  alpha: number;
  color: string;
}

interface IRenderParticle {
  x: number;
  y: number;
  size: number;
  alpha: number;
  color: string;
  region: FaceRegion;
  depthIntensity: number;
}

// ─── Particle System ───────────────────────────────────────────

export class ParticleSystem {
  private particles: IParticle[] = [];
  private config: IParticleConfig;
  private time: number = 0;
  private isFormed: boolean = false;
  private isDissolved: boolean = false;

  /** Expression offsets applied per region (from ExpressionEngine) */
  private regionOffsets: Map<FaceRegion, { dx: number; dy: number; scale: number }> = new Map();

  /** Mouth parameters from lip sync */
  private mouthOpenness: number = 0;
  private mouthWidth: number = 0.5;
  private mouthRound: number = 0;

  /** Performance degradation flags */
  private glowDisabled: boolean = false;
  private trailsDisabled: boolean = false;
  private meshDisabled: boolean = false;

  /** Pre-rendered glow sprites (keyed by "color|radius") */
  private glowSpriteCache: Map<string, HTMLCanvasElement> = new Map();
  /** Parsed RGB cache for fast per-particle depth shading */
  private rgbCache: Map<string, { r: number; g: number; b: number }> = new Map();
  /** Depth-shaded voxel color cache (quantized key) */
  private voxelColorCache: Map<string, string> = new Map();

  /** Idle animation state (from IdleAnimationController) */
  private idleState: IIdleAnimationState | undefined;

  /** Timestamp when entrance animation started (0 = no entrance in progress) */
  private entranceStartTime: number = 0;

  /** Per-particle delay (ms) before spring captures it during entrance */
  private entranceDelays: number[] = [];

  /** Ephemeral thought particles (separate from main particles array) */
  private thoughtParticles: IThoughtParticle[] = [];
  /** Whether thinking expression is active (spawns new thought particles) */
  private thinkingActive: boolean = false;
  /** Runtime cap for ephemeral thought particles (for perf tier degradation). */
  private maxThoughtParticles: number = 300;
  /** Runtime cap for thought particle spawn per frame. */
  private maxThoughtSpawnPerFrame: number = 3;

  /** Cached background gradient (recreated only on canvas resize or config change) */
  private bgGradient: CanvasGradient | undefined;
  private bgGradientHeight: number = 0;
  private bgGradientStart: string = '';
  private bgGradientEnd: string = '';

  constructor(config: Partial<IParticleConfig> = {}) {
    this.config = { ...DEFAULT_PARTICLE_CONFIG, ...config };
  }

  /**
   * Get or create a pre-rendered glow sprite for a given color + radius.
   * Replaces per-particle shadowBlur with a single drawImage() per particle.
   */
  private getGlowSprite(color: string, radius: number): HTMLCanvasElement {
    const quantizedRadius = Math.max(0.5, Math.round(radius * 2) / 2);
    const key = `${this.config.particleShape}|${color}|${quantizedRadius}`;
    let sprite = this.glowSpriteCache.get(key);
    if (sprite) return sprite;

    const blur = this.config.glowRadius;
    const size = Math.ceil((quantizedRadius + blur) * 2 + 2);
    sprite = document.createElement('canvas');
    sprite.width = size;
    sprite.height = size;
    const sCtx = sprite.getContext('2d');
    if (sCtx) {
      const half = size / 2;
      sCtx.shadowColor = this.config.glowColor;
      sCtx.shadowBlur = blur;
      sCtx.shadowOffsetX = 0;
      sCtx.shadowOffsetY = 0;
      sCtx.fillStyle = color;
      sCtx.fillRect(
        half - quantizedRadius,
        half - quantizedRadius,
        quantizedRadius * 2,
        quantizedRadius * 2
      );
    }
    this.glowSpriteCache.set(key, sprite);
    return sprite;
  }

  private static clamp01(value: number): number {
    if (value < 0) return 0;
    if (value > 1) return 1;
    return value;
  }

  private static pseudoDepth(
    region: FaceRegion,
    nx: number,
    ny: number
  ): number {
    const centerBias = 1 - Math.min(1, Math.hypot(nx - 0.5, ny - 0.52) * 1.75);
    let base: number;

    switch (region) {
      case 'left_pupil':
      case 'right_pupil':
        base = 0.88;
        break;
      case 'left_eye':
      case 'right_eye':
        base = 0.76;
        break;
      case 'left_eyebrow':
      case 'right_eyebrow':
        base = 0.72;
        break;
      case 'nose': {
        const ridge = 1 - Math.min(1, Math.abs(nx - 0.5) * 3.6);
        base = 0.74 + ridge * 0.24;
        break;
      }
      case 'mouth_upper':
      case 'mouth_lower': {
        const lipMid = 1 - Math.min(1, Math.abs(nx - 0.5) * 2.4);
        base = 0.58 + lipMid * 0.22;
        break;
      }
      case 'ambient':
        base = 0.30;
        break;
      case 'chin':
      case 'head_contour':
      case 'left_ear':
      case 'right_ear':
        base = 0.40;
        break;
      default:
        base = 0.50;
        break;
    }

    const verticalAtten = Math.max(0, Math.abs(ny - 0.5) - 0.24) * 0.20;
    return ParticleSystem.clamp01(base + centerBias * 0.12 - verticalAtten);
  }

  private static edgeHash(nx: number, ny: number, nz: number): number {
    const value = Math.sin(nx * 127.1 + ny * 311.7 + nz * 78.233) * 43758.5453;
    return value - Math.floor(value);
  }

  private getRGB(hexColor: string): { r: number; g: number; b: number } {
    const cached = this.rgbCache.get(hexColor);
    if (cached) return cached;

    let value = hexColor.trim().toLowerCase();
    if (value.startsWith('#')) value = value.slice(1);

    let r = 168;
    let g = 216;
    let b = 255;

    if (value.length === 3) {
      const rr = parseInt(value[0] + value[0], 16);
      const gg = parseInt(value[1] + value[1], 16);
      const bb = parseInt(value[2] + value[2], 16);
      if (!Number.isNaN(rr) && !Number.isNaN(gg) && !Number.isNaN(bb)) {
        r = rr;
        g = gg;
        b = bb;
      }
    } else if (value.length === 6) {
      const rr = parseInt(value.slice(0, 2), 16);
      const gg = parseInt(value.slice(2, 4), 16);
      const bb = parseInt(value.slice(4, 6), 16);
      if (!Number.isNaN(rr) && !Number.isNaN(gg) && !Number.isNaN(bb)) {
        r = rr;
        g = gg;
        b = bb;
      }
    }

    const rgb = { r, g, b };
    this.rgbCache.set(hexColor, rgb);
    return rgb;
  }

  private getVoxelColor(baseColor: string, depthIntensity: number): string {
    const quantized = Math.max(0, Math.min(20, Math.round(depthIntensity * 20)));
    const key = `${baseColor}|${quantized}|${this.config.voxelDepthShading}`;
    const cached = this.voxelColorCache.get(key);
    if (cached) return cached;

    const rgb = this.getRGB(baseColor);
    const shade = this.config.voxelDepthShading;
    const boost = (quantized / 20) * shade;
    const r = Math.max(0, Math.min(255, Math.round(rgb.r * (0.70 + boost * 0.45) + boost * 26)));
    const g = Math.max(0, Math.min(255, Math.round(rgb.g * (0.72 + boost * 0.40) + boost * 22)));
    const b = Math.max(0, Math.min(255, Math.round(rgb.b * (0.76 + boost * 0.38) + boost * 18)));
    const shaded = `rgb(${r}, ${g}, ${b})`;
    this.voxelColorCache.set(key, shaded);
    return shaded;
  }

  /**
   * Map normalized 0-1 coordinate to canvas pixel using uniform (square)
   * scaling so the face keeps its proportions regardless of aspect ratio.
   */
  private static mapCoord(
    pt: { x: number; y: number },
    canvasWidth: number,
    canvasHeight: number
  ): { tx: number; ty: number } {
    const scale = Math.min(canvasWidth, canvasHeight);
    const offsetX = (canvasWidth - scale) / 2;
    const offsetY = (canvasHeight - scale) / 2;
    return {
      tx: offsetX + pt.x * scale,
      ty: offsetY + pt.y * scale
    };
  }

  /**
   * Initialize particles from face template.
   * When `entrance` is true, particles scatter across the full canvas (and beyond)
   * with random velocities and per-particle delays. Spring force is suppressed
   * until each particle's delay fires, so they drift freely with noise and then
   * get "captured" one-by-one into the face — true particle wind effect.
   */
  public initFromTemplate(
    templates: IFaceTemplate[],
    canvasWidth: number,
    canvasHeight: number,
    entrance: boolean = false
  ): void {
    this.particles = [];
    this.entranceDelays = [];
    this.glowSpriteCache.clear();
    this.isFormed = true;
    this.isDissolved = false;

    for (const tmpl of templates) {
      for (const pt of tmpl.points) {
        const { tx, ty } = ParticleSystem.mapCoord(pt, canvasWidth, canvasHeight);
        const nx = ParticleSystem.clamp01(pt.x);
        const ny = ParticleSystem.clamp01(pt.y);
        const nz = ParticleSystem.pseudoDepth(tmpl.region, nx, ny);
        const isAmbient = tmpl.region === 'ambient';
        const restAlpha = isAmbient ? 0.18 + Math.random() * 0.28 : 0.52 + Math.random() * 0.40;
        const size = isAmbient ? 0.70 + Math.random() * 0.75 : 1.18 + Math.random() * 1.45;

        if (entrance) {
          // Scatter across full canvas + 40% margin in every direction
          const sx = canvasWidth * (-0.4 + Math.random() * 1.8);
          const sy = canvasHeight * (-0.4 + Math.random() * 1.8);
          // Random velocities in any direction (spring will capture them later)
          const angle = Math.random() * Math.PI * 2;
          const speed = 0.5 + Math.random() * 2;

          this.particles.push({
            x: sx,
            y: sy,
            tx,
            ty,
            vx: Math.cos(angle) * speed,
            vy: Math.sin(angle) * speed,
            size,
            alpha: 0,
            targetAlpha: 0, // unlocked per-particle in update()
            region: tmpl.region,
            nx,
            ny,
            nz,
            noiseOffsetX: Math.random() * 1000,
            noiseOffsetY: Math.random() * 1000,
            color: Math.random() > 0.3 ? this.config.primaryColor : this.config.secondaryColor,
            trail: [],
            trailWrite: 0,
            trailCount: 0
          });

          // Per-particle delay: region base + wide individual randomness
          const r = tmpl.region;
          let base = 0;
          if (r === 'ambient' || r === 'head_contour' || r === 'left_ear' || r === 'right_ear' || r === 'chin') {
            base = 0;
          } else if (r === 'left_eyebrow' || r === 'right_eyebrow' || r === 'mouth_upper' || r === 'mouth_lower') {
            base = 400;
          } else if (r === 'left_eye' || r === 'right_eye' || r === 'nose') {
            base = 800;
          } else if (r === 'left_pupil' || r === 'right_pupil') {
            base = 1200;
          }
          // Add wide per-particle jitter so particles within the same region don't arrive together
          this.entranceDelays.push(base + Math.random() * 600);
        } else {
          this.particles.push({
            x: tx,
            y: ty,
            tx,
            ty,
            vx: 0,
            vy: 0,
            size,
            alpha: restAlpha,
            targetAlpha: restAlpha,
            region: tmpl.region,
            nx,
            ny,
            nz,
            noiseOffsetX: Math.random() * 1000,
            noiseOffsetY: Math.random() * 1000,
            color: Math.random() > 0.3 ? this.config.primaryColor : this.config.secondaryColor,
            trail: [],
            trailWrite: 0,
            trailCount: 0
          });
          this.entranceDelays.push(0);
        }
      }
    }

    this.entranceStartTime = entrance ? performance.now() : 0;
  }

  /**
   * Update target positions for existing particles after canvas resize.
   * Does NOT re-create particles — they smoothly spring to new targets.
   */
  public rescaleTargets(
    templates: IFaceTemplate[],
    canvasWidth: number,
    canvasHeight: number
  ): void {
    let idx = 0;
    for (const tmpl of templates) {
      for (const pt of tmpl.points) {
        if (idx >= this.particles.length) return;
        const { tx, ty } = ParticleSystem.mapCoord(pt, canvasWidth, canvasHeight);
        const particle = this.particles[idx];
        const nx = ParticleSystem.clamp01(pt.x);
        const ny = ParticleSystem.clamp01(pt.y);
        particle.tx = tx;
        particle.ty = ty;
        particle.nx = nx;
        particle.ny = ny;
        particle.nz = ParticleSystem.pseudoDepth(tmpl.region, nx, ny);
        idx++;
      }
    }
  }

  /**
   * Initialize particles in ambient mode — random positions, no face template.
   * Particles drift with Perlin noise only (no spring targets).
   */
  public initAmbient(canvasWidth: number, canvasHeight: number, count: number): void {
    this.particles = [];
    this.glowSpriteCache.clear();
    this.isFormed = false;
    this.isDissolved = false;

    for (let i = 0; i < count; i++) {
      const x = Math.random() * canvasWidth;
      const y = Math.random() * canvasHeight;

      this.particles.push({
        x,
        y,
        tx: x,
        ty: y,
        vx: (Math.random() - 0.5) * 0.5,
        vy: (Math.random() - 0.5) * 0.5,
        size: 0.8 + Math.random() * 1.2,
        alpha: 0.15 + Math.random() * 0.35,
        targetAlpha: 0.15 + Math.random() * 0.35,
        region: 'head_contour',
        nx: ParticleSystem.clamp01(canvasWidth > 0 ? x / canvasWidth : 0.5),
        ny: ParticleSystem.clamp01(canvasHeight > 0 ? y / canvasHeight : 0.5),
        nz: 0.18 + Math.random() * 0.24,
        noiseOffsetX: Math.random() * 1000,
        noiseOffsetY: Math.random() * 1000,
        color: Math.random() > 0.3 ? this.config.primaryColor : this.config.secondaryColor,
        trail: [],
        trailWrite: 0,
        trailCount: 0
      });
    }
  }

  /**
   * Form: particles spring toward their face positions.
   */
  public form(): void {
    if (this.entranceStartTime > 0) return; // entrance stagger handles alpha
    this.isFormed = true;
    this.isDissolved = false;
    for (const p of this.particles) {
      p.targetAlpha = 0.5 + Math.random() * 0.5;
    }
  }

  /**
   * Dissolve: particles drift away and fade.
   */
  public dissolve(): void {
    this.isDissolved = true;
    this.isFormed = false;
    for (const p of this.particles) {
      p.targetAlpha = 0;
      // Push particles outward
      const dx = p.x - p.tx;
      const dy = p.y - p.ty;
      const len = Math.sqrt(dx * dx + dy * dy) || 1;
      p.vx += (dx / len) * 3;
      p.vy += (dy / len) * 3;
    }
  }

  /**
   * Update particle configuration (for personality switching).
   */
  public setConfig(config: Partial<IParticleConfig>): void {
    const oldPrimary = this.config.primaryColor;
    const oldSecondary = this.config.secondaryColor;
    this.config = { ...this.config, ...config };

    // Invalidate glow sprite cache when colors or glow radius change
    if (
      config.primaryColor ||
      config.secondaryColor ||
      config.glowColor ||
      config.glowRadius !== undefined ||
      config.particleShape
    ) {
      this.glowSpriteCache.clear();
    }

    if (
      config.primaryColor ||
      config.secondaryColor ||
      config.voxelDepthShading !== undefined
    ) {
      this.voxelColorCache.clear();
    }

    // Update particle colors if changed
    if (config.primaryColor || config.secondaryColor) {
      for (const p of this.particles) {
        if (p.color === oldPrimary) {
          p.color = this.config.primaryColor;
        } else if (p.color === oldSecondary) {
          p.color = this.config.secondaryColor;
        }
      }
    }
  }

  /**
   * Set expression region offsets (from ExpressionEngine).
   */
  public setRegionOffset(
    region: FaceRegion,
    dx: number,
    dy: number,
    scale: number
  ): void {
    this.regionOffsets.set(region, { dx, dy, scale });
  }

  /**
   * Clear all region offsets.
   */
  public clearRegionOffsets(): void {
    this.regionOffsets.clear();
  }

  /**
   * Set mouth parameters from lip sync.
   */
  public setMouthParams(openness: number, width: number, round: number = 0): void {
    this.mouthOpenness = openness;
    this.mouthWidth = width;
    this.mouthRound = round;
  }

  /**
   * Set idle animation state (from IdleAnimationController).
   * Applied each frame in update().
   */
  public setIdleState(state: IIdleAnimationState | undefined): void {
    this.idleState = state;
  }

  /**
   * Enable/disable thinking particle spawning.
   * When deactivated, existing particles finish their lifecycle naturally.
   */
  public setThinkingActive(active: boolean): void {
    this.thinkingActive = active;
  }

  /**
   * Physics tick. Call once per frame.
   */
  public update(dt: number): void {
    this.time += dt;
    const { springStrength, damping, noiseAmplitude, noiseSpeed, chaosFactor } = this.config;

    // ─── Entrance: per-particle delayed capture ───
    // During entrance, each particle has an individual delay. Before its delay
    // fires it drifts freely (no spring). Once the delay elapses, targetAlpha
    // is unlocked and spring kicks in — the particle gets "captured" into the face.
    let entranceActive = false;
    let entranceElapsed = 0;
    if (this.entranceStartTime > 0) {
      entranceActive = true;
      entranceElapsed = performance.now() - this.entranceStartTime;
      // Unlock targetAlpha for particles whose delay has elapsed
      for (let i = 0; i < this.particles.length; i++) {
        const p = this.particles[i];
        if (p.targetAlpha > 0) continue; // already unlocked
        if (entranceElapsed >= this.entranceDelays[i]) {
          const isAmbient = p.region === 'ambient';
          p.targetAlpha = isAmbient ? 0.18 + Math.random() * 0.28 : 0.52 + Math.random() * 0.40;
        }
      }
      // Max possible delay is ~1800ms (base 1200 + jitter 600) — stop at 2200ms
      if (entranceElapsed >= 2200) {
        this.entranceStartTime = 0;
        this.entranceDelays = [];
      }
    }

    // Idle animation overrides
    const effSpring = this.idleState ? springStrength * this.idleState.springMultiplier : springStrength;
    const effNoise = this.idleState ? noiseAmplitude * this.idleState.driftMultiplier : noiseAmplitude;
    const breathAmp = this.idleState ? this.idleState.breathingAmplitude : 0;
    const wind = this.idleState ? this.idleState.wind : undefined;
    const sparkleCount = this.idleState ? this.idleState.sparkleCount : 0;

    // Sparkle probability per particle (avoids per-frame Set allocation)
    const sparkleProb = sparkleCount > 0 && this.particles.length > 0
      ? sparkleCount / this.particles.length
      : 0;
    const t = this.time * noiseSpeed;

    for (let pi = 0; pi < this.particles.length; pi++) {
      const p = this.particles[pi];
      // During entrance, skip spring for particles not yet captured (still drifting)
      const captured = !entranceActive || p.targetAlpha > 0;
      if (captured && this.isFormed && !this.isDissolved) {
        // Get effective target with expression offsets
        let etx = p.tx;
        let ety = p.ty;

        const offset = this.regionOffsets.get(p.region);
        if (offset) {
          etx += offset.dx;
          ety += offset.dy;
          // Scale is applied relative to region center (simplified: just offset)
        }

        // Breathing: gentle sine wave on Y
        if (breathAmp > 0) {
          ety += Math.sin(this.time * Math.PI * 2 / 4) * breathAmp;
        }

        // Mouth deformation for lip sync (open, width, round)
        if (p.region === 'mouth_lower') {
          ety += this.mouthOpenness * 12; // Open mouth downward
        }
        if (p.region === 'mouth_upper' || p.region === 'mouth_lower') {
          const centerX = p.tx;
          const dx = p.x - centerX;
          // Width: spreads or narrows mouth horizontally
          etx += dx * (this.mouthWidth - 0.5) * 0.3;
          // Round: pulls outer particles inward (pursing lips)
          etx -= dx * this.mouthRound * 0.25;
        }

        // Spring toward target — ambient particles use weaker spring (0.3x)
        const springMul = p.region === 'ambient' ? 0.3 : 1;
        const dx = etx - p.x;
        const dy = ety - p.y;
        p.vx += dx * effSpring * springMul;
        p.vy += dy * effSpring * springMul;
      }

      // Perlin noise drift — ambient particles get 3x stronger noise
      const noiseMul = p.region === 'ambient' ? 3 : 1;
      const nx = perlin2(p.noiseOffsetX + t, p.noiseOffsetY) * effNoise * noiseMul;
      const ny = perlin2(p.noiseOffsetX, p.noiseOffsetY + t) * effNoise * noiseMul;
      p.vx += nx * 0.1;
      p.vy += ny * 0.1;

      // Wind force (idle tier 4) — ambient particles get 2x wind effect
      if (wind) {
        const windMul = p.region === 'ambient' ? 2 : 1;
        p.vx += wind.x * wind.strength * windMul;
        p.vy += wind.y * wind.strength * windMul;
      }

      // Sparkle: temporarily boost alpha for randomly selected particles
      if (sparkleProb > 0 && Math.random() < sparkleProb) {
        p.alpha = Math.min(1, p.alpha + 0.3 * Math.abs(Math.sin(this.time * 8 + p.noiseOffsetX)));
      }

      // Chaos
      if (chaosFactor > 0) {
        p.vx += (Math.random() - 0.5) * chaosFactor * 2;
        p.vy += (Math.random() - 0.5) * chaosFactor * 2;
      }

      // Damping
      p.vx *= damping;
      p.vy *= damping;

      // Integrate
      p.x += p.vx;
      p.y += p.vy;

      // Record trail history (ring buffer — avoids O(n) unshift)
      if (this.config.trailLength > 0) {
        const tLen = this.config.trailLength;
        // Pre-allocate trail array slots if needed
        if (p.trail.length < tLen) {
          p.trail.push({ x: p.x, y: p.y });
          p.trailWrite = p.trail.length % tLen;
          p.trailCount = p.trail.length;
        } else {
          p.trail[p.trailWrite] = { x: p.x, y: p.y };
          p.trailWrite = (p.trailWrite + 1) % tLen;
          p.trailCount = tLen;
        }
      }

      // Fade toward target alpha (fast: visible within ~2s)
      p.alpha = lerp(p.alpha, p.targetAlpha, 0.12);
    }

    // ─── Thought particles (ephemeral, thinking expression) ───
    const dtMs = dt * 1000;

    // Spawn new thought particles above eyebrows when thinking + formed
    if (this.thinkingActive && this.isFormed && !this.isDissolved) {
      // Find average eyebrow position
      let browSumX = 0;
      let browSumY = 0;
      let browCount = 0;
      for (let pi = 0; pi < this.particles.length; pi++) {
        const bp = this.particles[pi];
        if (bp.region === 'left_eyebrow' || bp.region === 'right_eyebrow') {
          browSumX += bp.x;
          browSumY += bp.y;
          browCount++;
        }
      }

      if (browCount > 0) {
        const browCx = browSumX / browCount;
        const browCy = browSumY / browCount;

        // Spawn particles with runtime caps (balanced perf floor under degradation).
        const availableSlots = Math.max(0, this.maxThoughtParticles - this.thoughtParticles.length);
        const intendedSpawn = 2 + (Math.random() < 0.5 ? 1 : 0);
        const spawnCount = Math.min(availableSlots, this.maxThoughtSpawnPerFrame, intendedSpawn);
        for (let s = 0; s < spawnCount; s++) {
          this.thoughtParticles.push({
            x: browCx + (Math.random() - 0.5) * 80,
            y: browCy - 30 - Math.random() * 30,
            vx: (Math.random() - 0.5) * 0.4,
            vy: -(0.3 + Math.random() * 0.5),
            size: 0.8 + Math.random() * 0.7,
            alpha: 0,
            color: this.config.primaryColor,
            age: 0,
            lifetime: 1200 + Math.random() * 800,
            noiseOffset: Math.random() * 1000
          });
        }
      }
    }

    // Update existing thought particles
    let writeIdx = 0;
    for (let ti = 0; ti < this.thoughtParticles.length; ti++) {
      const tp = this.thoughtParticles[ti];
      tp.age += dtMs;

      // Remove expired
      if (tp.age >= tp.lifetime) continue;

      // Perlin noise drift (gentle)
      const tn = this.time * noiseSpeed;
      tp.vx += perlin2(tp.noiseOffset + tn, tp.noiseOffset) * 0.05;
      tp.vy += perlin2(tp.noiseOffset, tp.noiseOffset + tn) * 0.03;

      // Light damping (keeps them drifting)
      tp.vx *= 0.98;
      tp.vy *= 0.98;

      // Integrate
      tp.x += tp.vx;
      tp.y += tp.vy;

      // Alpha curve: fade in (0-200ms), hold, fade out (last 400ms)
      const fadeInEnd = 200;
      const fadeOutStart = tp.lifetime - 400;
      if (tp.age < fadeInEnd) {
        tp.alpha = tp.age / fadeInEnd * 0.6;
      } else if (tp.age > fadeOutStart) {
        tp.alpha = (1 - (tp.age - fadeOutStart) / 400) * 0.6;
      } else {
        tp.alpha = 0.6;
      }

      // Compact in-place (avoid array allocation)
      this.thoughtParticles[writeIdx++] = tp;
    }
    this.thoughtParticles.length = writeIdx;
  }

  /**
   * Render all particles to canvas.
   */
  public render(ctx: CanvasRenderingContext2D, width: number, height: number): void {
    // Background gradient (cached — recreated only on height or color change)
    if (
      !this.bgGradient ||
      this.bgGradientHeight !== height ||
      this.bgGradientStart !== this.config.bgGradientStart ||
      this.bgGradientEnd !== this.config.bgGradientEnd
    ) {
      this.bgGradient = ctx.createLinearGradient(0, 0, 0, height);
      this.bgGradient.addColorStop(0, this.config.bgGradientStart);
      this.bgGradient.addColorStop(1, this.config.bgGradientEnd);
      this.bgGradientHeight = height;
      this.bgGradientStart = this.config.bgGradientStart;
      this.bgGradientEnd = this.config.bgGradientEnd;
    }
    ctx.fillStyle = this.bgGradient;
    ctx.fillRect(0, 0, width, height);

    const { trailStyle, trailLength, trailOpacity } = this.config;
    const swayX = Math.sin(this.time * this.config.swaySpeed) * this.config.swayAmplitudeX;
    const swayY = Math.sin(this.time * this.config.swaySpeed * 0.7) * this.config.swayAmplitudeY;

    // Draw trails (behind particles) — skipped when degraded
    if (!this.trailsDisabled && trailStyle !== 'none' && trailLength > 0 && trailOpacity > 0) {
      this.renderTrails(ctx, trailStyle, trailLength, trailOpacity);
    }

    const renderParticles: IRenderParticle[] = [];

    for (const p of this.particles) {
      if (p.alpha < 0.01) continue;

      const px = p.x + swayX * p.nz;
      const py = p.y + swayY * p.nz * 0.6;
      const centerDist = Math.min(1, Math.hypot(p.nx - 0.5, p.ny - 0.52) * 1.58);
      const edgeWeight = Math.max(0, centerDist - 0.54) / 0.46;
      const breakup = this.config.voxelEdgeBreakup;
      const breakupEligible =
        p.region === 'ambient' ||
        p.region === 'head_contour' ||
        p.region === 'chin' ||
        p.region === 'left_ear' ||
        p.region === 'right_ear';
      if (breakupEligible && breakup > 0 && edgeWeight > 0) {
        const hash = ParticleSystem.edgeHash(p.nx, p.ny, p.nz);
        if (hash < breakup * edgeWeight * 0.58) continue;
      }

      const centerBias = 1 - centerDist;
      const depthIntensity = ParticleSystem.clamp01(p.nz * 0.74 + centerBias * 0.26);
      const regionWeight =
        p.region === 'ambient' ? 0.76 :
          (p.region === 'left_pupil' || p.region === 'right_pupil' ? 1.10 :
            (p.region === 'left_eye' || p.region === 'right_eye' ? 1.02 :
              (p.region === 'nose' ? 1.00 : 0.96)));
      const voxelSize = lerp(this.config.voxelSizeMin, this.config.voxelSizeMax, depthIntensity) * regionWeight;
      const alphaMul = p.region === 'ambient' ? 1.00 : 1.06;
      const alphaLift = p.region === 'ambient' ? 0 : 0.08;
      const alpha = Math.min(1, p.alpha * alphaMul * (0.72 + depthIntensity * 0.30) + alphaLift);
      const shadedColor = this.getVoxelColor(p.color, depthIntensity);

      renderParticles.push({
        x: px,
        y: py,
        size: voxelSize,
        alpha,
        color: shadedColor,
        region: p.region,
        depthIntensity
      });
    }

    // Draw mesh lines between nearby particles (behind particles, creates wireframe)
    if (this.config.meshEnabled && !this.meshDisabled && this.config.meshStyle !== 'off') {
      this.renderMesh(ctx, renderParticles);
    }

    // Draw main voxel particles
    for (const p of renderParticles) {
      if (p.alpha < 0.01) continue;
      const half = p.size * 0.5;
      ctx.globalAlpha = p.alpha;
      ctx.fillStyle = p.color;
      ctx.fillRect(p.x - half, p.y - half, p.size, p.size);

      if (p.region !== 'ambient' && p.alpha > 0.12) {
        const coreSize = Math.max(this.config.voxelSizeMin * 0.62, p.size * 0.52);
        const coreHalf = coreSize * 0.5;
        ctx.globalAlpha = Math.min(1, p.alpha * 0.56);
        ctx.fillStyle = this.getVoxelColor(this.config.secondaryColor, Math.min(1, p.depthIntensity + 0.16));
        ctx.fillRect(p.x - coreHalf, p.y - coreHalf, coreSize, coreSize);
      } else if (p.region === 'ambient' && p.alpha > 0.10) {
        const speck = Math.max(0.30, p.size * 0.28);
        const speckHalf = speck * 0.5;
        ctx.globalAlpha = Math.min(1, p.alpha * 0.36);
        ctx.fillStyle = this.getVoxelColor(this.config.secondaryColor, Math.min(1, p.depthIntensity + 0.10));
        ctx.fillRect(p.x + half * 0.36 - speckHalf, p.y - half * 0.36 - speckHalf, speck, speck);
      }

      if (!this.glowDisabled) {
        const sprite = this.getGlowSprite(p.color, p.size);
        const spriteHalf = sprite.width / 2;
        const glowMul = p.region === 'ambient' ? 0.24 : 0.40;
        ctx.globalAlpha = p.alpha * glowMul;
        ctx.drawImage(sprite, p.x - spriteHalf, p.y - spriteHalf);
      }
    }

    // Draw thought particles (same voxel style, above face)
    for (const tp of this.thoughtParticles) {
      if (tp.alpha < 0.01) continue;
      const forward = 0.48;
      const tx = tp.x + swayX * forward;
      const ty = tp.y + swayY * forward * 0.6;
      const size = Math.max(this.config.voxelSizeMin * 0.75, tp.size * 1.6);
      const half = size * 0.5;
      const color = this.getVoxelColor(tp.color, 0.66);
      ctx.globalAlpha = tp.alpha * 0.90;
      ctx.fillStyle = color;
      ctx.fillRect(tx - half, ty - half, size, size);

      if (!this.glowDisabled) {
        const sprite = this.getGlowSprite(color, size);
        const spriteHalf = sprite.width / 2;
        ctx.globalAlpha = tp.alpha * 0.35;
        ctx.drawImage(sprite, tx - spriteHalf, ty - spriteHalf);
      }
    }

    // Reset
    ctx.globalAlpha = 1;
  }

  /**
   * Render proximity mesh lines between nearby particles.
   * Creates a wireframe/connected-graph aesthetic. O(n²) but with
   * ~200 particles that's ~20k checks — trivially fast per frame.
   */
  private renderMesh(ctx: CanvasRenderingContext2D, particles: IRenderParticle[]): void {
    if (this.config.meshStyle === 'off') return;

    const maxDistance = this.config.meshMaxDistance * 0.66;
    const maxDist2 = maxDistance * maxDistance;
    const meshOpacity = this.config.meshOpacity * 0.34;
    const meshLineWidth = Math.max(0.35, this.config.meshLineWidth * 0.85);

    const maxMeshParticles = 220;
    const coreParticles: IRenderParticle[] = [];
    const ambientParticles: IRenderParticle[] = [];
    for (const p of particles) {
      if (p.region === 'ambient') {
        ambientParticles.push(p);
      } else {
        coreParticles.push(p);
      }
    }
    const meshParticles: IRenderParticle[] = [...coreParticles];
    const ambientSlots = Math.max(0, maxMeshParticles - meshParticles.length);
    if (ambientSlots > 0 && ambientParticles.length > 0) {
      const step = Math.max(1, Math.floor(ambientParticles.length / ambientSlots));
      for (let i = 0; i < ambientParticles.length && meshParticles.length < maxMeshParticles; i += step) {
        meshParticles.push(ambientParticles[i]);
      }
    }
    const count = meshParticles.length;

    ctx.lineWidth = meshLineWidth;
    ctx.strokeStyle = this.config.primaryColor;

    // Batch lines into alpha buckets (10 levels) to minimize state changes
    const BUCKET_COUNT = 10;
    const buckets: Array<Array<number>> = [];
    for (let b = 0; b < BUCKET_COUNT; b++) buckets.push([]);

    for (let i = 0; i < count; i++) {
      const a = meshParticles[i];
      if (a.alpha < 0.05) continue;

      for (let j = i + 1; j < count; j++) {
        const b = meshParticles[j];
        if (b.alpha < 0.05) continue;
        if (a.region === 'ambient' && b.region === 'ambient') continue;

        const dx = a.x - b.x;
        const dy = a.y - b.y;
        const dist2 = dx * dx + dy * dy;
        const pairMaxDist2 =
          (a.region === 'ambient' || b.region === 'ambient')
            ? maxDist2 * 0.54
            : maxDist2;

        if (dist2 < pairMaxDist2) {
          const dist = Math.sqrt(dist2);
          const pairMaxDistance = Math.sqrt(pairMaxDist2);
          const fade = 1 - dist / pairMaxDistance;
          const alpha = fade * meshOpacity * Math.min(a.alpha, b.alpha);
          const bucketIdx = Math.min(BUCKET_COUNT - 1, Math.floor(alpha * BUCKET_COUNT));
          buckets[bucketIdx].push(a.x, a.y, b.x, b.y);
        }
      }
    }

    // Draw each bucket as a single path
    for (let b = 0; b < BUCKET_COUNT; b++) {
      const coords = buckets[b];
      if (coords.length === 0) continue;
      ctx.globalAlpha = (b + 0.5) / BUCKET_COUNT;
      ctx.beginPath();
      for (let k = 0; k < coords.length; k += 4) {
        ctx.moveTo(coords[k], coords[k + 1]);
        ctx.lineTo(coords[k + 2], coords[k + 3]);
      }
      ctx.stroke();
    }

    ctx.globalAlpha = 1;
  }

  /**
   * Get trail point at age index (0 = newest) from ring buffer.
   */
  private static trailAt(p: IParticle, age: number): { x: number; y: number } {
    // trailWrite points to the next write slot, so newest is at (trailWrite - 1)
    const idx = (p.trailWrite - 1 - age + p.trailCount * 2) % p.trailCount;
    return p.trail[idx];
  }

  /**
   * Render particle trails per style.
   */
  private renderTrails(
    ctx: CanvasRenderingContext2D,
    style: TrailStyle,
    maxLen: number,
    baseOpacity: number
  ): void {
    ctx.shadowBlur = 0;

    for (const p of this.particles) {
      if (p.alpha < 0.01 || p.trailCount < 2) continue;

      const len = Math.min(p.trailCount, maxLen);

      switch (style) {
        case 'afterglow': {
          // Soft glowing dots that fade out — ghostly residue (sprite-based)
          const agSprite = this.getGlowSprite(p.color, p.size);
          const agHalf = agSprite.width / 2;
          for (let i = 1; i < len; i++) {
            const t = i / len;
            const pt = ParticleSystem.trailAt(p, i);
            ctx.globalAlpha = p.alpha * baseOpacity * (1 - t);
            ctx.drawImage(agSprite, pt.x - agHalf, pt.y - agHalf);
          }
          break;
        }

        case 'sparkle':
          // Scattered tiny sparkles along the trail path — playful
          for (let i = 1; i < len; i += 2) {
            const t = i / len;
            const pt = ParticleSystem.trailAt(p, i);
            // Offset sparkle position slightly for scatter effect
            const jitterX = (Math.sin(p.noiseOffsetX + i * 7.3) * 3);
            const jitterY = (Math.cos(p.noiseOffsetY + i * 5.1) * 3);
            ctx.globalAlpha = p.alpha * baseOpacity * (1 - t) * (0.5 + Math.sin(this.time * 8 + i) * 0.5);
            ctx.fillStyle = p.color;
            const size = Math.max(0.6, p.size * 1.1 * (1 - t));
            const half = size * 0.5;
            ctx.fillRect(pt.x + jitterX - half, pt.y + jitterY - half, size, size);
          }
          break;

        case 'sharp':
          // Thin connected line segments — precise, mechanical
          ctx.strokeStyle = p.color;
          ctx.lineWidth = 0.8;
          ctx.beginPath();
          ctx.moveTo(p.x, p.y);
          for (let i = 0; i < len; i++) {
            const t = i / len;
            const pt = ParticleSystem.trailAt(p, i);
            ctx.globalAlpha = p.alpha * baseOpacity * (1 - t);
            ctx.lineTo(pt.x, pt.y);
          }
          ctx.stroke();
          break;

        case 'ember': {
          // Glowing dots that shrink and redden — fiery decay (sprite-based)
          const emSprite = this.getGlowSprite(p.color, p.size);
          const emHalf = emSprite.width / 2;
          for (let i = 1; i < len; i++) {
            const t = i / len;
            const pt = ParticleSystem.trailAt(p, i);
            const flicker = 0.8 + Math.sin(this.time * 12 + p.noiseOffsetX + i * 3) * 0.2;
            ctx.globalAlpha = p.alpha * baseOpacity * (1 - t * t) * flicker;
            ctx.drawImage(emSprite, pt.x - emHalf, pt.y - emHalf);
          }
          break;
        }

        default:
          break;
      }
    }

    ctx.globalAlpha = 1;
    ctx.shadowBlur = 0;
  }

  /**
   * Get current particle count (for adaptive performance).
   */
  public getParticleCount(): number {
    return this.particles.length;
  }

  /**
   * Reduce particle count for performance (randomly samples to target count).
   */
  public reduceParticles(keepRatio: number): void {
    if (keepRatio >= 1 || this.particles.length === 0) return;
    const targetCount = Math.max(1, Math.floor(this.particles.length * keepRatio));
    if (targetCount >= this.particles.length) return;
    // Fisher-Yates partial shuffle: randomly select targetCount particles into the tail
    const arr = this.particles;
    for (let i = arr.length - 1; i > arr.length - 1 - targetCount; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
    }
    this.particles = arr.slice(arr.length - targetCount);
  }

  /**
   * Check if particles have converged (formed).
   */
  public isConverged(threshold: number = 5): boolean {
    if (this.particles.length === 0) return false;
    let totalDist = 0;
    for (const p of this.particles) {
      const dx = p.x - p.tx;
      const dy = p.y - p.ty;
      totalDist += Math.sqrt(dx * dx + dy * dy);
    }
    return (totalDist / this.particles.length) < threshold;
  }

  public getConfig(): IParticleConfig {
    return { ...this.config };
  }

  public getParticles(): readonly IParticle[] {
    return this.particles;
  }

  public getThoughtParticles(): IThoughtParticleSnapshot[] {
    return this.thoughtParticles.map((p) => ({
      x: p.x,
      y: p.y,
      size: p.size,
      alpha: p.alpha,
      color: p.color
    }));
  }

  /**
   * Disable glow (shadowBlur) for performance recovery.
   */
  public disableGlow(): void {
    this.glowDisabled = true;
  }

  /**
   * Re-enable glow.
   */
  public enableGlow(): void {
    this.glowDisabled = false;
  }

  public isGlowDisabled(): boolean {
    return this.glowDisabled;
  }

  /**
   * Disable trail rendering for performance recovery.
   */
  public disableTrails(): void {
    this.trailsDisabled = true;
  }

  /**
   * Re-enable trail rendering.
   */
  public enableTrails(): void {
    this.trailsDisabled = false;
  }

  public isTrailsDisabled(): boolean {
    return this.trailsDisabled;
  }

  /**
   * Disable mesh rendering for performance recovery.
   */
  public disableMesh(): void {
    this.meshDisabled = true;
  }

  /**
   * Re-enable mesh rendering.
   */
  public enableMesh(): void {
    this.meshDisabled = false;
  }

  public isMeshDisabled(): boolean {
    return this.meshDisabled;
  }

  /**
   * Set thought particle limits used by tiered degradation policy.
   */
  public setThoughtParticleLimits(maxParticles: number, maxSpawnPerFrame: number): void {
    this.maxThoughtParticles = Math.max(0, Math.floor(maxParticles));
    this.maxThoughtSpawnPerFrame = Math.max(0, Math.floor(maxSpawnPerFrame));
    if (this.thoughtParticles.length > this.maxThoughtParticles) {
      this.thoughtParticles = this.thoughtParticles.slice(-this.maxThoughtParticles);
    }
  }
}

// ─── Performance Monitor ──────────────────────────────────────

/** Degradation tier — lower is better quality */
export type PerfTier = 0 | 1 | 2 | 3;

/**
 * PerformanceMonitor
 * Tracks frame times and recommends degradation tiers.
 *
 * Tier 0: Full quality (glow + trails + mesh + full thought particles)
 * Tier 1: Disable glow (biggest perf win, least visual loss)
 * Tier 2: Disable trails
 * Tier 3: Disable mesh + cap thought particles
 *
 * Recovers one tier at a time when performance improves.
 */
export class PerformanceMonitor {
  private readonly windowSize: number = 90;
  /** Circular buffer for frame times */
  private frameTimes: Float64Array = new Float64Array(this.windowSize);
  private writeIndex: number = 0;
  private frameCount: number = 0;
  private runningSum: number = 0;

  private currentTier: PerfTier = 0;
  private cooldownFrames: number = 0;
  private readonly cooldownLength: number = 120;
  private degradeConsecutive: number = 0;
  private recoverConsecutive: number = 0;
  private readonly requiredConsecutiveFrames: number = 24;
  private readonly degradeThresholds: ReadonlyArray<number> = [18.8, 21.8, 24.8];
  private readonly recoverThresholds: ReadonlyArray<number> = [Number.NEGATIVE_INFINITY, 15.4, 13.8, 12.6];

  /**
   * Record a frame time and return the recommended tier.
   */
  public recordFrame(frameTimeMs: number): PerfTier {
    const normalizedFrameMs = Math.max(4, Math.min(100, frameTimeMs));

    // Subtract the value being overwritten from running sum
    if (this.frameCount >= this.windowSize) {
      this.runningSum -= this.frameTimes[this.writeIndex];
    }
    this.frameTimes[this.writeIndex] = normalizedFrameMs;
    this.runningSum += normalizedFrameMs;
    this.writeIndex = (this.writeIndex + 1) % this.windowSize;
    if (this.frameCount < this.windowSize) this.frameCount++;

    if (this.cooldownFrames > 0) {
      this.cooldownFrames--;
      return this.currentTier;
    }

    // Need a full window before making decisions
    if (this.frameCount < this.windowSize) {
      return this.currentTier;
    }

    const avg = this.runningSum / this.frameCount;
    const degradeThreshold = this.degradeThresholds[this.currentTier] ?? Number.POSITIVE_INFINITY;
    const recoverThreshold = this.recoverThresholds[this.currentTier] ?? Number.NEGATIVE_INFINITY;

    if (avg > degradeThreshold && this.currentTier < 3) {
      this.degradeConsecutive++;
      this.recoverConsecutive = 0;
    } else if (avg < recoverThreshold && this.currentTier > 0) {
      this.recoverConsecutive++;
      this.degradeConsecutive = 0;
    } else {
      this.degradeConsecutive = 0;
      this.recoverConsecutive = 0;
    }

    if (this.degradeConsecutive >= this.requiredConsecutiveFrames && this.currentTier < 3) {
      this.currentTier = (this.currentTier + 1) as PerfTier;
      this.cooldownFrames = this.cooldownLength;
      this.degradeConsecutive = 0;
      this.recoverConsecutive = 0;
    } else if (this.recoverConsecutive >= this.requiredConsecutiveFrames && this.currentTier > 0) {
      this.currentTier = (this.currentTier - 1) as PerfTier;
      this.cooldownFrames = this.cooldownLength;
      this.degradeConsecutive = 0;
      this.recoverConsecutive = 0;
    }

    return this.currentTier;
  }

  public getTier(): PerfTier {
    return this.currentTier;
  }

  public getAvgFrameTime(): number {
    if (this.frameCount === 0) return 0;
    return this.runningSum / this.frameCount;
  }

  public getSampleCount(): number {
    return this.frameCount;
  }

  public reset(): void {
    this.frameCount = 0;
    this.writeIndex = 0;
    this.runningSum = 0;
    this.currentTier = 0;
    this.cooldownFrames = 0;
    this.degradeConsecutive = 0;
    this.recoverConsecutive = 0;
  }
}
