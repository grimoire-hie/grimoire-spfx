/**
 * IdleAnimationController — Multi-tier inactivity animations for the particle avatar.
 *
 * Tiers (by idle time):
 *   0: Normal (0-30s)
 *   1: Breathing (30-60s)    — gentle sine wave on Y offset
 *   2: Sparkle (1-2min)      — random particles brighten
 *   3: Slow drift (2-5min)   — increased noise, looser springs
 *   4: Wind dissolution (5+min) — directional force scatters particles
 *
 * Any user activity triggers reformation (tier 0 with reforming flag).
 */

export type InactivityTier = 0 | 1 | 2 | 3 | 4;

/** Animation parameters for a single frame, consumed by ParticleSystem */
export interface IIdleAnimationState {
  tier: InactivityTier;
  /** Breathing: sine amplitude for Y offset (0 = none) */
  breathingAmplitude: number;
  /** Sparkle: number of particles to brighten this frame (0 = none) */
  sparkleCount: number;
  /** Drift: multiplier for noiseAmplitude (1.0 = normal) */
  driftMultiplier: number;
  /** Drift: multiplier for springStrength (1.0 = normal) */
  springMultiplier: number;
  /** Wind: directional force vector (undefined = no wind) */
  wind: { x: number; y: number; strength: number } | undefined;
  /** Whether to trigger reformation animation */
  reforming: boolean;
}

/** Tier thresholds in milliseconds */
const TIER_THRESHOLDS: number[] = [
  0,      // Tier 0: Normal (0-30s)
  30000,  // Tier 1: Breathing (30-60s)
  60000,  // Tier 2: Sparkle (1-2min)
  120000, // Tier 3: Slow drift (2-5min)
  300000  // Tier 4: Wind dissolution (5+min)
];

export class IdleAnimationController {
  private lastActivityTime: number = Date.now();
  private currentTier: InactivityTier = 0;
  private wasHighTier: boolean = false;
  private ambientCap: InactivityTier | 5 = 5; // 5 = no cap
  private windAngle: number = 0;

  /** Reusable state object (avoids per-frame allocation) */
  private readonly state: IIdleAnimationState = {
    tier: 0, breathingAmplitude: 0, sparkleCount: 0,
    driftMultiplier: 1.0, springMultiplier: 1.0, wind: undefined, reforming: false
  };
  /** Reusable wind object (avoids per-frame allocation for tier 4) */
  private readonly windState: { x: number; y: number; strength: number } = { x: 0, y: 0, strength: 0 };

  /** Call on any user activity (mouse move, voice, typing, click) */
  public recordActivity(): void {
    this.lastActivityTime = Date.now();
    if (this.currentTier > 0) {
      this.wasHighTier = true;
    }
  }

  /** Override maximum tier from ambient sound state */
  public setAmbientCap(cap: InactivityTier | 5): void {
    this.ambientCap = cap;
  }

  /** Call per frame. Returns animation parameters for this frame. */
  public update(dt: number): IIdleAnimationState {
    const elapsed = Date.now() - this.lastActivityTime;

    // Determine tier from elapsed time
    let newTier: InactivityTier = 0;
    for (let i = TIER_THRESHOLDS.length - 1; i >= 0; i--) {
      if (elapsed >= TIER_THRESHOLDS[i]) {
        newTier = i as InactivityTier;
        break;
      }
    }

    // Apply ambient cap
    if (this.ambientCap < 5 && newTier > this.ambientCap) {
      newTier = this.ambientCap as InactivityTier;
    }

    // Detect reformation (was at tier > 0, now back to 0)
    const reforming = this.wasHighTier && newTier === 0;
    if (reforming) {
      this.wasHighTier = false;
    }

    this.currentTier = newTier;

    // Slowly rotate wind angle for tier 4 (wrap to prevent precision loss)
    this.windAngle = (this.windAngle + dt * 0.3) % (Math.PI * 2);

    // Mutate persistent state object (avoids per-frame allocation)
    const s = this.state;
    s.tier = newTier;
    s.reforming = reforming;

    switch (newTier) {
      case 0:
        s.breathingAmplitude = 0;
        s.sparkleCount = 0;
        s.driftMultiplier = 1.0;
        s.springMultiplier = 1.0;
        s.wind = undefined;
        break;

      case 1:
        s.breathingAmplitude = 2.0;
        s.sparkleCount = 0;
        s.driftMultiplier = 1.0;
        s.springMultiplier = 1.0;
        s.wind = undefined;
        s.reforming = false;
        break;

      case 2:
        s.breathingAmplitude = 2.5;
        s.sparkleCount = 3;
        s.driftMultiplier = 1.3;
        s.springMultiplier = 0.9;
        s.wind = undefined;
        s.reforming = false;
        break;

      case 3:
        s.breathingAmplitude = 3.0;
        s.sparkleCount = 5;
        s.driftMultiplier = 2.0;
        s.springMultiplier = 0.6;
        s.wind = undefined;
        s.reforming = false;
        break;

      case 4:
        this.windState.x = Math.cos(this.windAngle) * 0.5;
        this.windState.y = Math.sin(this.windAngle * 0.7) * 0.3 - 0.2;
        this.windState.strength = 0.8;
        s.breathingAmplitude = 0;
        s.sparkleCount = 8;
        s.driftMultiplier = 3.0;
        s.springMultiplier = 0.2;
        s.wind = this.windState;
        s.reforming = false;
        break;

      default:
        s.breathingAmplitude = 0;
        s.sparkleCount = 0;
        s.driftMultiplier = 1.0;
        s.springMultiplier = 1.0;
        s.wind = undefined;
        break;
    }

    return s;
  }

  public getCurrentTier(): InactivityTier {
    return this.currentTier;
  }

  public reset(): void {
    this.lastActivityTime = Date.now();
    this.currentTier = 0;
    this.wasHighTier = false;
    this.ambientCap = 5;
  }
}
