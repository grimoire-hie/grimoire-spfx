/**
 * PersonalityEngine
 * 4 personality modes for Grimm, each with visual + behavior config.
 * Colors, glow, movement, background, and system prompt modifiers.
 */

import { IParticleConfig, DEFAULT_PARTICLE_CONFIG } from './ParticleSystem';
import { PERSONALITY_PROMPT_MODIFIERS } from '../../config/promptCatalog';

// ─── Types ─────────────────────────────────────────────────────

export type PersonalityMode = 'normal' | 'funny' | 'harsh' | 'devil';

export interface IPersonalityConfig {
  mode: PersonalityMode;
  label: string;
  /** Short description shown in the carousel card */
  description: string;
  /** Icon color for the selector button */
  iconColor: string;
  /** Particle system config overrides */
  particleOverrides: Partial<IParticleConfig>;
  /** System prompt personality modifier */
  promptModifier: string;
  /** Eye glow effect (for devil mode) */
  eyeGlow?: {
    color: string;
    radius: number;
    intensity: number;
  };
}

// ─── Personality Definitions ───────────────────────────────────

export const PERSONALITIES: Record<PersonalityMode, IPersonalityConfig> = {
  normal: {
    mode: 'normal',
    label: 'Normal',
    description: 'Calm and professional',
    iconColor: '#a8d8ff',
    particleOverrides: {
      trailStyle: 'afterglow',
      trailLength: 6,
      trailOpacity: 0.15
    },
    promptModifier: PERSONALITY_PROMPT_MODIFIERS.normal
  },

  funny: {
    mode: 'funny',
    label: 'Funny',
    description: 'Witty and playful',
    iconColor: '#ffd700',
    particleOverrides: {
      primaryColor: '#ffd700',
      secondaryColor: '#ffe680',
      glowColor: 'rgba(255, 200, 50, 0.6)',
      glowRadius: 10,
      noiseAmplitude: 5,
      chaosFactor: 0.3,
      bgGradientStart: '#1a1000',
      bgGradientEnd: '#2a1800',
      trailStyle: 'sparkle',
      trailLength: 10,
      trailOpacity: 0.4
    },
    promptModifier: PERSONALITY_PROMPT_MODIFIERS.funny
  },

  harsh: {
    mode: 'harsh',
    label: 'Harsh',
    description: 'Blunt and direct',
    iconColor: '#888899',
    particleOverrides: {
      primaryColor: '#8899aa',
      secondaryColor: '#667788',
      glowColor: 'rgba(100, 120, 140, 0.4)',
      glowRadius: 4,
      noiseAmplitude: 1.5,
      chaosFactor: 0,
      springStrength: 0.06,
      bgGradientStart: '#0a0a0e',
      bgGradientEnd: '#151518',
      trailStyle: 'sharp',
      trailLength: 4,
      trailOpacity: 0.2
    },
    promptModifier: PERSONALITY_PROMPT_MODIFIERS.harsh
  },

  devil: {
    mode: 'devil',
    label: 'Devil',
    description: 'Dark and theatrical',
    iconColor: '#ff2020',
    particleOverrides: {
      primaryColor: '#ff3030',
      secondaryColor: '#cc1010',
      glowColor: 'rgba(255, 30, 30, 0.6)',
      glowRadius: 12,
      noiseAmplitude: 4,
      chaosFactor: 0.2,
      springStrength: 0.05,
      bgGradientStart: '#1a0505',
      bgGradientEnd: '#2a0808',
      trailStyle: 'ember',
      trailLength: 8,
      trailOpacity: 0.5
    },
    eyeGlow: {
      color: '#ff0000',
      radius: 15,
      intensity: 1.5
    },
    promptModifier: PERSONALITY_PROMPT_MODIFIERS.devil
  }
};

// ─── Engine ────────────────────────────────────────────────────

export class PersonalityEngine {
  private current: PersonalityMode = 'normal';

  public setMode(mode: PersonalityMode): void {
    this.current = mode;
  }

  public getMode(): PersonalityMode {
    return this.current;
  }

  public getConfig(): IPersonalityConfig {
    return PERSONALITIES[this.current];
  }

  /**
   * Get the full particle config by merging defaults with personality overrides.
   */
  public getParticleConfig(): IParticleConfig {
    const personality = PERSONALITIES[this.current];
    return { ...DEFAULT_PARTICLE_CONFIG, ...personality.particleOverrides };
  }

  /**
   * Get the system prompt personality modifier to append/replace in the prompt.
   */
  public getPromptModifier(): string {
    return PERSONALITIES[this.current].promptModifier;
  }

  /**
   * Get eye glow config (for devil mode rendering).
   */
  public getEyeGlow(): IPersonalityConfig['eyeGlow'] | undefined {
    return PERSONALITIES[this.current].eyeGlow;
  }

  /**
   * Get all available modes for the UI selector.
   */
  public static getModes(): Array<{ mode: PersonalityMode; label: string; color: string }> {
    return (Object.keys(PERSONALITIES) as PersonalityMode[]).map((key) => ({
      mode: key,
      label: PERSONALITIES[key].label,
      color: PERSONALITIES[key].iconColor
    }));
  }
}
