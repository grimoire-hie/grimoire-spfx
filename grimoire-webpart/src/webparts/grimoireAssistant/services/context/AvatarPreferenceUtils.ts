import { VISAGE_OPTIONS, VisageMode } from '../avatar/FaceTemplateData';
import { PERSONALITIES, PersonalityMode } from '../avatar/PersonalityEngine';
import {
  isSupportedRealtimeVoiceId,
  SUPPORTED_REALTIME_VOICE_IDS
} from '../realtime/RealtimeVoiceCatalog';

export const AVATAR_SETTINGS_KEY = 'avatar.settings.v1';

export const VALID_VOICE_IDS = SUPPORTED_REALTIME_VOICE_IDS;
export type VoiceId = typeof VALID_VOICE_IDS[number];
const VALID_PERSONALITY_SET = new Set<string>(Object.keys(PERSONALITIES));
const VALID_VISAGE_SET = new Set<string>(Object.keys(VISAGE_OPTIONS));
const LEGACY_VISAGE_ALIASES: Record<string, VisageMode> = {
  guyFawkes: 'anonyMousse',
  clippy: 'classic',
  friendlyRobot: 'robot',
  pixelAi: 'robot',
  grimoire_test_1: 'classic',
  grimoire_test_2: 'anonyMousse'
};

export interface IAvatarSettingsSnapshot {
  avatarEnabled: boolean;
  voiceId: string;
  personality: PersonalityMode;
  visage: VisageMode;
}

interface IAvatarSettingsPreferencePayload {
  avatarEnabled?: boolean;
  voiceId?: string;
  personality?: string;
  visage?: string;
  updatedAt?: string;
}

export function parseAvatarSettingsPreference(raw: string | undefined): Partial<IAvatarSettingsSnapshot> {
  if (!raw) return {};

  let parsed: IAvatarSettingsPreferencePayload;
  try {
    parsed = JSON.parse(raw) as IAvatarSettingsPreferencePayload;
  } catch {
    return {};
  }

  const result: Partial<IAvatarSettingsSnapshot> = {};
  if (typeof parsed.avatarEnabled === 'boolean') {
    result.avatarEnabled = parsed.avatarEnabled;
  }
  if (isSupportedRealtimeVoiceId(parsed.voiceId)) {
    result.voiceId = parsed.voiceId;
  }
  if (parsed.personality && VALID_PERSONALITY_SET.has(parsed.personality)) {
    result.personality = parsed.personality as PersonalityMode;
  }
  const mappedVisage = parsed.visage ? (LEGACY_VISAGE_ALIASES[parsed.visage] || parsed.visage) : undefined;
  if (mappedVisage && VALID_VISAGE_SET.has(mappedVisage)) {
    result.visage = mappedVisage as VisageMode;
  }

  return result;
}

export function serializeAvatarSettingsPreference(settings: IAvatarSettingsSnapshot): string {
  return JSON.stringify({
    avatarEnabled: settings.avatarEnabled,
    voiceId: settings.voiceId,
    personality: settings.personality,
    visage: settings.visage,
    updatedAt: new Date().toISOString()
  });
}
