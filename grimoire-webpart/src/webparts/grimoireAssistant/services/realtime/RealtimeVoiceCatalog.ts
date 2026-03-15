import type { IDropdownOption } from '@fluentui/react/lib/Dropdown';

export const SUPPORTED_REALTIME_VOICE_IDS = ['alloy', 'echo', 'shimmer', 'marin', 'cedar'] as const;
export type RealtimeVoiceId = typeof SUPPORTED_REALTIME_VOICE_IDS[number];

export const DEFAULT_REALTIME_VOICE_ID: RealtimeVoiceId = 'alloy';

const SUPPORTED_REALTIME_VOICE_ID_SET = new Set<string>(SUPPORTED_REALTIME_VOICE_IDS);

export const REALTIME_VOICE_OPTIONS: IDropdownOption[] = [
  { key: 'alloy', text: 'Alloy - Neutral' },
  { key: 'echo', text: 'Echo - Deep' },
  { key: 'shimmer', text: 'Shimmer - Bright' },
  { key: 'marin', text: 'Marin - Professional' },
  { key: 'cedar', text: 'Cedar - Natural' }
];

export function isSupportedRealtimeVoiceId(value: string | undefined): value is RealtimeVoiceId {
  return typeof value === 'string' && SUPPORTED_REALTIME_VOICE_ID_SET.has(value);
}

export function normalizeRealtimeVoiceId(value: string | undefined): RealtimeVoiceId {
  return isSupportedRealtimeVoiceId(value) ? value : DEFAULT_REALTIME_VOICE_ID;
}
