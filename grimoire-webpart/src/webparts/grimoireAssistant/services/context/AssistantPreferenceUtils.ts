export const ASSISTANT_SETTINGS_KEY = 'assistant.settings.v1';
export type SearchRecapMode = 'off' | 'auto' | 'always';

export interface IAssistantSettingsSnapshot {
  publicWebSearchEnabled: boolean;
  copilotWebGroundingEnabled: boolean;
  searchRecapMode: SearchRecapMode;
}

interface IAssistantSettingsPayload {
  publicWebSearchEnabled?: boolean;
  copilotWebGroundingEnabled?: boolean;
  searchRecapMode?: SearchRecapMode;
  updatedAt?: string;
}

export function parseAssistantSettingsPreference(raw: string | undefined): Partial<IAssistantSettingsSnapshot> {
  if (!raw) return {};

  let parsed: IAssistantSettingsPayload;
  try {
    parsed = JSON.parse(raw) as IAssistantSettingsPayload;
  } catch {
    return {};
  }

  const result: Partial<IAssistantSettingsSnapshot> = {};
  const hasLegacyCopilotFlag = typeof parsed.copilotWebGroundingEnabled === 'boolean';
  if (typeof parsed.publicWebSearchEnabled === 'boolean') {
    result.publicWebSearchEnabled = parsed.publicWebSearchEnabled;
  } else if (hasLegacyCopilotFlag) {
    result.publicWebSearchEnabled = parsed.copilotWebGroundingEnabled;
  }
  if (typeof parsed.copilotWebGroundingEnabled === 'boolean') {
    result.copilotWebGroundingEnabled = parsed.copilotWebGroundingEnabled;
  }
  if (parsed.searchRecapMode === 'off' || parsed.searchRecapMode === 'auto' || parsed.searchRecapMode === 'always') {
    result.searchRecapMode = parsed.searchRecapMode;
  }
  return result;
}

export function serializeAssistantSettingsPreference(settings: IAssistantSettingsSnapshot): string {
  return JSON.stringify({
    publicWebSearchEnabled: settings.publicWebSearchEnabled,
    copilotWebGroundingEnabled: settings.copilotWebGroundingEnabled,
    searchRecapMode: settings.searchRecapMode,
    updatedAt: new Date().toISOString()
  });
}
