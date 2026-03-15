import {
  parseAssistantSettingsPreference,
  serializeAssistantSettingsPreference
} from './AssistantPreferenceUtils';

describe('AssistantPreferenceUtils', () => {
  it('parses valid assistant settings payload', () => {
    const parsed = parseAssistantSettingsPreference(JSON.stringify({
      publicWebSearchEnabled: true,
      copilotWebGroundingEnabled: true,
      searchRecapMode: 'always',
      updatedAt: '2026-03-04T12:00:00.000Z'
    }));
    expect(parsed).toEqual({
      publicWebSearchEnabled: true,
      copilotWebGroundingEnabled: true,
      searchRecapMode: 'always'
    });
  });

  it('migrates legacy single-flag payloads to both settings', () => {
    const parsed = parseAssistantSettingsPreference(JSON.stringify({
      copilotWebGroundingEnabled: true,
      updatedAt: '2026-03-04T12:00:00.000Z'
    }));

    expect(parsed).toEqual({
      publicWebSearchEnabled: true,
      copilotWebGroundingEnabled: true
    });
  });

  it('returns empty object for invalid payloads', () => {
    expect(parseAssistantSettingsPreference(undefined)).toEqual({});
    expect(parseAssistantSettingsPreference('')).toEqual({});
    expect(parseAssistantSettingsPreference('not-json')).toEqual({});
    expect(parseAssistantSettingsPreference(JSON.stringify({ copilotWebGroundingEnabled: 'yes' }))).toEqual({});
    expect(parseAssistantSettingsPreference(JSON.stringify({ searchRecapMode: 'sometimes' }))).toEqual({});
  });

  it('serializes parseable assistant settings', () => {
    const raw = serializeAssistantSettingsPreference({
      publicWebSearchEnabled: true,
      copilotWebGroundingEnabled: false,
      searchRecapMode: 'auto'
    });
    const reparsed = parseAssistantSettingsPreference(raw);
    expect(reparsed).toEqual({
      publicWebSearchEnabled: true,
      copilotWebGroundingEnabled: false,
      searchRecapMode: 'auto'
    });
  });
});
