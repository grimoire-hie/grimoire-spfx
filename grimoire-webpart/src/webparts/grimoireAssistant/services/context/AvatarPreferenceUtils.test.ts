import {
  parseAvatarSettingsPreference,
  serializeAvatarSettingsPreference
} from './AvatarPreferenceUtils';

describe('AvatarPreferenceUtils', () => {
  it('applies a fully valid preference payload', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      avatarEnabled: false,
      voiceId: 'echo',
      personality: 'funny',
      visage: 'robot',
      updatedAt: '2026-01-01T00:00:00.000Z'
    }));

    expect(parsed).toEqual({
      avatarEnabled: false,
      voiceId: 'echo',
      personality: 'funny',
      visage: 'robot'
    });
  });

  it('ignores unknown enum values and keeps valid ones', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      voiceId: 'invalid-voice',
      personality: 'invalid-personality',
      visage: 'invalid-visage'
    }));

    expect(parsed).toEqual({});
  });

  it('returns empty object for invalid or missing payloads', () => {
    expect(parseAvatarSettingsPreference(undefined)).toEqual({});
    expect(parseAvatarSettingsPreference('')).toEqual({});
    expect(parseAvatarSettingsPreference('not-json')).toEqual({});
  });

  it('maps legacy Guy Fawkes visage value to AnonyMousse', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      visage: 'guyFawkes'
    }));

    expect(parsed).toEqual({
      visage: 'anonyMousse'
    });
  });

  it('maps removed Friendly Robot visage value to Robot', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      visage: 'friendlyRobot'
    }));

    expect(parsed).toEqual({
      visage: 'robot'
    });
  });

  it('maps removed Pixel AI visage value to Robot', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      visage: 'pixelAi'
    }));

    expect(parsed).toEqual({
      visage: 'robot'
    });
  });

  it('maps removed Clip visage value to GriMoire', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      visage: 'clippy'
    }));

    expect(parsed).toEqual({
      visage: 'classic'
    });
  });

  it('maps Grimoire Test 1 to GriMoire', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      visage: 'grimoire_test_1'
    }));

    expect(parsed).toEqual({
      visage: 'classic'
    });
  });

  it('maps Grimoire Test 2 to AnonyMousse', () => {
    const parsed = parseAvatarSettingsPreference(JSON.stringify({
      visage: 'grimoire_test_2'
    }));

    expect(parsed).toEqual({
      visage: 'anonyMousse'
    });
  });

  it('serializes payload with updatedAt and parseable settings', () => {
    const raw = serializeAvatarSettingsPreference({
      avatarEnabled: false,
      voiceId: 'alloy',
      personality: 'normal',
      visage: 'classic'
    });

    const asObject = JSON.parse(raw) as Record<string, string>;
    expect(typeof asObject.updatedAt).toBe('string');

    const reparsed = parseAvatarSettingsPreference(raw);
    expect(reparsed).toEqual({
      avatarEnabled: false,
      voiceId: 'alloy',
      personality: 'normal',
      visage: 'classic'
    });
  });
});
