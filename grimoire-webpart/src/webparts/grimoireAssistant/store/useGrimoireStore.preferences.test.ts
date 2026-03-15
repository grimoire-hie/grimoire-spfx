import { useGrimoireStore } from './useGrimoireStore';

describe('useGrimoireStore preference dirty flags', () => {
  beforeEach(() => {
    useGrimoireStore.setState({
      avatarSettingsDirty: false,
      assistantSettingsDirty: false
    });
  });

  it('marks and clears avatar settings as dirty', () => {
    useGrimoireStore.getState().markAvatarSettingsDirty();
    expect(useGrimoireStore.getState().avatarSettingsDirty).toBe(true);

    useGrimoireStore.getState().clearAvatarSettingsDirty();
    expect(useGrimoireStore.getState().avatarSettingsDirty).toBe(false);
  });

  it('marks and clears assistant settings as dirty', () => {
    useGrimoireStore.getState().markAssistantSettingsDirty();
    expect(useGrimoireStore.getState().assistantSettingsDirty).toBe(true);

    useGrimoireStore.getState().clearAssistantSettingsDirty();
    expect(useGrimoireStore.getState().assistantSettingsDirty).toBe(false);
  });
});
