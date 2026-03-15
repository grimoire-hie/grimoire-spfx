import { useGrimoireStore } from './useGrimoireStore';

describe('useGrimoireStore avatar action cue', () => {
  beforeEach(() => {
    useGrimoireStore.getState().resetSession();
    useGrimoireStore.setState({ avatarActionCue: undefined });
  });

  it('sets avatarActionCue when triggerAvatarActionCue is called', () => {
    useGrimoireStore.getState().triggerAvatarActionCue('focus');
    const cue = useGrimoireStore.getState().avatarActionCue;

    expect(cue).toBeDefined();
    expect(cue?.type).toBe('focus');
    expect(typeof cue?.id).toBe('number');
    expect(cue?.at).toBeGreaterThan(0);
  });

  it('increments cue id and replaces cue type on consecutive triggers', () => {
    useGrimoireStore.getState().triggerAvatarActionCue('focus');
    const firstCue = useGrimoireStore.getState().avatarActionCue;
    expect(firstCue).toBeDefined();

    useGrimoireStore.getState().triggerAvatarActionCue('chat');
    const secondCue = useGrimoireStore.getState().avatarActionCue;
    expect(secondCue).toBeDefined();

    expect(secondCue?.id).toBeGreaterThan(firstCue?.id || 0);
    expect(secondCue?.type).toBe('chat');
  });

  it('clears avatarActionCue on resetSession', () => {
    useGrimoireStore.getState().triggerAvatarActionCue('summarize');
    expect(useGrimoireStore.getState().avatarActionCue).toBeDefined();

    useGrimoireStore.getState().resetSession();
    expect(useGrimoireStore.getState().avatarActionCue).toBeUndefined();
  });
});
