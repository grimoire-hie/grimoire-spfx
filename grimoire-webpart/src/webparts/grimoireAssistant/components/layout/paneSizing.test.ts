import {
  ACTION_PANE_MIN_WIDTH,
  AVATAR_PANE_MAX_WIDTH,
  AVATAR_PANE_MIN_WIDTH,
  DEFAULT_AVATAR_PANE_RATIO,
  MAIN_PANE_RESIZE_HANDLE_WIDTH,
  resolveAvatarActionPaneLayout,
  resolveAvatarPaneRatioFromPointer
} from './paneSizing';

describe('paneSizing', () => {
  it('uses the default split ratio when there is enough room for both panes', () => {
    const layout = resolveAvatarActionPaneLayout(1000, DEFAULT_AVATAR_PANE_RATIO);

    expect(layout.availablePaneWidth).toBe(1000 - MAIN_PANE_RESIZE_HANDLE_WIDTH);
    expect(layout.avatarWidth).toBe(415);
    expect(layout.actionWidth).toBe(573);
  });

  it('clamps the avatar pane to its minimum width', () => {
    const layout = resolveAvatarActionPaneLayout(1200, 0.08);

    expect(layout.avatarWidth).toBe(AVATAR_PANE_MIN_WIDTH);
    expect(layout.actionWidth).toBe(868);
  });

  it('clamps the action pane to its minimum width by limiting the avatar pane', () => {
    const layout = resolveAvatarActionPaneLayout(1000, 0.9);

    expect(layout.avatarWidth).toBe(608);
    expect(layout.actionWidth).toBe(ACTION_PANE_MIN_WIDTH);
  });

  it('clamps the avatar pane to its maximum width on wide layouts', () => {
    const layout = resolveAvatarActionPaneLayout(1600, 0.9);

    expect(layout.avatarWidth).toBe(AVATAR_PANE_MAX_WIDTH);
    expect(layout.actionWidth).toBe(888);
  });

  it('recomputes the avatar width from the stored drag ratio when the container width changes', () => {
    const ratio = resolveAvatarPaneRatioFromPointer(620, 120, 1200);
    const wideLayout = resolveAvatarActionPaneLayout(1500, ratio);
    const narrowLayout = resolveAvatarActionPaneLayout(960, ratio);
    const expectedWideAvatarWidth = Math.round((1500 - MAIN_PANE_RESIZE_HANDLE_WIDTH) * ratio);
    const expectedNarrowAvatarWidth = Math.round((960 - MAIN_PANE_RESIZE_HANDLE_WIDTH) * ratio);

    expect(wideLayout.avatarWidth).toBe(expectedWideAvatarWidth);
    expect(wideLayout.actionWidth).toBe(wideLayout.availablePaneWidth - expectedWideAvatarWidth);
    expect(narrowLayout.avatarWidth).toBe(expectedNarrowAvatarWidth);
    expect(narrowLayout.actionWidth).toBe(narrowLayout.availablePaneWidth - expectedNarrowAvatarWidth);
  });
});
