export const DEFAULT_AVATAR_PANE_RATIO = 0.42;
export const AVATAR_PANE_MIN_WIDTH = 320;
export const AVATAR_PANE_MAX_WIDTH = 700;
export const ACTION_PANE_MIN_WIDTH = 380;
export const MAIN_PANE_RESIZE_HANDLE_WIDTH = 12;

export interface IAvatarActionPaneLayout {
  availablePaneWidth: number;
  avatarWidth: number;
  actionWidth: number;
  avatarMinWidth: number;
  avatarMaxWidth: number;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

export function clampAvatarPaneRatio(ratio: number): number {
  if (!Number.isFinite(ratio)) {
    return DEFAULT_AVATAR_PANE_RATIO;
  }

  return clamp(ratio, 0, 1);
}

export function resolveAvatarActionPaneLayout(
  containerWidth: number,
  preferredAvatarRatio: number
): IAvatarActionPaneLayout {
  const normalizedContainerWidth = Math.max(0, Math.round(containerWidth));
  const availablePaneWidth = Math.max(0, normalizedContainerWidth - MAIN_PANE_RESIZE_HANDLE_WIDTH);

  // If the container gets too narrow, preserve the action pane minimum first.
  const avatarMinWidth = Math.min(
    AVATAR_PANE_MIN_WIDTH,
    Math.max(0, availablePaneWidth - ACTION_PANE_MIN_WIDTH)
  );
  const avatarMaxWidth = Math.min(
    AVATAR_PANE_MAX_WIDTH,
    Math.max(avatarMinWidth, availablePaneWidth - ACTION_PANE_MIN_WIDTH)
  );
  const targetAvatarWidth = Math.round(availablePaneWidth * clampAvatarPaneRatio(preferredAvatarRatio));
  const avatarWidth = clamp(targetAvatarWidth, avatarMinWidth, avatarMaxWidth);

  return {
    availablePaneWidth,
    avatarWidth,
    actionWidth: Math.max(0, availablePaneWidth - avatarWidth),
    avatarMinWidth,
    avatarMaxWidth
  };
}

export function resolveAvatarPaneRatioFromPointer(
  clientX: number,
  containerLeft: number,
  containerWidth: number
): number {
  const { availablePaneWidth, avatarMinWidth, avatarMaxWidth } = resolveAvatarActionPaneLayout(
    containerWidth,
    DEFAULT_AVATAR_PANE_RATIO
  );

  if (availablePaneWidth <= 0) {
    return DEFAULT_AVATAR_PANE_RATIO;
  }

  const avatarWidth = clamp(Math.round(clientX - containerLeft), avatarMinWidth, avatarMaxWidth);
  return clampAvatarPaneRatio(avatarWidth / availablePaneWidth);
}
