export type CloseBehaviorAction = 'close-logs' | 'close-settings' | 'close-grimoire';

export interface IEscapeCloseBehaviorOptions {
  isMobile: boolean;
  logSidebarOpen: boolean;
  isSettingsOpen: boolean;
}

export function resolveEscapeCloseAction(options: IEscapeCloseBehaviorOptions): CloseBehaviorAction {
  const { isMobile, logSidebarOpen, isSettingsOpen } = options;

  if (isMobile) {
    if (logSidebarOpen) return 'close-logs';
    if (isSettingsOpen) return 'close-settings';
    return 'close-grimoire';
  }

  if (isSettingsOpen) return 'close-settings';
  if (logSidebarOpen) return 'close-logs';
  return 'close-grimoire';
}

export function resolveActionPanelCloseAction(isSettingsOpen: boolean): CloseBehaviorAction {
  return isSettingsOpen ? 'close-settings' : 'close-grimoire';
}

export function getActionPanelCloseCopy(isSettingsOpen: boolean): { ariaLabel: string; title: string } {
  if (isSettingsOpen) {
    return {
      ariaLabel: 'Close settings',
      title: 'Close settings'
    };
  }

  return {
    ariaLabel: 'Close Grimoire (Esc)',
    title: 'Close (Esc)'
  };
}
