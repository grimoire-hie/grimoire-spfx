import {
  getActionPanelCloseCopy,
  resolveActionPanelCloseAction,
  resolveEscapeCloseAction
} from './closeBehavior';

describe('closeBehavior', () => {
  it('closes settings before logs and app close on desktop', () => {
    expect(resolveEscapeCloseAction({
      isMobile: false,
      isSettingsOpen: true,
      logSidebarOpen: true
    })).toBe('close-settings');

    expect(resolveEscapeCloseAction({
      isMobile: false,
      isSettingsOpen: false,
      logSidebarOpen: true
    })).toBe('close-logs');

    expect(resolveEscapeCloseAction({
      isMobile: false,
      isSettingsOpen: false,
      logSidebarOpen: false
    })).toBe('close-grimoire');
  });

  it('closes logs before settings and app close on mobile', () => {
    expect(resolveEscapeCloseAction({
      isMobile: true,
      isSettingsOpen: true,
      logSidebarOpen: true
    })).toBe('close-logs');

    expect(resolveEscapeCloseAction({
      isMobile: true,
      isSettingsOpen: true,
      logSidebarOpen: false
    })).toBe('close-settings');

    expect(resolveEscapeCloseAction({
      isMobile: true,
      isSettingsOpen: false,
      logSidebarOpen: false
    })).toBe('close-grimoire');
  });

  it('makes the action-panel close button settings-aware', () => {
    expect(resolveActionPanelCloseAction(true)).toBe('close-settings');
    expect(resolveActionPanelCloseAction(false)).toBe('close-grimoire');
    expect(getActionPanelCloseCopy(true)).toEqual({
      ariaLabel: 'Close settings',
      title: 'Close settings'
    });
    expect(getActionPanelCloseCopy(false)).toEqual({
      ariaLabel: 'Close Grimoire (Esc)',
      title: 'Close (Esc)'
    });
  });
});
