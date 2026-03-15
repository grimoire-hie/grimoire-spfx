jest.mock('../../hooks/useVoiceSession', () => ({
  useVoiceSession: (() => {
    const session = {
      connect: jest.fn().mockResolvedValue(undefined),
      reconnect: jest.fn().mockResolvedValue(undefined),
      disconnect: jest.fn(),
      toggleMute: jest.fn(),
      sendText: jest.fn(),
      checkHealth: jest.fn().mockResolvedValue(undefined)
    };
    return () => session;
  })()
}));

jest.mock('./ParticleAvatar', () => ({
  ParticleAvatar: () => null
}));

jest.mock('./TranscriptOverlay', () => ({
  TranscriptOverlay: () => null
}));

jest.mock('./TextInput', () => ({
  TextInput: () => null
}));

jest.mock('./AvatarSettingsPanel', () => ({
  AvatarSettingsPanel: () => null
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';

import { AvatarPanel } from './AvatarPanel';
import { DEFAULT_SP_THEME_COLORS, useGrimoireStore } from '../../store/useGrimoireStore';

describe('AvatarPanel activity status', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    useGrimoireStore.setState({
      avatarEnabled: false,
      avatarRenderState: 'placeholder',
      avatarPrefsHydrated: true,
      connectionState: 'idle',
      personality: 'normal',
      visage: 'classic',
      expression: 'idle',
      gazeTarget: 'none',
      avatarActionCue: undefined,
      remoteStream: undefined,
      micStream: undefined,
      backendOk: true,
      lastHealthCheckAt: Date.now(),
      micGranted: true,
      transcript: [],
      activityStatus: '',
      textInputValue: '',
      mcpConnections: [],
      logSidebarOpen: false,
      proxyConfig: {
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test-key',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      spThemeColors: DEFAULT_SP_THEME_COLORS
    });
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    jest.clearAllMocks();
  });

  it('renders the activity status line only when a status is present', async () => {
    await act(async () => {
      ReactDom.render(
        <AvatarPanel
          isSettingsOpen={false}
          onOpenSettings={() => undefined}
          onDismissSettings={() => undefined}
        />,
        container
      );
      await Promise.resolve();
    });

    expect(container.textContent).not.toContain('Rate-limited, retrying in 5s...');

    await act(async () => {
      useGrimoireStore.getState().setActivityStatus('Rate-limited, retrying in 5s... (1/3)');
      await Promise.resolve();
    });

    expect(container.textContent).toContain('Rate-limited, retrying in 5s... (1/3)');
    const liveRegion = container.querySelector('[aria-live="polite"]');
    expect(liveRegion?.textContent).toBe('Rate-limited, retrying in 5s... (1/3)');

    await act(async () => {
      useGrimoireStore.getState().setActivityStatus('');
      await Promise.resolve();
    });

    expect(container.textContent).not.toContain('Rate-limited, retrying in 5s... (1/3)');
    expect(container.querySelector('[aria-live="polite"]')).toBeNull();
  });

  it('shows an avatar placeholder while svg render work is still in progress', async () => {
    useGrimoireStore.setState({
      avatarEnabled: true,
      avatarPrefsHydrated: false,
      avatarRenderState: 'svg-loading'
    });

    await act(async () => {
      ReactDom.render(
        <AvatarPanel
          isSettingsOpen={false}
          onOpenSettings={() => undefined}
          onDismissSettings={() => undefined}
        />,
        container
      );
      await Promise.resolve();
    });

    expect(container.textContent).toContain('Preparing avatar...');

    ReactDom.unmountComponentAtNode(container);
  });
});
