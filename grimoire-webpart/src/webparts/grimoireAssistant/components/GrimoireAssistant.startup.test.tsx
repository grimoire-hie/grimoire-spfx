const mockGetPreferences = jest.fn();
const mockSetPreference = jest.fn().mockResolvedValue(undefined);
const mockProbeAvailability = jest.fn();
const mockFetchBackendHealth = jest.fn();
const mockEnrichUserContextWithGraph = jest.fn(async (context) => ({
  ...context,
  department: 'Engineering'
}));

jest.mock('./layout/AppLayout', () => ({
  AppLayout: () => <div>Active shell</div>
}));

jest.mock('../services/context/PersistenceService', () => ({
  PersistenceService: {
    getPreferences: (...args: unknown[]) => mockGetPreferences(...args),
    setPreference: (...args: unknown[]) => mockSetPreference(...args)
  }
}));

jest.mock('../services/web/PublicWebSearchService', () => ({
  PublicWebSearchService: jest.fn().mockImplementation(() => ({
    probeAvailability: (...args: unknown[]) => mockProbeAvailability(...args)
  }))
}));

jest.mock('../services/startup/BackendHealthService', () => ({
  fetchBackendHealth: (...args: unknown[]) => mockFetchBackendHealth(...args)
}));

jest.mock('../services/context/ContextService', () => {
  return {
    buildBaseUserContext: (pageContext: {
      user?: { displayName?: string; email?: string; loginName?: string };
      web?: { title?: string; absoluteUrl?: string };
      site?: { absoluteUrl?: string };
      list?: { title?: string };
      listItem?: { id?: number };
    }) => ({
      displayName: pageContext.user?.displayName || 'User',
      email: pageContext.user?.email || '',
      loginName: pageContext.user?.loginName || '',
      resolvedLanguage: 'en',
      currentWebTitle: pageContext.web?.title || '',
      currentWebUrl: pageContext.web?.absoluteUrl || '',
      currentSiteTitle: pageContext.web?.title || '',
      currentSiteUrl: pageContext.site?.absoluteUrl || '',
      currentListTitle: pageContext.list?.title,
      currentItemId: pageContext.listItem?.id
    }),
    enrichUserContextWithGraph: (context: unknown) => mockEnrichUserContextWithGraph(context)
  };
});

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';

import { GrimoireAssistant } from './GrimoireAssistant';
import { useGrimoireStore } from '../store/useGrimoireStore';
import {
  resolveUserSessionScope,
  writeSessionPreferenceSeed
} from '../services/startup/StartupSessionCache';
import { serializeAvatarSettingsPreference } from '../services/context/AvatarPreferenceUtils';
import { serializeAssistantSettingsPreference } from '../services/context/AssistantPreferenceUtils';

describe('GrimoireAssistant startup split', () => {
  let container: HTMLDivElement;

  async function flushAsyncWork(): Promise<void> {
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
  }

  function resetStore(): void {
    useGrimoireStore.setState({
      uiMode: 'widget',
      spThemeColors: {
        bodyBackground: '#ffffff',
        bodyText: '#323130',
        bodySubtext: '#605e5c',
        cardBackground: '#ffffff',
        cardBorder: '#edebe9',
        isDark: false
      },
      startupPhase: 'bootstrap-widget',
      startupMetrics: {},
      avatarEnabled: true,
      avatarRenderState: 'placeholder',
      personality: 'normal',
      visage: 'classic',
      expression: 'idle',
      gazeTarget: 'none',
      avatarActionCue: undefined,
      remoteStream: undefined,
      micStream: undefined,
      isMuted: false,
      backendOk: false,
      lastHealthCheckAt: undefined,
      lastHealthCheckSource: 'none',
      micGranted: false,
      proxyConfig: undefined,
      aadHttpClient: undefined,
      userApiClient: undefined,
      userApiClientReady: false,
      getToken: undefined,
      mcpEnvironmentId: undefined,
      userContext: undefined,
      transcript: [],
      textInputValue: '',
      textChatActive: false,
      activityStatus: '',
      conversationLanguage: undefined,
      voiceId: 'alloy',
      assistantPlaybackState: 'idle',
      publicWebSearchEnabled: false,
      publicWebSearchCapability: 'unknown',
      publicWebSearchCapabilityDetail: undefined,
      publicWebSearchCapabilityCheckedAt: undefined,
      publicWebSearchCapabilitySource: 'none',
      copilotWebGroundingEnabled: false,
      searchRecapMode: 'off',
      avatarPrefsHydrated: false,
      avatarSettingsDirty: false,
      assistantSettingsDirty: false,
      blocks: [],
      freshBlockIds: [],
      activeActionBlockId: undefined,
      selectedActionIndices: [],
      focusedContext: undefined,
      mcpConnections: [],
      logEntries: [],
      logSidebarOpen: false,
      logSidebarWidth: 420,
      logFilter: 'all'
    });
  }

  function buildProps(): {
    props: React.ComponentProps<typeof GrimoireAssistant>;
    aadGetClient: jest.Mock;
    getToken: jest.Mock;
  } {
    const aadGetClient = jest.fn().mockImplementation(async (resource: string) => ({ resource }));
    const getToken = jest.fn().mockResolvedValue('token');

    return {
      props: {
        isDarkTheme: false,
        hasTeamsContext: false,
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test-key',
        backendApiResource: 'api://grimoire-backend',
        proxyBackend: 'reasoning',
        mcpEnvironmentId: 'env-123',
        spThemeColors: {
          bodyBackground: '#ffffff',
          bodyText: '#323130',
          bodySubtext: '#605e5c',
          cardBackground: '#ffffff',
          cardBorder: '#edebe9',
          isDark: false
        },
        context: {
          pageContext: {
            user: {
              displayName: 'Test User',
              email: 'test.user@example.com',
              loginName: 'test.user@example.com'
            },
            web: {
              title: 'GriMoire',
              absoluteUrl: 'https://contoso.sharepoint.com/sites/grimoire'
            },
            site: {
              absoluteUrl: 'https://contoso.sharepoint.com/sites/grimoire'
            }
          },
          aadHttpClientFactory: {
            getClient: aadGetClient
          },
          aadTokenProviderFactory: {
            getTokenProvider: jest.fn().mockResolvedValue({
              getToken
            })
          }
        } as never
      },
      aadGetClient,
      getToken
    };
  }

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    resetStore();
    sessionStorage.clear();
    mockGetPreferences.mockResolvedValue({});
    mockSetPreference.mockResolvedValue(undefined);
    mockProbeAvailability.mockResolvedValue({ status: 'available' });
    mockFetchBackendHealth.mockResolvedValue({
      backendOk: true,
      checkedAt: Date.now(),
      source: 'network',
      durationMs: 12
    });
    mockEnrichUserContextWithGraph.mockImplementation(async (context) => ({
      ...context,
      department: 'Engineering'
    }));

    Object.defineProperty(window, 'requestAnimationFrame', {
      configurable: true,
      writable: true,
      value: (callback: FrameRequestCallback) => {
        callback(0);
        return 1;
      }
    });
    Object.defineProperty(window, 'cancelAnimationFrame', {
      configurable: true,
      writable: true,
      value: jest.fn()
    });
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    jest.clearAllMocks();
    sessionStorage.clear();
  });

  it('keeps widget bootstrap free of deferred network work while seeding base context immediately', async () => {
    const { props, aadGetClient } = buildProps();

    await act(async () => {
      ReactDom.render(<GrimoireAssistant {...props} />, container);
      await flushAsyncWork();
    });

    expect(container.textContent).toContain('Grimoire');
    expect(useGrimoireStore.getState().userContext?.displayName).toBe('Test User');
    expect(useGrimoireStore.getState().userContext?.department).toBeUndefined();
    expect(mockEnrichUserContextWithGraph).not.toHaveBeenCalled();
    expect(aadGetClient).not.toHaveBeenCalled();
    expect(mockGetPreferences).not.toHaveBeenCalled();
    expect(mockFetchBackendHealth).not.toHaveBeenCalled();
    expect(mockProbeAvailability).not.toHaveBeenCalled();
  });

  it('renders the active shell immediately on Start while deferred startup work continues in the background', async () => {
    const deferredPreferences = new Promise<Record<string, string>>(() => undefined);
    mockGetPreferences.mockReturnValue(deferredPreferences);
    const { props, aadGetClient } = buildProps();

    await act(async () => {
      ReactDom.render(<GrimoireAssistant {...props} />, container);
      await flushAsyncWork();
    });

    const startButton = Array.from(container.querySelectorAll('button')).find((button) => button.textContent === 'Start');
    expect(startButton).toBeDefined();

    await act(async () => {
      startButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await flushAsyncWork();
    });

    expect(container.textContent).toContain('Active shell');
    expect(aadGetClient).toHaveBeenCalled();
    expect(mockEnrichUserContextWithGraph).toHaveBeenCalledTimes(1);
    expect(mockFetchBackendHealth).toHaveBeenCalledTimes(1);
    expect(mockGetPreferences).toHaveBeenCalledTimes(1);

    ReactDom.unmountComponentAtNode(container);
  });

  it('applies session-stored avatar preferences before backend hydration completes', async () => {
    const { props } = buildProps();
    writeSessionPreferenceSeed(
      resolveUserSessionScope('test.user@example.com'),
      serializeAvatarSettingsPreference({
        avatarEnabled: false,
        voiceId: 'alloy',
        personality: 'normal',
        visage: 'classic'
      }),
      serializeAssistantSettingsPreference({
        publicWebSearchEnabled: false,
        copilotWebGroundingEnabled: false,
        searchRecapMode: 'off'
      })
    );

    await act(async () => {
      ReactDom.render(<GrimoireAssistant {...props} />, container);
      await flushAsyncWork();
    });

    expect(useGrimoireStore.getState().avatarEnabled).toBe(false);
    expect(useGrimoireStore.getState().avatarPrefsHydrated).toBe(false);
    expect(mockGetPreferences).not.toHaveBeenCalled();

    ReactDom.unmountComponentAtNode(container);
  });
});
