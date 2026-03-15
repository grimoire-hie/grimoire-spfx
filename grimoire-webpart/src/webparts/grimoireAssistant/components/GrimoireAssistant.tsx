/**
 * GrimoireAssistant — Root component
 * Splits startup into widget bootstrap, active shell, and deferred runtimes.
 */

import * as React from 'react';
import type { AadHttpClient } from '@microsoft/sp-http';
import { shallow } from 'zustand/shallow';
import { LayoutProvider } from '../context/LayoutContext';
import { AppLayout } from './layout/AppLayout';
import { WidgetPanel } from './widget/WidgetPanel';
import { ErrorBoundary } from './ErrorBoundary';
import { IGrimoireAssistantProps, getProxyConfig } from './IGrimoireAssistantProps';
import { DEFAULT_SP_THEME_COLORS, useGrimoireStore } from '../store/useGrimoireStore';
import {
  buildBaseUserContext,
  enrichUserContextWithGraph
} from '../services/context/ContextService';
import { PersistenceService } from '../services/context/PersistenceService';
import {
  AVATAR_SETTINGS_KEY,
  parseAvatarSettingsPreference,
  serializeAvatarSettingsPreference
} from '../services/context/AvatarPreferenceUtils';
import {
  ASSISTANT_SETTINGS_KEY,
  parseAssistantSettingsPreference,
  serializeAssistantSettingsPreference
} from '../services/context/AssistantPreferenceUtils';
import { getPreferencePersistenceDecision } from '../services/context/PreferencePersistenceGate';
import { logService } from '../services/logging/LogService';
import { PublicWebSearchService } from '../services/web/PublicWebSearchService';
import { fetchBackendHealth } from '../services/startup/BackendHealthService';
import {
  beginStartupMetric,
  completeStartupMetric,
  setStartupPhase
} from '../services/startup/StartupInstrumentation';
import {
  readCachedPublicWebProbe,
  readSessionPreferenceSeed,
  resolveProxySessionScope,
  resolveUserSessionScope,
  writeCachedPublicWebProbe,
  writeSessionPreferenceSeed
} from '../services/startup/StartupSessionCache';

function scheduleAfterPaint(callback: () => void): () => void {
  if (typeof window === 'undefined' || typeof window.requestAnimationFrame !== 'function') {
    const timeoutId = setTimeout(callback, 0);
    return () => clearTimeout(timeoutId);
  }

  const frameId = window.requestAnimationFrame(() => callback());
  return () => window.cancelAnimationFrame(frameId);
}

const CoreStoreSync: React.FC<{ props: IGrimoireAssistantProps }> = ({ props }) => {
  const setProxyConfig = useGrimoireStore((s) => s.setProxyConfig);
  const setMcpEnvironmentId = useGrimoireStore((s) => s.setMcpEnvironmentId);
  const setSpThemeColors = useGrimoireStore((s) => s.setSpThemeColors);
  const setUserContext = useGrimoireStore((s) => s.setUserContext);
  const {
    avatarEnabled,
    voiceId,
    publicWebSearchEnabled,
    copilotWebGroundingEnabled,
    searchRecapMode,
    personality,
    visage,
    setAvatarEnabled,
    setVoiceId,
    setPublicWebSearchEnabled,
    setCopilotWebGroundingEnabled,
    setSearchRecapMode,
    setPersonality,
    setVisage
  } = useGrimoireStore((s) => ({
    avatarEnabled: s.avatarEnabled,
    voiceId: s.voiceId,
    publicWebSearchEnabled: s.publicWebSearchEnabled,
    copilotWebGroundingEnabled: s.copilotWebGroundingEnabled,
    searchRecapMode: s.searchRecapMode,
    personality: s.personality,
    visage: s.visage,
    setAvatarEnabled: s.setAvatarEnabled,
    setVoiceId: s.setVoiceId,
    setPublicWebSearchEnabled: s.setPublicWebSearchEnabled,
    setCopilotWebGroundingEnabled: s.setCopilotWebGroundingEnabled,
    setSearchRecapMode: s.setSearchRecapMode,
    setPersonality: s.setPersonality,
    setVisage: s.setVisage
  }), shallow);
  const [sessionSeedReady, setSessionSeedReady] = React.useState(false);

  const userSessionScope = React.useMemo(
    () => resolveUserSessionScope(
      props.context?.pageContext?.user?.loginName || props.context?.pageContext?.user?.email
    ),
    [
      props.context?.pageContext?.user?.email,
      props.context?.pageContext?.user?.loginName
    ]
  );

  const avatarSettingsPayload = serializeAvatarSettingsPreference({
    avatarEnabled,
    voiceId,
    personality,
    visage
  });
  const assistantSettingsPayload = serializeAssistantSettingsPreference({
    publicWebSearchEnabled,
    copilotWebGroundingEnabled,
    searchRecapMode
  });

  React.useEffect(() => {
    logService.setHandler((entry) => {
      useGrimoireStore.getState().addLogEntry(entry);
    });
  }, []);

  React.useEffect(() => {
    const config = getProxyConfig(props);
    setProxyConfig(config);
  }, [props.proxyUrl, props.proxyApiKey, props.backendApiResource, props.proxyBackend, props.deploymentPrefix, setProxyConfig]);

  React.useEffect(() => {
    setMcpEnvironmentId(props.mcpEnvironmentId?.trim() || undefined);
  }, [props.mcpEnvironmentId, setMcpEnvironmentId]);

  React.useEffect(() => {
    setSpThemeColors(props.spThemeColors || DEFAULT_SP_THEME_COLORS);
  }, [props.spThemeColors, setSpThemeColors]);

  React.useEffect(() => {
    if (!props.context?.pageContext) {
      return;
    }

    const pageCtx = props.context.pageContext;
    setUserContext(buildBaseUserContext(pageCtx as Parameters<typeof buildBaseUserContext>[0]));
  }, [props.context, setUserContext]);

  React.useEffect(() => {
    setSessionSeedReady(false);
  }, [userSessionScope]);

  React.useEffect(() => {
    if (sessionSeedReady) {
      return;
    }

    const sessionSeed = readSessionPreferenceSeed(userSessionScope);
    if (sessionSeed?.avatarRaw) {
      const parsedAvatarPrefs = parseAvatarSettingsPreference(sessionSeed.avatarRaw);
      if (typeof parsedAvatarPrefs.avatarEnabled === 'boolean') {
        setAvatarEnabled(parsedAvatarPrefs.avatarEnabled);
      }
      if (parsedAvatarPrefs.voiceId) {
        setVoiceId(parsedAvatarPrefs.voiceId);
      }
      if (parsedAvatarPrefs.personality) {
        setPersonality(parsedAvatarPrefs.personality);
      }
      if (parsedAvatarPrefs.visage) {
        setVisage(parsedAvatarPrefs.visage);
      }
    }

    if (sessionSeed?.assistantRaw) {
      const parsedAssistantPrefs = parseAssistantSettingsPreference(sessionSeed.assistantRaw);
      if (typeof parsedAssistantPrefs.publicWebSearchEnabled === 'boolean') {
        setPublicWebSearchEnabled(parsedAssistantPrefs.publicWebSearchEnabled);
      }
      if (typeof parsedAssistantPrefs.copilotWebGroundingEnabled === 'boolean') {
        setCopilotWebGroundingEnabled(parsedAssistantPrefs.copilotWebGroundingEnabled);
      }
      if (parsedAssistantPrefs.searchRecapMode) {
        setSearchRecapMode(parsedAssistantPrefs.searchRecapMode);
      }
    }

    setSessionSeedReady(true);
  }, [
    sessionSeedReady,
    setAvatarEnabled,
    setCopilotWebGroundingEnabled,
    setPersonality,
    setPublicWebSearchEnabled,
    setSearchRecapMode,
    setVisage,
    setVoiceId,
    userSessionScope
  ]);

  React.useEffect(() => {
    if (!sessionSeedReady) {
      return;
    }

    writeSessionPreferenceSeed(
      userSessionScope,
      avatarSettingsPayload,
      assistantSettingsPayload
    );
  }, [
    assistantSettingsPayload,
    avatarSettingsPayload,
    sessionSeedReady,
    userSessionScope
  ]);

  return null;
};

const DeferredActiveSync: React.FC<{ props: IGrimoireAssistantProps }> = ({ props }) => {
  const setAadHttpClient = useGrimoireStore((s) => s.setAadHttpClient);
  const setUserApiClient = useGrimoireStore((s) => s.setUserApiClient);
  const setUserApiClientReady = useGrimoireStore((s) => s.setUserApiClientReady);
  const setGetToken = useGrimoireStore((s) => s.setGetToken);
  const setUserContext = useGrimoireStore((s) => s.setUserContext);
  const {
    proxyConfig,
    getToken,
    userApiClient,
    userApiClientReady,
    aadHttpClient,
    avatarEnabled,
    voiceId,
    publicWebSearchEnabled,
    publicWebSearchCapability,
    copilotWebGroundingEnabled,
    searchRecapMode,
    personality,
    visage,
    avatarPrefsHydrated,
    avatarSettingsDirty,
    assistantSettingsDirty,
    setAvatarEnabled,
    setHealthCheckState,
    setPublicWebSearchEnabled,
    setPublicWebSearchCapability,
    setCopilotWebGroundingEnabled,
    setSearchRecapMode,
    setPersonality,
    setVisage,
    setVoiceId,
    setAvatarPrefsHydrated,
    clearAvatarSettingsDirty,
    clearAssistantSettingsDirty
  } = useGrimoireStore((s) => ({
    proxyConfig: s.proxyConfig,
    getToken: s.getToken,
    userApiClient: s.userApiClient,
    userApiClientReady: s.userApiClientReady,
    aadHttpClient: s.aadHttpClient,
    avatarEnabled: s.avatarEnabled,
    voiceId: s.voiceId,
    publicWebSearchEnabled: s.publicWebSearchEnabled,
    publicWebSearchCapability: s.publicWebSearchCapability,
    copilotWebGroundingEnabled: s.copilotWebGroundingEnabled,
    searchRecapMode: s.searchRecapMode,
    personality: s.personality,
    visage: s.visage,
    avatarPrefsHydrated: s.avatarPrefsHydrated,
    avatarSettingsDirty: s.avatarSettingsDirty,
    assistantSettingsDirty: s.assistantSettingsDirty,
    setAvatarEnabled: s.setAvatarEnabled,
    setHealthCheckState: s.setHealthCheckState,
    setPublicWebSearchEnabled: s.setPublicWebSearchEnabled,
    setPublicWebSearchCapability: s.setPublicWebSearchCapability,
    setCopilotWebGroundingEnabled: s.setCopilotWebGroundingEnabled,
    setSearchRecapMode: s.setSearchRecapMode,
    setPersonality: s.setPersonality,
    setVisage: s.setVisage,
    setVoiceId: s.setVoiceId,
    setAvatarPrefsHydrated: s.setAvatarPrefsHydrated,
    clearAvatarSettingsDirty: s.clearAvatarSettingsDirty,
    clearAssistantSettingsDirty: s.clearAssistantSettingsDirty
  }), shallow);
  const avatarSettingsPersistBaselineRef = React.useRef<string | undefined>(undefined);
  const assistantSettingsPersistBaselineRef = React.useRef<string | undefined>(undefined);
  const avatarPrefsHydrationStartedRef = React.useRef(false);

  const avatarSettingsPayload = serializeAvatarSettingsPreference({
    avatarEnabled,
    voiceId,
    personality,
    visage
  });
  const assistantSettingsPayload = serializeAssistantSettingsPreference({
    publicWebSearchEnabled,
    copilotWebGroundingEnabled,
    searchRecapMode
  });

  React.useEffect(() => {
    if (props.context?.aadHttpClientFactory) {
      props.context.aadHttpClientFactory
        .getClient('https://graph.microsoft.com')
        .then((client: AadHttpClient) => {
          setAadHttpClient(client);
        })
        .catch(() => {
          setAadHttpClient(undefined);
        });

      const tokenFactory = async (resource: string): Promise<string> => {
        const tokenProvider = await props.context.aadTokenProviderFactory.getTokenProvider();
        return tokenProvider.getToken(resource);
      };
      setGetToken(tokenFactory);
      return;
    }

    setAadHttpClient(undefined);
    setGetToken(undefined);
  }, [props.context, setAadHttpClient, setGetToken]);

  React.useEffect(() => {
    setUserApiClient(undefined);
    setUserApiClientReady(false);

    if (!props.context?.aadHttpClientFactory || !proxyConfig?.backendApiResource) {
      setUserApiClientReady(true);
      return;
    }

    props.context.aadHttpClientFactory
      .getClient(proxyConfig.backendApiResource)
      .then((client: AadHttpClient) => {
        setUserApiClient(client);
      })
      .catch(() => {
        setUserApiClient(undefined);
      })
      .finally(() => {
        setUserApiClientReady(true);
      });
  }, [props.context, proxyConfig?.backendApiResource, setUserApiClient, setUserApiClientReady]);

  React.useEffect(() => {
    if (!props.context?.pageContext || !aadHttpClient) {
      return;
    }

    let cancelled = false;
    beginStartupMetric('graph-enrichment');
    const baseContext = buildBaseUserContext(
      props.context.pageContext as Parameters<typeof buildBaseUserContext>[0]
    );

    enrichUserContextWithGraph(baseContext, aadHttpClient)
      .then((context) => {
        if (cancelled) {
          return;
        }
        setUserContext(context);
        completeStartupMetric('graph-enrichment', `user=${context.loginName || context.email || 'unknown'}`);
      })
      .catch((error: Error) => {
        if (cancelled) {
          return;
        }
        completeStartupMetric('graph-enrichment', `error=${error.message}`);
      });

    return () => {
      cancelled = true;
    };
  }, [aadHttpClient, props.context, setUserContext]);

  React.useEffect(() => {
    if (!proxyConfig) {
      return;
    }

    let cancelled = false;
    beginStartupMetric('health-check');

    fetchBackendHealth(proxyConfig, { allowSessionCache: true })
      .then((result) => {
        if (cancelled) {
          return;
        }
        setHealthCheckState(result.backendOk, result.checkedAt, result.source);
        completeStartupMetric('health-check', `source=${result.source}; ok=${result.backendOk}`);
      })
      .catch((error: Error) => {
        if (cancelled) {
          return;
        }
        setHealthCheckState(false, Date.now(), 'network');
        completeStartupMetric('health-check', `error=${error.message}`);
      });

    return () => {
      cancelled = true;
    };
  }, [proxyConfig, setHealthCheckState]);

  React.useEffect(() => {
    if (!proxyConfig || avatarPrefsHydrated) {
      return;
    }

    if (!proxyConfig.backendApiResource) {
      logService.warning('system', 'Secure settings persistence is unavailable: backendApiResource is not configured.');
      setAvatarPrefsHydrated(true);
      return;
    }

    if (!userApiClientReady && !getToken) {
      return;
    }

    if (!userApiClient && !getToken) {
      logService.warning('system', 'Secure settings persistence is unavailable: no secured backend client or token provider.');
      setAvatarPrefsHydrated(true);
      return;
    }

    if (avatarPrefsHydrationStartedRef.current) {
      return;
    }

    let cancelled = false;
    avatarPrefsHydrationStartedRef.current = true;
    beginStartupMetric('preferences-hydration');

    PersistenceService.getPreferences(proxyConfig, userApiClient, getToken)
      .then((prefs) => {
        if (cancelled) {
          return;
        }
        const parsedAvatarPrefs = parseAvatarSettingsPreference(prefs[AVATAR_SETTINGS_KEY]);
        if (typeof parsedAvatarPrefs.avatarEnabled === 'boolean') {
          setAvatarEnabled(parsedAvatarPrefs.avatarEnabled);
        }
        if (parsedAvatarPrefs.voiceId) {
          setVoiceId(parsedAvatarPrefs.voiceId);
        }
        if (parsedAvatarPrefs.personality) {
          setPersonality(parsedAvatarPrefs.personality);
        }
        if (parsedAvatarPrefs.visage) {
          setVisage(parsedAvatarPrefs.visage);
        }

        const assistantPrefs = parseAssistantSettingsPreference(prefs[ASSISTANT_SETTINGS_KEY]);
        if (typeof assistantPrefs.publicWebSearchEnabled === 'boolean') {
          setPublicWebSearchEnabled(assistantPrefs.publicWebSearchEnabled);
        }
        if (typeof assistantPrefs.copilotWebGroundingEnabled === 'boolean') {
          setCopilotWebGroundingEnabled(assistantPrefs.copilotWebGroundingEnabled);
        }
        if (assistantPrefs.searchRecapMode) {
          setSearchRecapMode(assistantPrefs.searchRecapMode);
        }

        completeStartupMetric('preferences-hydration', 'source=backend');
      })
      .catch((error: Error) => {
        if (cancelled) {
          return;
        }
        logService.warning('system', `Avatar settings hydration failed: ${error.message}`);
        completeStartupMetric('preferences-hydration', `error=${error.message}`);
      })
      .finally(() => {
        if (cancelled) {
          return;
        }
        setAvatarPrefsHydrated(true);
      });

    return () => {
      cancelled = true;
    };
  }, [
    avatarPrefsHydrated,
    getToken,
    proxyConfig,
    setAvatarEnabled,
    setAvatarPrefsHydrated,
    setCopilotWebGroundingEnabled,
    setPersonality,
    setPublicWebSearchEnabled,
    setSearchRecapMode,
    setVisage,
    setVoiceId,
    userApiClient,
    userApiClientReady
  ]);

  React.useEffect(() => {
    if (!avatarPrefsHydrated) {
      avatarPrefsHydrationStartedRef.current = false;
      avatarSettingsPersistBaselineRef.current = undefined;
      assistantSettingsPersistBaselineRef.current = undefined;
    }
  }, [avatarPrefsHydrated]);

  React.useEffect(() => {
    if (!avatarPrefsHydrated || !proxyConfig || (!userApiClient && !getToken)) {
      return;
    }

    if (avatarSettingsDirty) {
      if (avatarSettingsPersistBaselineRef.current === avatarSettingsPayload) {
        clearAvatarSettingsDirty();
        return;
      }
    } else {
      const decision = getPreferencePersistenceDecision(
        avatarSettingsPersistBaselineRef.current,
        avatarSettingsPayload
      );
      if (!decision.shouldPersist) {
        avatarSettingsPersistBaselineRef.current = decision.nextBaseline;
        return;
      }
    }

    let cancelled = false;

    PersistenceService.setPreference(proxyConfig, userApiClient, getToken, AVATAR_SETTINGS_KEY, avatarSettingsPayload)
      .then(() => {
        if (cancelled) {
          return;
        }
        avatarSettingsPersistBaselineRef.current = avatarSettingsPayload;
        clearAvatarSettingsDirty();
      })
      .catch((error: Error) => {
        if (cancelled) {
          return;
        }
        logService.warning('system', `Avatar settings persistence failed: ${error.message}`);
      });

    return () => {
      cancelled = true;
    };
  }, [
    avatarPrefsHydrated,
    avatarSettingsDirty,
    avatarSettingsPayload,
    clearAvatarSettingsDirty,
    getToken,
    proxyConfig,
    userApiClient
  ]);

  React.useEffect(() => {
    if (!avatarPrefsHydrated || !proxyConfig || (!userApiClient && !getToken)) {
      return;
    }

    if (assistantSettingsDirty) {
      if (assistantSettingsPersistBaselineRef.current === assistantSettingsPayload) {
        clearAssistantSettingsDirty();
        return;
      }
    } else {
      const decision = getPreferencePersistenceDecision(
        assistantSettingsPersistBaselineRef.current,
        assistantSettingsPayload
      );
      if (!decision.shouldPersist) {
        assistantSettingsPersistBaselineRef.current = decision.nextBaseline;
        return;
      }
    }

    let cancelled = false;

    PersistenceService.setPreference(proxyConfig, userApiClient, getToken, ASSISTANT_SETTINGS_KEY, assistantSettingsPayload)
      .then(() => {
        if (cancelled) {
          return;
        }
        assistantSettingsPersistBaselineRef.current = assistantSettingsPayload;
        clearAssistantSettingsDirty();
      })
      .catch((error: Error) => {
        if (cancelled) {
          return;
        }
        logService.warning('system', `Assistant settings persistence failed: ${error.message}`);
      });

    return () => {
      cancelled = true;
    };
  }, [
    assistantSettingsDirty,
    assistantSettingsPayload,
    avatarPrefsHydrated,
    clearAssistantSettingsDirty,
    getToken,
    proxyConfig,
    userApiClient
  ]);

  React.useEffect(() => {
    if (!avatarPrefsHydrated || !publicWebSearchEnabled || !proxyConfig || publicWebSearchCapability !== 'unknown') {
      return;
    }

    let cancelled = false;
    const probeCacheScope = resolveProxySessionScope(proxyConfig);
    const cachedProbe = readCachedPublicWebProbe(probeCacheScope);

    if (cachedProbe) {
      setPublicWebSearchCapability(
        cachedProbe.status,
        cachedProbe.detail,
        cachedProbe.checkedAt,
        'session-cache'
      );
      beginStartupMetric('public-web-probe', 'source=session-cache');
      completeStartupMetric('public-web-probe', `source=session-cache; status=${cachedProbe.status}`);
      return;
    }

    beginStartupMetric('public-web-probe');
    const service = new PublicWebSearchService(proxyConfig);

    service.probeAvailability()
      .then((result) => {
        if (cancelled) {
          return;
        }
        const checkedAt = Date.now();
        setPublicWebSearchCapability(result.status, result.detail, checkedAt, 'network');
        writeCachedPublicWebProbe(probeCacheScope, {
          status: result.status,
          detail: result.detail,
          checkedAt
        });
        completeStartupMetric('public-web-probe', `source=network; status=${result.status}`);
      })
      .catch((error: Error) => {
        if (cancelled) {
          return;
        }
        const checkedAt = Date.now();
        setPublicWebSearchCapability('error', error.message, checkedAt, 'network');
        writeCachedPublicWebProbe(probeCacheScope, {
          status: 'error',
          detail: error.message,
          checkedAt
        });
        completeStartupMetric('public-web-probe', `error=${error.message}`);
      });

    return () => {
      cancelled = true;
    };
  }, [
    avatarPrefsHydrated,
    proxyConfig,
    publicWebSearchCapability,
    publicWebSearchEnabled,
    setPublicWebSearchCapability
  ]);

  return null;
};

const overlayStyle: React.CSSProperties = {
  position: 'fixed',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  zIndex: 1000
};

const ActiveOverlay: React.FC = () => {
  React.useEffect(() => {
    setStartupPhase('active-shell');
    completeStartupMetric('active-shell-paint');
  }, []);

  return (
    <div style={overlayStyle}>
      <AppLayout />
    </div>
  );
};

export const GrimoireAssistant: React.FC<IGrimoireAssistantProps> = (props) => {
  const uiMode = useGrimoireStore((s) => s.uiMode);
  const [deferredActiveSyncReady, setDeferredActiveSyncReady] = React.useState(false);

  React.useEffect(() => {
    if (uiMode !== 'active') {
      setDeferredActiveSyncReady(false);
      return;
    }

    beginStartupMetric('deferred-runtimes-start');
    const cancel = scheduleAfterPaint(() => {
      setStartupPhase('deferred-runtimes');
      setDeferredActiveSyncReady(true);
      completeStartupMetric('deferred-runtimes-start');
    });

    return cancel;
  }, [uiMode]);

  return (
    <ErrorBoundary>
      <LayoutProvider hasTeamsContext={props.hasTeamsContext}>
        <CoreStoreSync props={props} />
        {uiMode === 'widget' ? <WidgetPanel /> : <ActiveOverlay />}
        {uiMode === 'active' && deferredActiveSyncReady ? <DeferredActiveSync props={props} /> : null}
      </LayoutProvider>
    </ErrorBoundary>
  );
};

export default GrimoireAssistant;
