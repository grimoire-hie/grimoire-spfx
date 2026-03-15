/**
 * AvatarPanel
 * Left column — compact top bar (status), full-height particle avatar,
 * voice controls and text input at bottom.
 */

import * as React from 'react';
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import type { IButtonStyles } from '@fluentui/react/lib/Button';
import { shallow } from 'zustand/shallow';
import * as strings from 'GrimoireAssistantWebPartStrings';
import { ParticleAvatar } from './ParticleAvatar';
import { TranscriptOverlay } from './TranscriptOverlay';
import { TextInput } from './TextInput';
import { AvatarSettingsPanel } from './AvatarSettingsPanel';
import { generateVisageTemplate, randomizeFaceParams } from '../../services/avatar/FaceTemplateData';
import { PERSONALITIES } from '../../services/avatar/PersonalityEngine';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { useVoiceSession } from '../../hooks/useVoiceSession';
import type { ConnectionState } from '../../store/useGrimoireStore';

// ─── Pulse animation (injected once) ────────────────────────────

const PULSE_STYLE_ID = 'grimoire-pulse-anim';
const SHAREPOINT_BLUE = '#0f6cbd';
const SHAREPOINT_BLUE_HOVER = '#115ea3';

function ensurePulseStyle(): void {
  if (typeof document === 'undefined') return;
  if (document.getElementById(PULSE_STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = PULSE_STYLE_ID;
  style.textContent = `@keyframes grimPulse {
  0%, 100% { opacity: 1; transform: scale(1); box-shadow: 0 0 4px currentColor; }
  50% { opacity: 0.4; transform: scale(0.7); box-shadow: 0 0 8px currentColor; }
}`;
  document.head.appendChild(style);
}

// ─── Status dot helpers ─────────────────────────────────────────

function statusDotColor(ok: boolean): string {
  return ok ? '#40c040' : '#e04040';
}

function voiceDotColor(state: ConnectionState): string {
  if (state === 'connected' || state === 'speaking') return '#40c040';
  if (state === 'connecting') return '#f0a030';
  return '#e04040';
}

function withAlpha(color: string, alpha: number): string {
  const normalizedAlpha = Math.max(0, Math.min(1, alpha));
  const hex = color.trim();
  const shortHex = /^#([0-9a-fA-F]{3})$/;
  const longHex = /^#([0-9a-fA-F]{6})$/;
  const shortMatch = hex.match(shortHex);
  if (shortMatch) {
    const [r, g, b] = shortMatch[1].split('').map((part) => Number.parseInt(`${part}${part}`, 16));
    return `rgba(${r}, ${g}, ${b}, ${normalizedAlpha})`;
  }
  const longMatch = hex.match(longHex);
  if (longMatch) {
    const raw = longMatch[1];
    const r = Number.parseInt(raw.slice(0, 2), 16);
    const g = Number.parseInt(raw.slice(2, 4), 16);
    const b = Number.parseInt(raw.slice(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${normalizedAlpha})`;
  }
  return color;
}

const dotCss = (color: string, pulsing?: boolean): React.CSSProperties => ({
  width: 8,
  height: 8,
  borderRadius: '50%',
  backgroundColor: color,
  boxShadow: `0 0 4px ${color}`,
  color,
  flexShrink: 0,
  animation: pulsing ? 'grimPulse 1.2s ease-in-out infinite' : undefined
});

// ─── Styles ────────────────────────────────────────────────────

const panelStyle: React.CSSProperties = {
  position: 'relative',
  display: 'flex',
  flexDirection: 'column',
  flex: 1,
  minHeight: 0,
  overflow: 'hidden'
};

const topBarBaseStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  padding: '6px 12px',
  backdropFilter: 'blur(8px)',
  zIndex: 10,
  gap: 8,
  flexShrink: 0
};

const statusGroupStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  minWidth: 0
};

const actionGroupStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6
};

const statusItemStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 4
};

const statusLabelBaseStyle: React.CSSProperties = {
  fontSize: 10,
  fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  userSelect: 'none'
};

const MCP_ENDPOINT_DESCRIPTIONS: Record<string, string> = {
  'SharePoint & OneDrive': 'Files, folders, sites, and sharing',
  'SharePoint Lists': 'Lists, columns, and list items',
  'Outlook Mail': 'Email search, read, and send operations',
  'Outlook Calendar': 'Calendar events, availability, and meetings',
  'Microsoft Teams': 'Teams chats, channels, and messages',
  'User Profile': 'People profiles and org hierarchy',
  'Copilot Search': 'Cross-workload M365 search',
  'Word Documents': 'Create and edit Word content'
};

interface IMcpEndpointStatus {
  name: string;
  description?: string;
  connectedAtMs: number;
}

function getMcpStatusSummary(endpoints: IMcpEndpointStatus[]): { shortLabel: string; tooltip: string } {
  if (endpoints.length === 0) {
    return {
      shortLabel: '',
      tooltip: 'No active MCP connections'
    };
  }

  const firstName = endpoints[0].name;
  const shortLabel = endpoints.length === 1 ? firstName : `${firstName} +${endpoints.length - 1}`;
  const lines = endpoints.map((endpoint) => (
    endpoint.description
      ? `- ${endpoint.name}: ${endpoint.description}`
      : `- ${endpoint.name}`
  ));

  return {
    shortLabel,
    tooltip: `${endpoints.length} MCP endpoint${endpoints.length > 1 ? 's' : ''} connected\n${lines.join('\n')}`
  };
}

const avatarContainerStyle: React.CSSProperties = {
  flex: 1,
  width: '100%',
  position: 'relative',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  minHeight: 0,
  overflow: 'hidden'
};

const transcriptOverlayBaseStyle: React.CSSProperties = {
  position: 'absolute',
  left: 0,
  right: 0,
  bottom: 0,
  maxHeight: 240,
  overflowY: 'auto',
  padding: '24px 16px 12px'
};

const controlsRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  gap: 12,
  padding: '6px 16px',
  flexShrink: 0
};

const activityStatusStyle: React.CSSProperties = {
  padding: '6px 16px 2px',
  fontSize: 12,
  lineHeight: 1.4,
  textAlign: 'center',
  flexShrink: 0
};

const bottomStyle: React.CSSProperties = {
  width: '100%',
  padding: '0 12px 12px',
  flexShrink: 0,
  boxSizing: 'border-box'
};

const avatarPlaceholderStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 16,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  textAlign: 'center',
  borderRadius: 16,
  backdropFilter: 'blur(4px)',
  pointerEvents: 'none'
};

export interface IAvatarPanelProps {
  isSettingsOpen: boolean;
  onOpenSettings: () => void;
  onDismissSettings: () => void;
}

export const AvatarPanel: React.FC<IAvatarPanelProps> = ({
  isSettingsOpen,
  onOpenSettings,
  onDismissSettings
}) => {
  const {
    avatarEnabled,
    avatarRenderState,
    connectionState,
    personality,
    visage,
    expression,
    gazeTarget,
    avatarActionCue,
    remoteStream,
    micStream,
    backendOk,
    lastHealthCheckAt,
    micGranted,
    transcript,
    activityStatus,
    textInputValue,
    mcpConnections,
    logSidebarOpen,
    proxyConfig,
    spThemeColors,
    setLogSidebarOpen,
    setTextInputValue
  } = useGrimoireStore((s) => ({
    avatarEnabled: s.avatarEnabled,
    avatarRenderState: s.avatarRenderState,
    connectionState: s.connectionState,
    personality: s.personality,
    visage: s.visage,
    expression: s.expression,
    gazeTarget: s.gazeTarget,
    avatarActionCue: s.avatarActionCue,
    remoteStream: s.remoteStream,
    micStream: s.micStream,
    backendOk: s.backendOk,
    lastHealthCheckAt: s.lastHealthCheckAt,
    micGranted: s.micGranted,
    transcript: s.transcript,
    activityStatus: s.activityStatus,
    textInputValue: s.textInputValue,
    mcpConnections: s.mcpConnections,
    logSidebarOpen: s.logSidebarOpen,
    proxyConfig: s.proxyConfig,
    spThemeColors: s.spThemeColors,
    setLogSidebarOpen: s.setLogSidebarOpen,
    setTextInputValue: s.setTextInputValue
  }), shallow);

  const { connect, reconnect, disconnect, sendText } = useVoiceSession();

  const classicFaceParams = React.useMemo(() => randomizeFaceParams(), []);
  const faceParams = React.useMemo(
    () => (visage === 'classic' ? classicFaceParams : {}),
    [visage, classicFaceParams]
  );
  const faceTemplate = React.useMemo(
    () => generateVisageTemplate(visage, faceParams),
    [visage, faceParams]
  );
  const transcriptScrollRef = React.useRef<HTMLDivElement>(null);

  const isConnectedOrSpeaking = connectionState === 'connected' || connectionState === 'speaking';
  const isConnecting = connectionState === 'connecting';
  const hasProxy = !!proxyConfig;
  const apiPending = hasProxy && typeof lastHealthCheckAt !== 'number';

  const personalityConfig = PERSONALITIES[personality];
  const bgStyle: React.CSSProperties = {
    ...panelStyle,
    backgroundColor: spThemeColors.cardBackground,
    color: spThemeColors.bodyText
  };
  const topBarStyle: React.CSSProperties = {
    ...topBarBaseStyle,
    backgroundColor: withAlpha(spThemeColors.bodyBackground, spThemeColors.isDark ? 0.8 : 0.92),
    borderBottom: `1px solid ${spThemeColors.cardBorder}`
  };
  const statusLabelStyle: React.CSSProperties = {
    ...statusLabelBaseStyle,
    color: spThemeColors.bodySubtext
  };
  const mcpEndpointLabelStyle: React.CSSProperties = {
    ...statusLabelStyle,
    maxWidth: 150,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    opacity: 0.9
  };
  const transcriptOverlayStyle: React.CSSProperties = {
    ...transcriptOverlayBaseStyle,
    background: `linear-gradient(180deg, ${withAlpha(spThemeColors.cardBackground, 0)} 0%, ${withAlpha(spThemeColors.cardBackground, spThemeColors.isDark ? 0.82 : 0.92)} 100%)`
  };
  const transcriptPanelStyle: React.CSSProperties = {
    position: 'absolute',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    overflowY: 'auto',
    padding: '12px 16px',
    boxSizing: 'border-box',
    display: 'flex',
    flexDirection: 'column'
  };
  const transcriptPanelContentStyle: React.CSSProperties = {
    width: '100%',
    marginTop: 'auto'
  };
  const controlsSurfaceStyle: React.CSSProperties = {
    backgroundColor: spThemeColors.cardBackground,
    borderTop: `1px solid ${spThemeColors.cardBorder}`
  };
  const activityStatusTextStyle: React.CSSProperties = {
    ...activityStatusStyle,
    color: spThemeColors.bodySubtext,
    backgroundColor: spThemeColors.cardBackground,
    borderTop: activityStatus ? `1px solid ${spThemeColors.cardBorder}` : 'none'
  };
  const smallIconButtonStyles: IButtonStyles = {
    root: {
      width: 28,
      height: 28,
      padding: 0,
      minWidth: 28,
      background: spThemeColors.bodyBackground,
      color: spThemeColors.bodySubtext,
      border: `1px solid ${spThemeColors.cardBorder}`,
      borderRadius: 4
    },
    rootHovered: {
      background: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.2 : 0.08),
      color: spThemeColors.bodyText,
      border: `1px solid ${spThemeColors.cardBorder}`
    },
    icon: { fontSize: 14, fontWeight: 600 }
  };
  const logsButtonStyles: IButtonStyles = {
    root: {
      minWidth: 76,
      height: 28,
      padding: '0 10px',
      borderRadius: 999,
      background: logSidebarOpen ? SHAREPOINT_BLUE : spThemeColors.bodyBackground,
      color: logSidebarOpen ? '#ffffff' : spThemeColors.bodySubtext,
      border: `1px solid ${logSidebarOpen ? SHAREPOINT_BLUE : spThemeColors.cardBorder}`
    },
    rootHovered: {
      background: logSidebarOpen
        ? SHAREPOINT_BLUE_HOVER
        : withAlpha(SHAREPOINT_BLUE, spThemeColors.isDark ? 0.2 : 0.08),
      color: logSidebarOpen ? '#ffffff' : spThemeColors.bodyText,
      border: `1px solid ${logSidebarOpen ? SHAREPOINT_BLUE_HOVER : spThemeColors.cardBorder}`
    },
    icon: {
      color: logSidebarOpen ? '#ffffff' : spThemeColors.bodySubtext,
      fontSize: 12
    },
    iconHovered: {
      color: logSidebarOpen ? '#ffffff' : spThemeColors.bodyText
    },
    label: {
      fontSize: 11,
      fontWeight: 600
    }
  };
  const voiceButtonStyles: IButtonStyles = {
    root: {
      width: 48,
      height: 48,
      borderRadius: '50%',
      border: `2px solid ${spThemeColors.cardBorder}`,
      background: spThemeColors.bodyBackground,
      color: spThemeColors.bodyText
    },
    rootHovered: {
      background: withAlpha('#40c040', spThemeColors.isDark ? 0.22 : 0.12),
      borderColor: withAlpha('#40c040', 0.7)
    },
    icon: { color: spThemeColors.bodyText, fontSize: 20 }
  };
  const transcriptScrollClassName = spThemeColors.isDark ? 'grimoire-scroll-dark' : 'grimoire-scroll';
  const showAvatar = avatarEnabled;
  const showTranscriptOnly = !avatarEnabled && transcript.length > 0;
  const showAvatarPlaceholder = showAvatar
    && (avatarRenderState === 'placeholder' || avatarRenderState === 'svg-loading');
  const placeholderSurfaceStyle: React.CSSProperties = {
    ...avatarPlaceholderStyle,
    background: withAlpha(spThemeColors.bodyBackground, spThemeColors.isDark ? 0.72 : 0.86),
    border: `1px dashed ${spThemeColors.cardBorder}`,
    color: spThemeColors.bodySubtext
  };

  React.useEffect(() => { ensurePulseStyle(); }, []);

  React.useEffect(() => {
    if (transcriptScrollRef.current) {
      transcriptScrollRef.current.scrollTop = transcriptScrollRef.current.scrollHeight;
    }
  }, [transcript.length]);

  const handleConnect = React.useCallback((): void => {
    connect().catch(() => { /* errors handled in hook */ });
  }, [connect]);

  const handleTextSend = React.useCallback((text: string): void => {
    sendText(text);
  }, [sendText]);

  const handleToggleLogs = React.useCallback((): void => {
    setLogSidebarOpen(!logSidebarOpen);
  }, [logSidebarOpen, setLogSidebarOpen]);

  const connectedMcpEndpoints = React.useMemo((): IMcpEndpointStatus[] => {
    const deduped = new Map<string, IMcpEndpointStatus>();
    mcpConnections.forEach((connection) => {
      if (connection.state !== 'connected') return;
      const key = `${connection.serverUrl}|${connection.serverName}`;
      if (deduped.has(key)) return;

      const connectedAtMs = connection.connectedAt instanceof Date
        ? connection.connectedAt.getTime()
        : new Date(connection.connectedAt).getTime();

      deduped.set(key, {
        name: connection.serverName,
        description: MCP_ENDPOINT_DESCRIPTIONS[connection.serverName],
        connectedAtMs: Number.isFinite(connectedAtMs) ? connectedAtMs : 0
      });
    });

    return Array.from(deduped.values()).sort((a, b) => b.connectedAtMs - a.connectedAtMs);
  }, [mcpConnections]);

  const mcpCount = connectedMcpEndpoints.length;
  const mcpStatus = React.useMemo(
    () => getMcpStatusSummary(connectedMcpEndpoints),
    [connectedMcpEndpoints]
  );

  return (
    <div style={bgStyle}>
      <div style={topBarStyle}>
        <div style={statusGroupStyle}>
          <div
            style={statusItemStyle}
            title={!hasProxy ? strings.NoBackendTooltip : apiPending ? strings.CheckingApiTooltip : backendOk ? strings.BackendConnectedTooltip : strings.BackendUnavailableTooltip}
          >
            <div
              style={dotCss(
                !hasProxy ? withAlpha(spThemeColors.bodySubtext, 0.55) : apiPending ? '#f0a030' : statusDotColor(backendOk),
                apiPending
              )}
            />
            <span style={statusLabelStyle}>{!hasProxy ? strings.NoApiStatus : apiPending ? strings.ApiPendingStatus : strings.ApiStatus}</span>
          </div>

          <div
            style={{ ...statusItemStyle, cursor: isConnectedOrSpeaking ? 'pointer' : 'default' }}
            title={isConnectedOrSpeaking ? strings.ClickToDisconnectTooltip : isConnecting ? strings.ConnectingVoiceTooltip : strings.VoiceDisconnectedTooltip}
            onClick={isConnectedOrSpeaking ? () => disconnect() : undefined}
          >
            <div style={dotCss(voiceDotColor(connectionState), isConnecting)} />
            <span style={{
              ...statusLabelStyle,
              textDecoration: isConnectedOrSpeaking ? 'underline' : 'none',
              cursor: isConnectedOrSpeaking ? 'pointer' : 'default'
            }}>{isConnecting ? strings.VoicePendingStatus : strings.VoiceStatus}</span>
          </div>

          <div style={statusItemStyle} title={micGranted ? strings.MicAllowedTooltip : isConnecting ? strings.MicRequestingTooltip : strings.MicBlockedTooltip}>
            <div style={dotCss(isConnecting ? '#f0a030' : statusDotColor(micGranted), isConnecting && !micGranted)} />
            <span style={statusLabelStyle}>{isConnecting && !micGranted ? strings.MicPendingStatus : strings.MicStatus}</span>
          </div>

          <div style={{ ...statusItemStyle, minWidth: 0 }} title={mcpStatus.tooltip}>
            <div style={dotCss(mcpCount > 0 ? '#40c040' : withAlpha(spThemeColors.bodySubtext, 0.5))} />
            <span style={statusLabelStyle}>{mcpCount > 0 ? `${strings.McpLabel} (${mcpCount})` : strings.McpLabel}</span>
            {mcpStatus.shortLabel && (
              <span style={mcpEndpointLabelStyle}>{mcpStatus.shortLabel}</span>
            )}
          </div>
        </div>

        <div style={actionGroupStyle}>
          <DefaultButton
            text={strings.LogsButton}
            iconProps={{ iconName: 'ClipboardList' }}
            title={logSidebarOpen ? strings.CloseLogsLabel : strings.OpenLogsLabel}
            ariaLabel={logSidebarOpen ? strings.CloseLogsLabel : strings.OpenLogsLabel}
            styles={logsButtonStyles}
            onClick={handleToggleLogs}
          />
          <IconButton
            iconProps={{ iconName: 'Settings' }}
            title={strings.AvatarSettingsTitle}
            styles={smallIconButtonStyles}
            onClick={onOpenSettings}
          />
        </div>
      </div>

      <div style={avatarContainerStyle}>
        {showAvatar ? (
          <>
            <ParticleAvatar
              faceTemplate={faceTemplate}
              visage={visage}
              personality={personality}
              expression={expression}
              remoteStream={remoteStream}
              micStream={micStream}
              actionCue={avatarActionCue}
              isActive={true}
              eyeGlow={personalityConfig.eyeGlow}
              gazeTarget={gazeTarget}
            />
            {showAvatarPlaceholder && (
              <div style={placeholderSurfaceStyle}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>
                    {strings.PreparingAvatar}
                  </div>
                  <div style={{ fontSize: 11 }}>
                    {strings.PreparingAvatarDetail}
                  </div>
                </div>
              </div>
            )}

            {transcript.length > 0 && (
              <div ref={transcriptScrollRef} className={transcriptScrollClassName} style={transcriptOverlayStyle}>
                <TranscriptOverlay entries={transcript} maxVisible={6} />
              </div>
            )}
          </>
        ) : showTranscriptOnly ? (
          transcript.length > 0 && (
            <div ref={transcriptScrollRef} className={transcriptScrollClassName} style={transcriptPanelStyle}>
              <div style={transcriptPanelContentStyle}>
                <TranscriptOverlay entries={transcript} maxVisible={transcript.length} fadeOlderEntries={false} />
              </div>
            </div>
          )
        ) : null}
      </div>

      {activityStatus && (
        <div aria-live="polite" style={activityStatusTextStyle}>
          {activityStatus}
        </div>
      )}

      <div style={{ ...controlsRowStyle, ...controlsSurfaceStyle }}>
        {!isConnectedOrSpeaking && (
          <IconButton
            iconProps={{ iconName: 'Speech' }}
            title={strings.ConnectVoice}
            styles={voiceButtonStyles}
            onClick={handleConnect}
            disabled={isConnecting || !hasProxy}
          />
        )}
        {isConnectedOrSpeaking && (
          <IconButton
            iconProps={{ iconName: 'PowerButton' }}
            title={strings.TurnAudioOff}
            styles={voiceButtonStyles}
            onClick={disconnect}
          />
        )}
      </div>

      <div style={{ ...bottomStyle, ...controlsSurfaceStyle }}>
        <TextInput
          value={textInputValue}
          onChange={setTextInputValue}
          onSend={handleTextSend}
          disabled={!hasProxy}
          placeholder={hasProxy ? strings.TypeAMessage : strings.NoBackendConfigured}
        />
      </div>

      <AvatarSettingsPanel
        isOpen={isSettingsOpen}
        onDismiss={onDismissSettings}
        voiceConnected={isConnectedOrSpeaking || isConnecting}
        onVoiceReconnect={() => {
          reconnect().catch(() => { /* errors handled in hook */ });
        }}
      />
    </div>
  );
};
