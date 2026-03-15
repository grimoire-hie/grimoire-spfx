/**
 * StatusBar
 * 4 connection indicator dots: Backend, Voice, Microphone, MCP.
 * Floating bar at the top of the avatar panel.
 */

import * as React from 'react';
import * as strings from 'GrimoireAssistantWebPartStrings';
import type { ConnectionState } from '../../store/useGrimoireStore';

export interface IStatusBarProps {
  backendOk: boolean;
  voiceState: ConnectionState;
  micGranted: boolean;
  /** Number of active MCP server connections */
  mcpConnected: number;
}

const DOT_SIZE = 8;

const containerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 16,
  padding: '8px 16px',
  backgroundColor: 'rgba(0, 0, 0, 0.4)',
  borderRadius: 20,
  backdropFilter: 'blur(8px)',
  zIndex: 10
};

const indicatorStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6
};

const labelStyle: React.CSSProperties = {
  fontSize: 11,
  color: 'rgba(255, 255, 255, 0.7)',
  fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  userSelect: 'none'
};

function getDotColor(ok: boolean): string {
  return ok ? '#40c040' : '#e04040';
}

function getVoiceDotColor(state: ConnectionState): string {
  switch (state) {
    case 'connected':
    case 'speaking':
      return '#40c040';
    case 'connecting':
      return '#f0a030';
    default:
      return '#e04040';
  }
}

function getVoiceLabel(state: ConnectionState): string {
  switch (state) {
    case 'connected': return strings.StatusConnected;
    case 'speaking': return strings.StatusListening;
    case 'connecting': return strings.StatusConnecting;
    case 'error': return strings.StatusError;
    default: return strings.StatusIdle;
  }
}

const Dot: React.FC<{ color: string; pulse?: boolean }> = ({ color, pulse }) => (
  <div
    style={{
      width: DOT_SIZE,
      height: DOT_SIZE,
      borderRadius: '50%',
      backgroundColor: color,
      boxShadow: `0 0 ${pulse ? 8 : 4}px ${color}`,
      transition: 'background-color 0.3s, box-shadow 0.3s',
      animation: pulse ? 'pulse-dot 1.5s ease-in-out infinite' : undefined
    }}
  />
);

export const StatusBar: React.FC<IStatusBarProps> = ({
  backendOk,
  voiceState,
  micGranted,
  mcpConnected
}) => {
  return (
    <div style={containerStyle}>
      <div style={indicatorStyle} title={backendOk ? strings.BackendConnectedTooltip : strings.BackendUnavailableTooltip}>
        <Dot color={getDotColor(backendOk)} />
        <span style={labelStyle}>{strings.BackendLabel}</span>
      </div>

      <div style={indicatorStyle} title={getVoiceLabel(voiceState)}>
        <Dot
          color={getVoiceDotColor(voiceState)}
          pulse={voiceState === 'speaking'}
        />
        <span style={labelStyle}>{strings.VoiceLabel}</span>
      </div>

      <div style={indicatorStyle} title={micGranted ? strings.MicAllowed : strings.MicBlocked}>
        <Dot color={getDotColor(micGranted)} />
        <span style={labelStyle}>{strings.MicLabel}</span>
      </div>

      <div style={indicatorStyle} title={mcpConnected > 0 ? strings.McpServersTooltip.replace('{0}', String(mcpConnected)) : strings.NoMcpServersTooltip}>
        <Dot color={mcpConnected > 0 ? '#40c040' : '#888888'} />
        <span style={labelStyle}>{strings.McpLabel}</span>
      </div>
    </div>
  );
};
