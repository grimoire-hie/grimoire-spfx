/**
 * Header
 * Top bar with MCP servers info, connection status, and logs toggle.
 */

import * as React from 'react';
import { IconButton } from '@fluentui/react';
import * as strings from 'GrimoireAssistantWebPartStrings';
import type { ConnectionState } from '../../store/useGrimoireStore';
import type { ISpThemeColors } from '../../store/useGrimoireStore';

export interface IHeaderProps {
  connectionState: ConnectionState;
  mcpServerCount: number;
  logSidebarOpen: boolean;
  spThemeColors: ISpThemeColors;
  onToggleLogSidebar: () => void;
}

const leftStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 16
};

const rightStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 12
};


function getStatusDotColor(state: ConnectionState): string {
  switch (state) {
    case 'connected':
    case 'speaking':
      return '#40c040';
    case 'connecting':
      return '#f0a030';
    case 'error':
      return '#e04040';
    default:
      return '#888888';
  }
}

function getStatusLabel(state: ConnectionState): string {
  switch (state) {
    case 'connected': return strings.StatusConnected;
    case 'speaking': return strings.StatusSpeaking;
    case 'connecting': return strings.StatusConnecting;
    case 'error': return strings.StatusError;
    default: return strings.StatusOffline;
  }
}

export const Header: React.FC<IHeaderProps> = ({
  connectionState,
  mcpServerCount,
  logSidebarOpen,
  spThemeColors,
  onToggleLogSidebar
}) => {
  const dotColor = getStatusDotColor(connectionState);
  const statusLabel = getStatusLabel(connectionState);
  const headerStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '8px 16px',
    backgroundColor: spThemeColors.cardBackground,
    borderBottom: `1px solid ${spThemeColors.cardBorder}`,
    height: 48,
    flexShrink: 0
  };
  const titleStyle: React.CSSProperties = {
    fontSize: 16,
    fontWeight: 600,
    color: spThemeColors.bodyText,
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    letterSpacing: '0.2px'
  };
  const mcpBadgeStyle: React.CSSProperties = {
    fontSize: 12,
    color: spThemeColors.bodySubtext,
    backgroundColor: spThemeColors.bodyBackground,
    border: `1px solid ${spThemeColors.cardBorder}`,
    padding: '2px 10px',
    borderRadius: 12
  };

  return (
    <div style={headerStyle}>
      <div style={leftStyle}>
        <span style={titleStyle}>{strings.GrimoireTitle}</span>
        {mcpServerCount > 0 && (
          <span style={mcpBadgeStyle}>
            {mcpServerCount} MCP Server{mcpServerCount > 1 ? 's' : ''}
          </span>
        )}
      </div>

      <div style={rightStyle}>
        {/* Connection status light */}
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 6
          }}
          title={statusLabel}
        >
          <div
            style={{
              width: 8,
              height: 8,
              borderRadius: '50%',
              backgroundColor: dotColor,
              boxShadow: `0 0 6px ${dotColor}`
            }}
          />
          <span style={{
            fontSize: 12,
            color: spThemeColors.bodySubtext,
            fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif'
          }}>
            {statusLabel}
          </span>
        </div>

        {/* Logs toggle */}
        <IconButton
          iconProps={{ iconName: 'ClipboardList' }}
          ariaLabel={logSidebarOpen ? strings.CloseLogsLabel : strings.OpenLogsLabel}
          title={logSidebarOpen ? strings.CloseLogsLabel : strings.OpenLogsLabel}
          onClick={onToggleLogSidebar}
          styles={{
            root: {
              color: logSidebarOpen ? spThemeColors.bodyText : spThemeColors.bodySubtext,
              backgroundColor: logSidebarOpen ? spThemeColors.bodyBackground : 'transparent',
              border: `1px solid ${logSidebarOpen ? spThemeColors.cardBorder : 'transparent'}`,
              width: 32,
              height: 32
            },
            rootHovered: {
              backgroundColor: spThemeColors.bodyBackground,
              color: spThemeColors.bodyText
            }
          }}
        />

      </div>
    </div>
  );
};
