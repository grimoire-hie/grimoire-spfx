/**
 * WidgetPanel
 * Compact start mode — simple SP-themed card with a "Start Grimoire" button.
 * Rendered when uiMode === 'widget'.
 */

import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import {
  beginStartupMetric,
  completeStartupMetric,
  recordStartupMetric
} from '../../services/startup/StartupInstrumentation';

export const WidgetPanel: React.FC = () => {
  const setUiMode = useGrimoireStore((s) => s.setUiMode);
  const spTheme = useGrimoireStore((s) => s.spThemeColors);

  React.useEffect(() => {
    beginStartupMetric('widget-render');
    const frameId = typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function'
      ? window.requestAnimationFrame(() => completeStartupMetric('widget-render'))
      : undefined;

    if (typeof frameId !== 'number') {
      completeStartupMetric('widget-render');
      return undefined;
    }

    return () => window.cancelAnimationFrame(frameId);
  }, []);

  const handleStart = React.useCallback((): void => {
    recordStartupMetric('start-click');
    beginStartupMetric('active-shell-paint');
    setUiMode('active');
  }, [setUiMode]);

  const containerStyle: React.CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 12,
    padding: '32px 24px',
    borderRadius: 8,
    border: `1px solid ${spTheme.cardBorder}`,
    backgroundColor: spTheme.cardBackground
  };

  const titleStyle: React.CSSProperties = {
    fontSize: 18,
    fontWeight: 600,
    color: spTheme.bodyText,
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    letterSpacing: '0.5px'
  };

  const subtitleStyle: React.CSSProperties = {
    fontSize: 13,
    color: spTheme.bodySubtext,
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif'
  };

  return (
    <div style={containerStyle}>
      <span style={titleStyle}>Grimoire</span>
      <span style={subtitleStyle}>Hybrid Interaction Assistant</span>
      <PrimaryButton
        text="Start"
        onClick={handleStart}
        styles={{
          root: { borderRadius: 6, padding: '0 40px', height: 44, marginTop: 8, fontSize: 15, fontWeight: 600 }
        }}
      />
    </div>
  );
};
