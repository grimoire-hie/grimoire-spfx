/**
 * AppLayout
 * Desktop layout: Logs pane (left, resizable) | AvatarPanel | ActionPanel.
 * Mobile layout: avatar + action stacked, with the logs pane opening as a drawer.
 */

import * as React from 'react';
import { shallow } from 'zustand/shallow';
import * as strings from 'GrimoireAssistantWebPartStrings';
import { Header } from './Header';
import {
  DEFAULT_AVATAR_PANE_RATIO,
  MAIN_PANE_RESIZE_HANDLE_WIDTH,
  resolveAvatarActionPaneLayout,
  resolveAvatarPaneRatioFromPointer
} from './paneSizing';
import {
  DEFAULT_LOG_SIDEBAR_WIDTH,
  useGrimoireStore
} from '../../store/useGrimoireStore';
import { hybridInteractionEngine } from '../../services/hie/HybridInteractionEngine';
import { resolveEscapeCloseAction } from './closeBehavior';

const LazyAvatarPanel = React.lazy(async () => ({
  default: (await import(/* webpackChunkName: 'grimoire-avatar-panel' */ '../avatar/AvatarPanel')).AvatarPanel
}));
const LazyActionPanel = React.lazy(async () => ({
  default: (await import(/* webpackChunkName: 'grimoire-action-panel' */ '../action-panel/ActionPanel')).ActionPanel
}));
const LazyLogSidebar = React.lazy(async () => ({
  default: (await import(/* webpackChunkName: 'grimoire-log-sidebar' */ '../logging/LogSidebar')).LogSidebar
}));

const MOBILE_BREAKPOINT = 768;
const LOG_SIDEBAR_MIN_WIDTH = 320;
const LOG_SIDEBAR_MAX_WIDTH = 720;
const MIN_CONTENT_WIDTH = 440;
const MOBILE_DRAWER_MAX_WIDTH = 440;
const RESIZE_HANDLE_WIDTH = MAIN_PANE_RESIZE_HANDLE_WIDTH;

type ResizeTarget = 'logs-pane' | 'main-panes';

const bodyStyle: React.CSSProperties = {
  display: 'flex',
  flex: 1,
  overflow: 'hidden',
  minHeight: 0
};

interface IDeferredSurfaceProps {
  background: string;
  borderColor: string;
  textColor: string;
  label: string;
}

const DeferredSurface: React.FC<IDeferredSurfaceProps> = ({
  background,
  borderColor,
  textColor,
  label
}) => (
  <div
    style={{
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      width: '100%',
      height: '100%',
      minHeight: 0,
      padding: 20,
      boxSizing: 'border-box',
      background,
      color: textColor
    }}
  >
    <div
      style={{
        width: '100%',
        height: '100%',
        minHeight: 160,
        border: `1px dashed ${borderColor}`,
        borderRadius: 12,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontSize: 12,
        letterSpacing: 0.2,
        opacity: 0.85
      }}
    >
      {label}
    </div>
  </div>
);

function clampLogSidebarWidth(width: number, viewportWidth: number): number {
  const maxWidth = Math.min(
    LOG_SIDEBAR_MAX_WIDTH,
    Math.max(LOG_SIDEBAR_MIN_WIDTH, viewportWidth - MIN_CONTENT_WIDTH)
  );
  return Math.min(maxWidth, Math.max(LOG_SIDEBAR_MIN_WIDTH, Math.round(width)));
}

export const AppLayout: React.FC = () => {
  const {
    connectionState,
    logSidebarOpen,
    logSidebarWidth,
    mcpConnections,
    setLogSidebarOpen,
    setLogSidebarWidth,
    setUiMode,
    spThemeColors
  } = useGrimoireStore((s) => ({
    connectionState: s.connectionState,
    logSidebarOpen: s.logSidebarOpen,
    logSidebarWidth: s.logSidebarWidth,
    mcpConnections: s.mcpConnections,
    setLogSidebarOpen: s.setLogSidebarOpen,
    setLogSidebarWidth: s.setLogSidebarWidth,
    setUiMode: s.setUiMode,
    spThemeColors: s.spThemeColors
  }), shallow);

  const [viewportWidth, setViewportWidth] = React.useState(
    typeof window !== 'undefined' ? window.innerWidth : 1280
  );
  const [mainContentWidth, setMainContentWidth] = React.useState(
    typeof window !== 'undefined' ? window.innerWidth : 1280
  );
  const [isMobile, setIsMobile] = React.useState(
    typeof window !== 'undefined' ? window.innerWidth < MOBILE_BREAKPOINT : false
  );
  const [isSettingsOpen, setIsSettingsOpen] = React.useState(false);
  const [avatarPaneRatio, setAvatarPaneRatio] = React.useState(DEFAULT_AVATAR_PANE_RATIO);
  const activeResizeTargetRef = React.useRef<ResizeTarget | undefined>(undefined);
  const mainContentRef = React.useRef<HTMLDivElement>(null);

  const appBackground = spThemeColors.bodyBackground;
  const leftColumnBackground = spThemeColors.cardBackground;
  const mainPaneLayout = React.useMemo(
    () => resolveAvatarActionPaneLayout(mainContentWidth, avatarPaneRatio),
    [avatarPaneRatio, mainContentWidth]
  );
  const mobileDrawerWidth = Math.min(
    MOBILE_DRAWER_MAX_WIDTH,
    Math.max(LOG_SIDEBAR_MIN_WIDTH, viewportWidth - 32)
  );

  const rootStyle: React.CSSProperties = {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    height: '100vh',
    overflow: 'hidden',
    background: appBackground,
    color: spThemeColors.bodyText,
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif'
  };

  const avatarColumnStyle: React.CSSProperties = {
    flex: '0 0 auto',
    width: mainPaneLayout.avatarWidth,
    minWidth: 0,
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    background: leftColumnBackground
  };

  const mobileAvatarStyle: React.CSSProperties = {
    flex: '0 0 45%',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    background: leftColumnBackground,
    borderBottom: `1px solid ${spThemeColors.cardBorder}`
  };

  const mainContentStyle: React.CSSProperties = {
    display: 'flex',
    flex: 1,
    minWidth: 0,
    minHeight: 0,
    overflow: 'hidden'
  };

  const actionColumnStyle: React.CSSProperties = {
    flex: '1 1 auto',
    width: mainPaneLayout.actionWidth,
    minWidth: 0,
    overflow: 'hidden'
  };

  const resizeHandleStyle: React.CSSProperties = {
    width: RESIZE_HANDLE_WIDTH,
    flexShrink: 0,
    cursor: 'col-resize',
    display: 'flex',
    alignItems: 'stretch',
    justifyContent: 'center',
    background: spThemeColors.bodyBackground,
    touchAction: 'none'
  };

  const resizeRailStyle: React.CSSProperties = {
    width: 2,
    margin: '12px 0',
    borderRadius: 999,
    background: spThemeColors.cardBorder,
    opacity: 0.9
  };

  const mobileDrawerBackdropStyle: React.CSSProperties = {
    position: 'absolute',
    top: 48,
    right: 0,
    bottom: 0,
    left: 0,
    background: 'rgba(15, 23, 42, 0.18)',
    backdropFilter: 'blur(2px)',
    zIndex: 19
  };

  const mobileDrawerShellStyle: React.CSSProperties = {
    position: 'absolute',
    top: 48,
    bottom: 0,
    left: 0,
    width: mobileDrawerWidth,
    maxWidth: '92vw',
    zIndex: 20,
    boxShadow: '24px 0 48px rgba(15, 23, 42, 0.18)'
  };

  React.useEffect(() => {
    const handleResize = (): void => {
      const nextWidth = window.innerWidth;
      setViewportWidth(nextWidth);
      setIsMobile(nextWidth < MOBILE_BREAKPOINT);
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  React.useEffect(() => {
    if (isMobile) {
      return undefined;
    }

    const element = mainContentRef.current;
    if (!element) {
      return undefined;
    }

    const updateWidth = (): void => {
      const nextWidth = Math.round(element.getBoundingClientRect().width);
      setMainContentWidth((currentWidth) => currentWidth === nextWidth ? currentWidth : nextWidth);
    };

    updateWidth();

    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(() => updateWidth());
      observer.observe(element);
      return () => observer.disconnect();
    }

    window.addEventListener('resize', updateWidth);
    return () => window.removeEventListener('resize', updateWidth);
  }, [isMobile, logSidebarOpen]);

  React.useEffect(() => {
    hybridInteractionEngine.emitEvent({
      eventName: 'shell.logs.toggled',
      source: 'app-layout',
      surface: 'logs',
      correlationId: `shell-logs-${Date.now()}`,
      payload: { isOpen: logSidebarOpen },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });
  }, [logSidebarOpen]);

  React.useEffect(() => {
    hybridInteractionEngine.emitEvent({
      eventName: 'shell.settings.toggled',
      source: 'app-layout',
      surface: 'settings',
      correlationId: `shell-settings-${Date.now()}`,
      payload: { isOpen: isSettingsOpen },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });
  }, [isSettingsOpen]);

  React.useEffect(() => {
    hybridInteractionEngine.emitEvent({
      eventName: 'shell.app.visibility',
      source: 'app-layout',
      surface: 'app-shell',
      correlationId: `shell-app-open-${Date.now()}`,
      payload: { isOpen: true },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });
  }, []);

  React.useEffect(() => {
    if (isMobile) return;
    const clamped = clampLogSidebarWidth(logSidebarWidth || DEFAULT_LOG_SIDEBAR_WIDTH, viewportWidth);
    if (clamped !== logSidebarWidth) {
      setLogSidebarWidth(clamped);
    }
  }, [isMobile, logSidebarWidth, setLogSidebarWidth, viewportWidth]);

  React.useEffect(() => {
    const stopResizing = (): void => {
      if (!activeResizeTargetRef.current) return;
      activeResizeTargetRef.current = undefined;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    const handlePointerMove = (event: PointerEvent): void => {
      if (!activeResizeTargetRef.current || isMobile) return;

      if (activeResizeTargetRef.current === 'logs-pane') {
        const nextWidth = clampLogSidebarWidth(event.clientX, viewportWidth);
        setLogSidebarWidth(nextWidth);
        return;
      }

      const element = mainContentRef.current;
      if (!element) return;

      const rect = element.getBoundingClientRect();
      const nextRatio = resolveAvatarPaneRatioFromPointer(event.clientX, rect.left, rect.width);
      setAvatarPaneRatio((currentRatio) => (
        Math.abs(currentRatio - nextRatio) < 0.001 ? currentRatio : nextRatio
      ));
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', stopResizing);
    window.addEventListener('pointercancel', stopResizing);

    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', stopResizing);
      window.removeEventListener('pointercancel', stopResizing);
      stopResizing();
    };
  }, [isMobile, setLogSidebarWidth, viewportWidth]);

  const handleToggleLogSidebar = React.useCallback((): void => {
    setLogSidebarOpen(!logSidebarOpen);
  }, [logSidebarOpen, setLogSidebarOpen]);

  const handleClose = React.useCallback((): void => {
    hybridInteractionEngine.emitEvent({
      eventName: 'shell.app.visibility',
      source: 'app-layout',
      surface: 'app-shell',
      correlationId: `shell-app-${Date.now()}`,
      payload: { isOpen: false },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });
    setUiMode('widget');
  }, [setUiMode]);

  const handleLogPaneResizeStart = React.useCallback((event: React.PointerEvent<HTMLDivElement>): void => {
    if (isMobile) return;
    event.preventDefault();
    activeResizeTargetRef.current = 'logs-pane';
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  }, [isMobile]);

  const handleMainPaneResizeStart = React.useCallback((event: React.PointerEvent<HTMLDivElement>): void => {
    if (isMobile) return;
    event.preventDefault();
    activeResizeTargetRef.current = 'main-panes';
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  }, [isMobile]);

  React.useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        switch (resolveEscapeCloseAction({ isMobile, logSidebarOpen, isSettingsOpen })) {
          case 'close-logs':
            setLogSidebarOpen(false);
            return;
          case 'close-settings':
            setIsSettingsOpen(false);
            return;
          default:
            handleClose();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleClose, isMobile, isSettingsOpen, logSidebarOpen, setLogSidebarOpen]);

  const connectedMcpCount = mcpConnections.filter((connection) => connection.state === 'connected').length;

  if (isMobile) {
    return (
      <div style={rootStyle}>
        <Header
          connectionState={connectionState}
          mcpServerCount={connectedMcpCount}
          logSidebarOpen={logSidebarOpen}
          spThemeColors={spThemeColors}
          onToggleLogSidebar={handleToggleLogSidebar}
        />
        <div style={mobileAvatarStyle}>
          <React.Suspense
            fallback={(
              <DeferredSurface
                background={leftColumnBackground}
                borderColor={spThemeColors.cardBorder}
                textColor={spThemeColors.bodySubtext}
                label={strings.LoadingAvatarSurface}
              />
            )}
          >
            <LazyAvatarPanel
              isSettingsOpen={isSettingsOpen}
              onOpenSettings={() => setIsSettingsOpen(true)}
              onDismissSettings={() => setIsSettingsOpen(false)}
            />
          </React.Suspense>
        </div>
        <div style={{ flex: '1 1 55%', overflow: 'hidden' }}>
          <React.Suspense
            fallback={(
              <DeferredSurface
                background={spThemeColors.bodyBackground}
                borderColor={spThemeColors.cardBorder}
                textColor={spThemeColors.bodySubtext}
                label={strings.LoadingAssistantTools}
              />
            )}
          >
            <LazyActionPanel
              isSettingsOpen={isSettingsOpen}
              onCloseApp={handleClose}
              onCloseSettings={() => setIsSettingsOpen(false)}
            />
          </React.Suspense>
        </div>
        {logSidebarOpen && (
          <>
            <div
              aria-hidden={true}
              style={mobileDrawerBackdropStyle}
              onClick={() => setLogSidebarOpen(false)}
            />
            <div style={mobileDrawerShellStyle}>
              <React.Suspense
                fallback={(
                  <DeferredSurface
                    background={spThemeColors.cardBackground}
                    borderColor={spThemeColors.cardBorder}
                    textColor={spThemeColors.bodySubtext}
                    label={strings.LoadingLogs}
                  />
                )}
              >
                <LazyLogSidebar width={mobileDrawerWidth} />
              </React.Suspense>
            </div>
          </>
        )}
      </div>
    );
  }

  return (
    <div style={rootStyle}>
      <Header
        connectionState={connectionState}
        mcpServerCount={connectedMcpCount}
        logSidebarOpen={logSidebarOpen}
        spThemeColors={spThemeColors}
        onToggleLogSidebar={handleToggleLogSidebar}
      />
      <div style={bodyStyle}>
        {logSidebarOpen && (
          <>
            <React.Suspense
              fallback={(
                <div style={{ width: logSidebarWidth, minWidth: 0 }}>
                  <DeferredSurface
                    background={spThemeColors.cardBackground}
                    borderColor={spThemeColors.cardBorder}
                    textColor={spThemeColors.bodySubtext}
                    label={strings.LoadingLogs}
                  />
                </div>
              )}
            >
              <LazyLogSidebar width={logSidebarWidth} />
            </React.Suspense>
            <div
              role="separator"
              aria-orientation="vertical"
              aria-label={strings.ResizeLogsPaneLabel}
              title={strings.ResizeLogsPaneTitle}
              style={resizeHandleStyle}
              onPointerDown={handleLogPaneResizeStart}
            >
              <div style={resizeRailStyle} />
            </div>
          </>
        )}
        <div ref={mainContentRef} style={mainContentStyle}>
          <div style={avatarColumnStyle}>
            <React.Suspense
              fallback={(
                <DeferredSurface
                  background={leftColumnBackground}
                  borderColor={spThemeColors.cardBorder}
                  textColor={spThemeColors.bodySubtext}
                  label={strings.LoadingAvatarSurface}
                />
              )}
            >
              <LazyAvatarPanel
                isSettingsOpen={isSettingsOpen}
                onOpenSettings={() => setIsSettingsOpen(true)}
                onDismissSettings={() => setIsSettingsOpen(false)}
              />
            </React.Suspense>
          </div>
          <div
            role="separator"
            aria-orientation="vertical"
            aria-label={strings.ResizeAvatarActionLabel}
            title={strings.ResizeAvatarActionTitle}
            style={resizeHandleStyle}
            onPointerDown={handleMainPaneResizeStart}
          >
            <div style={resizeRailStyle} />
          </div>
          <div style={actionColumnStyle}>
            <React.Suspense
              fallback={(
                <DeferredSurface
                  background={spThemeColors.bodyBackground}
                  borderColor={spThemeColors.cardBorder}
                  textColor={spThemeColors.bodySubtext}
                  label={strings.LoadingAssistantTools}
                />
              )}
            >
              <LazyActionPanel
                isSettingsOpen={isSettingsOpen}
                onCloseApp={handleClose}
                onCloseSettings={() => setIsSettingsOpen(false)}
              />
            </React.Suspense>
          </div>
        </div>
      </div>
    </div>
  );
};
