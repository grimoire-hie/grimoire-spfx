/**
 * LogSidebar
 * Theme-aware logs pane shown on the left in desktop mode and as a drawer on mobile.
 */

import * as React from 'react';
import { IconButton, Pivot, PivotItem } from '@fluentui/react';
import { shallow } from 'zustand/shallow';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { ILogEntry } from '../../services/logging/LogTypes';
import type { LogCategory } from '../../services/logging/LogTypes';
import type {
  ContextType,
  HieExposureMode,
  IHieArtifactRecord,
  IHieEvent
} from '../../services/hie/HIETypes';
import {
  findLatestMcpExecutionTrace,
  getLogEntryToggleLabel,
  hasLogEntryDetails,
  parseMcpExecutionTrace,
  type IParsedMcpExecutionTrace
} from './logEntryHelpers';
import { hybridInteractionEngine } from '../../services/hie/HybridInteractionEngine';
import {
  buildStateCompassCards,
  buildHieTargetContextSummary,
  buildSituationPills,
  describeHieArtifact,
  describeHieEvent,
  formatHieContextPreview,
  getArtifactLinkageLabel,
  getDefaultExpandedCompassCard,
  shortTurnId,
  type CompassCardKey
} from './hieInspectorHelpers';

export interface ILogSidebarProps {
  width?: number;
}

const SHAREPOINT_BLUE = '#0f6cbd';

const levelColors: Record<string, string> = {
  info: '#1d4ed8',
  warning: '#b45309',
  error: '#b91c1c',
  debug: '#64748b'
};

const categoryColors: Record<string, string> = {
  mcp: '#b45309',
  llm: '#047857',
  search: '#a16207',
  graph: '#0f766e',
  system: '#64748b',
  voice: '#2563eb'
};

const FILTER_TABS: Array<{ key: LogCategory | 'all'; label: string }> = [
  { key: 'all', label: 'All' },
  { key: 'mcp', label: 'MCP' },
  { key: 'llm', label: 'LLM' },
  { key: 'search', label: 'Search' },
  { key: 'graph', label: 'Graph' },
  { key: 'voice', label: 'Voice' },
  { key: 'system', label: 'System' }
];

type HieInspectorView = 'flow' | 'events' | 'llm';

function withAlpha(color: string, alpha: number): string {
  const normalizedAlpha = Math.max(0, Math.min(1, alpha));
  const shortHex = color.match(/^#([0-9a-fA-F]{3})$/);
  if (shortHex) {
    const [r, g, b] = shortHex[1].split('').map((part) => Number.parseInt(`${part}${part}`, 16));
    return `rgba(${r}, ${g}, ${b}, ${normalizedAlpha})`;
  }

  const longHex = color.match(/^#([0-9a-fA-F]{6})$/);
  if (longHex) {
    const raw = longHex[1];
    const r = Number.parseInt(raw.slice(0, 2), 16);
    const g = Number.parseInt(raw.slice(2, 4), 16);
    const b = Number.parseInt(raw.slice(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${normalizedAlpha})`;
  }

  return color;
}

interface ILogEntryRowProps {
  entry: ILogEntry;
  textColor: string;
  subtleTextColor: string;
  borderColor: string;
  entryBackground: string;
  entryShadow: string;
  detailBackground: string;
  detailBorderColor: string;
}

interface IExecutionTracePanelProps {
  trace: IParsedMcpExecutionTrace;
  title: string;
  textColor: string;
  subtleTextColor: string;
  borderColor: string;
  background: string;
  compact?: boolean;
}

function formatRelativeTime(timestampMs: number): string {
  return new Date(timestampMs).toLocaleTimeString([], {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });
}

function getExposureModeColor(mode: HieExposureMode): string {
  switch (mode) {
    case 'response-triggering':
      return '#b91c1c';
    case 'silent-context':
      return '#0f766e';
    default:
      return '#64748b';
  }
}

function getArtifactAccent(artifact: IHieArtifactRecord): string {
  switch (artifact.artifactKind) {
    case 'summary':
      return '#2563eb';
    case 'preview':
      return '#0f766e';
    case 'lookup':
      return '#0f6cbd';
    case 'error':
      return '#b91c1c';
    case 'block':
      return SHAREPOINT_BLUE;
    case 'recap':
      return '#2563eb';
    case 'share':
      return '#b45309';
    case 'form':
      return '#7c3aed';
    default:
      return '#64748b';
  }
}

function getContextTypeColor(type: ContextType): string {
  switch (type) {
    case 'interaction':
      return '#b91c1c';
    case 'flow':
      return '#a16207';
    default:
      return '#0f766e';
  }
}

const ExecutionTracePanel: React.FC<IExecutionTracePanelProps> = ({
  trace,
  title,
  textColor,
  subtleTextColor,
  borderColor,
  background,
  compact
}) => {
  const labelStyle: React.CSSProperties = {
    color: subtleTextColor,
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: 0.25,
    textTransform: 'uppercase'
  };
  const valueStyle: React.CSSProperties = {
    color: textColor,
    fontSize: 11,
    lineHeight: 1.45,
    wordBreak: 'break-word'
  };
  const codeStyle: React.CSSProperties = {
    color: subtleTextColor,
    fontSize: 10,
    lineHeight: 1.45,
    wordBreak: 'break-word',
    fontFamily: 'Consolas, "SFMono-Regular", Menlo, monospace'
  };
  const sectionStyle: React.CSSProperties = {
    padding: compact ? '8px 10px' : '10px 12px',
    borderRadius: 10,
    border: `1px solid ${borderColor}`,
    background,
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
    minWidth: 0
  };

  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        gap: compact ? 8 : 10
      }}
    >
      <div style={{ color: textColor, fontSize: 11, fontWeight: 700 }}>
        {title}
      </div>
      <div style={sectionStyle}>
        <div style={labelStyle}>Run</div>
        <div style={valueStyle}>{trace.toolLabel}</div>
        {trace.serverLabel && trace.serverLabel !== trace.toolLabel && (
          <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
            Server: {trace.serverLabel}
          </div>
        )}
        {trace.targetLabel && (
          <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
            Target: {trace.targetLabel}
          </div>
        )}
        {trace.targetSourceLabel && (
          <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
            Source: {trace.targetSourceLabel}
          </div>
        )}
        {trace.requiredLabel && (
          <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
            Required: {trace.requiredLabel}
          </div>
        )}
        {trace.resultLabel && (
          <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
            Result: {trace.resultLabel}
          </div>
        )}
      </div>

      {(trace.rawArgsLabel || trace.normalizedArgsLabel || trace.resolvedArgsLabel) && (
        <div style={sectionStyle}>
          <div style={labelStyle}>Arguments</div>
          {trace.rawArgsLabel && <div style={codeStyle}>Raw: {trace.rawArgsLabel}</div>}
          {trace.normalizedArgsLabel && <div style={codeStyle}>Normalized: {trace.normalizedArgsLabel}</div>}
          {trace.resolvedArgsLabel && <div style={codeStyle}>Resolved: {trace.resolvedArgsLabel}</div>}
        </div>
      )}

      {(trace.recoveryLabel || trace.unwrapLabel) && (
        <div style={sectionStyle}>
          <div style={labelStyle}>Processing</div>
          {trace.recoveryLabel && (
            <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
              Recovery: {trace.recoveryLabel}
            </div>
          )}
          {trace.unwrapLabel && (
            <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.45 }}>
              Unwrapped via: {trace.unwrapLabel}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

const LogEntryRow: React.FC<ILogEntryRowProps> = ({
  entry,
  textColor,
  subtleTextColor,
  borderColor,
  entryBackground,
  entryShadow,
  detailBackground,
  detailBorderColor
}) => {
  const [expanded, setExpanded] = React.useState(false);
  const isExpandable = hasLogEntryDetails(entry);
  const timeStr = entry.timestamp.toLocaleTimeString([], {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });
  const categoryColor = categoryColors[entry.category] || subtleTextColor;
  const levelColor = levelColors[entry.level] || textColor;
  const mcpTrace = React.useMemo(() => parseMcpExecutionTrace(entry), [entry]);

  const containerStyle: React.CSSProperties = {
    width: '100%',
    padding: 0,
    border: `1px solid ${borderColor}`,
    borderRadius: 14,
    background: entryBackground,
    boxShadow: entryShadow,
    textAlign: 'left',
    cursor: isExpandable ? 'pointer' : 'default'
  };

  const content = (
    <>
      <div style={{ padding: '12px 14px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'flex-start' }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, flex: 1, minWidth: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
              <span
                style={{
                  color: categoryColor,
                  background: withAlpha(categoryColor, 0.12),
                  border: `1px solid ${withAlpha(categoryColor, 0.2)}`,
                  borderRadius: 999,
                  padding: '2px 8px',
                  fontSize: 10,
                  fontWeight: 700,
                  letterSpacing: 0.4,
                  textTransform: 'uppercase',
                  flexShrink: 0
                }}
              >
                {entry.category}
              </span>
              <span style={{ color: subtleTextColor, fontSize: 11 }}>
                {timeStr}
              </span>
              {entry.durationMs !== undefined && (
                <span style={{ color: subtleTextColor, fontSize: 11 }}>
                  {entry.durationMs}ms
                </span>
              )}
            </div>
            <span style={{ color: levelColor, fontSize: 12, lineHeight: 1.45, wordBreak: 'break-word' }}>
              {entry.message}
            </span>
          </div>
          {isExpandable && (
            <span
              aria-hidden={true}
              style={{
                color: subtleTextColor,
                fontSize: 11,
                lineHeight: '20px',
                flexShrink: 0
              }}
            >
              {getLogEntryToggleLabel(expanded)}
            </span>
          )}
        </div>

        {expanded && isExpandable && (
          <>
            {mcpTrace && (
              <div
                style={{
                  marginTop: 10,
                  padding: 10,
                  borderRadius: 10,
                  background: detailBackground,
                  border: `1px solid ${detailBorderColor}`
                }}
              >
                <ExecutionTracePanel
                  trace={mcpTrace}
                  title="Execution Trace"
                  textColor={textColor}
                  subtleTextColor={subtleTextColor}
                  borderColor={detailBorderColor}
                  background={entryBackground}
                  compact={true}
                />
              </div>
            )}
            <pre
              style={{
                margin: '10px 0 0',
                padding: 10,
                borderRadius: 10,
                background: detailBackground,
                border: `1px solid ${detailBorderColor}`,
                color: textColor,
                fontSize: 11,
                lineHeight: 1.45,
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-word',
                maxHeight: 420,
                overflowY: 'auto'
              }}
            >
              {entry.detail}
            </pre>
          </>
        )}
      </div>
    </>
  );

  if (!isExpandable) {
    return (
      <div style={containerStyle}>
        {content}
      </div>
    );
  }

  return (
    <button
      type="button"
      style={containerStyle}
      onClick={() => setExpanded(!expanded)}
    >
      {content}
    </button>
  );
};

export const LogSidebar: React.FC<ILogSidebarProps> = ({ width }) => {
  const {
    clearLogEntries,
    logEntries,
    logFilter,
    setLogFilter,
    setLogSidebarOpen,
    spThemeColors
  } = useGrimoireStore((s) => ({
    clearLogEntries: s.clearLogEntries,
    logEntries: s.logEntries,
    logFilter: s.logFilter,
    setLogFilter: s.setLogFilter,
    setLogSidebarOpen: s.setLogSidebarOpen,
    spThemeColors: s.spThemeColors
  }), shallow);
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const hasAutoSelectedCompassRef = React.useRef(false);
  const [hieInspectorView, setHieInspectorView] = React.useState<HieInspectorView>('flow');
  const [expandedCompassKey, setExpandedCompassKey] = React.useState<CompassCardKey | undefined>(undefined);
  const [executionTraceExpanded, setExecutionTraceExpanded] = React.useState(false);

  const filteredEntries = logFilter === 'all'
    ? logEntries
    : logEntries.filter((entry) => entry.category === logFilter);
  const hieTaskContext = React.useMemo(() => hybridInteractionEngine.getCurrentTaskContext(), [logEntries]);
  const hieCurrentTurn = React.useMemo(() => hybridInteractionEngine.getCurrentTurnLineage(), [logEntries]);
  const hieRecentEvents = React.useMemo(
    () => hybridInteractionEngine.getRecentEvents().slice(-12).reverse(),
    [logEntries]
  );
  const hieArtifactMap = React.useMemo(
    () => hybridInteractionEngine.getCurrentArtifacts(),
    [logEntries]
  );
  const hieArtifacts = React.useMemo(
    () => Object.values(hieArtifactMap).sort((left, right) => right.updatedAt - left.updatedAt),
    [hieArtifactMap]
  );
  const hieContextHistory = React.useMemo(
    () => hybridInteractionEngine.getContextHistory().slice(-6).reverse(),
    [logEntries]
  );
  const hieSnapshot = React.useMemo(
    () => hybridInteractionEngine.getVisualStateSnapshot(),
    [logEntries]
  );
  const hieProjectedSummary = React.useMemo(
    () => hybridInteractionEngine.getProjectedCurrentStateSummary(),
    [logEntries]
  );
  const hieLatestThreadEvent = React.useMemo(
    () => hieRecentEvents.find((entry) => entry.eventName.startsWith('thread.')),
    [hieRecentEvents]
  );
  const hieTargetContextSummary = React.useMemo(
    () => buildHieTargetContextSummary(hieTaskContext, hieArtifactMap),
    [hieArtifactMap, hieTaskContext]
  );
  const latestMcpTrace = React.useMemo(
    () => findLatestMcpExecutionTrace(logEntries),
    [logEntries]
  );
  const compassCards = React.useMemo(
    () => buildStateCompassCards(hieCurrentTurn, hieLatestThreadEvent, hieTaskContext, hieArtifactMap, hieSnapshot),
    [hieArtifactMap, hieCurrentTurn, hieLatestThreadEvent, hieSnapshot, hieTaskContext]
  );
  const hieFlowState = React.useMemo(
    () => hybridInteractionEngine.getActiveFlowState(),
    [logEntries]
  );
  const hieVerbosity = React.useMemo(
    () => hybridInteractionEngine.getVerbosityLevel(),
    [logEntries]
  );
  const hieExpressionTrigger = React.useMemo(
    () => hybridInteractionEngine.getLastExpressionTrigger(),
    [logEntries]
  );
  const situationPills = React.useMemo(
    () => buildSituationPills({
      flowState: hieFlowState,
      verbosity: hieVerbosity,
      expressionTrigger: hieExpressionTrigger,
      targetSummary: hieTargetContextSummary
    }),
    [hieExpressionTrigger, hieFlowState, hieTargetContextSummary, hieVerbosity]
  );
  const hieArtifactCount = hieArtifacts.length;

  React.useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [filteredEntries.length]);

  React.useEffect(() => {
    if (hieInspectorView !== 'flow') {
      return;
    }

    if (expandedCompassKey && compassCards.some((card) => card.key === expandedCompassKey && !card.isEmpty)) {
      hasAutoSelectedCompassRef.current = true;
      return;
    }

    if (hasAutoSelectedCompassRef.current) {
      return;
    }

    const defaultKey = getDefaultExpandedCompassCard(compassCards);
    if (defaultKey) {
      setExpandedCompassKey(defaultKey);
      hasAutoSelectedCompassRef.current = true;
    }
  }, [expandedCompassKey, compassCards, hieInspectorView]);

  const borderColor = spThemeColors.isDark
    ? withAlpha(spThemeColors.cardBorder, 0.95)
    : '#d0d7de';
  const textColor = spThemeColors.bodyText;
  const subtleTextColor = spThemeColors.bodySubtext;
  const headerBackground = spThemeColors.isDark
    ? withAlpha(SHAREPOINT_BLUE, 0.18)
    : withAlpha(SHAREPOINT_BLUE, 0.05);
  const entryBackground = spThemeColors.cardBackground;
  const entryShadow = spThemeColors.isDark
    ? '0 1px 2px rgba(0, 0, 0, 0.28)'
    : '0 1px 2px rgba(15, 23, 42, 0.08)';
  const detailBackground = spThemeColors.isDark
    ? withAlpha(spThemeColors.bodyBackground, 0.82)
    : withAlpha(SHAREPOINT_BLUE, 0.03);
  const detailBorderColor = spThemeColors.isDark
    ? withAlpha(spThemeColors.cardBorder, 0.85)
    : withAlpha(SHAREPOINT_BLUE, 0.2);
  const controlButtonStyles = {
    root: {
      color: subtleTextColor,
      width: 30,
      height: 30,
      backgroundColor: spThemeColors.bodyBackground,
      border: `1px solid ${borderColor}`,
      borderRadius: 999
    },
    rootHovered: {
      color: textColor,
      backgroundColor: withAlpha(SHAREPOINT_BLUE, spThemeColors.isDark ? 0.2 : 0.08),
      border: `1px solid ${withAlpha(SHAREPOINT_BLUE, spThemeColors.isDark ? 0.6 : 0.22)}`
    }
  };
  const inspectorSectionStyle: React.CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    gap: 8,
    padding: 10,
    borderRadius: 12,
    border: `1px solid ${detailBorderColor}`,
    background: withAlpha(spThemeColors.bodyBackground, spThemeColors.isDark ? 0.36 : 0.65)
  };
  const inspectorMiniCardStyle: React.CSSProperties = {
    padding: '8px 10px',
    borderRadius: 12,
    border: `1px solid ${detailBorderColor}`,
    background: entryBackground,
    minWidth: 0
  };
  const inspectorToggleStyle = (selected: boolean): React.CSSProperties => ({
    border: `1px solid ${selected ? withAlpha(SHAREPOINT_BLUE, 0.55) : detailBorderColor}`,
    background: selected
      ? withAlpha(SHAREPOINT_BLUE, spThemeColors.isDark ? 0.22 : 0.12)
      : withAlpha(spThemeColors.bodyBackground, spThemeColors.isDark ? 0.28 : 0.8),
    color: selected ? SHAREPOINT_BLUE : subtleTextColor,
    borderRadius: 999,
    padding: '4px 10px',
    fontSize: 10,
    fontWeight: 700,
    letterSpacing: 0.3,
    cursor: 'pointer'
  });
  const compassAccent: Record<CompassCardKey, string> = {
    thread: '#64748b',
    task: SHAREPOINT_BLUE,
    content: '#0f766e',
    artifacts: '#b45309'
  };

  return (
    <aside
      style={{
        width: width ?? 420,
        height: '100%',
        background: `linear-gradient(180deg, ${spThemeColors.cardBackground} 0%, ${headerBackground} 100%)`,
        borderRight: `1px solid ${borderColor}`,
        display: 'flex',
        flexDirection: 'column',
        overflow: 'hidden',
        flexShrink: 0
      }}
    >
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '12px 14px',
          borderBottom: `1px solid ${borderColor}`,
          background: headerBackground,
          gap: 12
        }}
      >
        <div style={{ minWidth: 0 }}>
          <div style={{ color: textColor, fontSize: 15, fontWeight: 700 }}>
            Logs
          </div>
          <div style={{ color: subtleTextColor, fontSize: 12, marginTop: 2 }}>
            {filteredEntries.length} message{filteredEntries.length === 1 ? '' : 's'}
          </div>
        </div>
        <div style={{ display: 'flex', gap: 4, flexShrink: 0 }}>
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            ariaLabel="Clear log messages"
            title="Clear log messages"
            onClick={clearLogEntries}
            styles={controlButtonStyles}
          />
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close logs"
            title="Close logs"
            onClick={() => setLogSidebarOpen(false)}
            styles={controlButtonStyles}
          />
        </div>
      </div>

      <div
        style={{
          padding: '6px 10px 0',
          borderBottom: `1px solid ${borderColor}`,
          background: withAlpha(SHAREPOINT_BLUE, spThemeColors.isDark ? 0.1 : 0.02)
        }}
      >
        <div
          style={{
            margin: '4px 2px 8px',
            padding: 12,
            borderRadius: 14,
            border: `1px solid ${detailBorderColor}`,
            background: entryBackground,
            boxShadow: entryShadow,
            display: 'flex',
            flexDirection: 'column',
            gap: 8
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'center' }}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <span style={{ color: textColor, fontSize: 12, fontWeight: 700, letterSpacing: 0.3 }}>
                HIE Inspector
              </span>
              <span style={{ color: subtleTextColor, fontSize: 11 }}>
                {hieArtifactCount} output{hieArtifactCount === 1 ? '' : 's'} · {hieSnapshot.blocks.length} visible block{hieSnapshot.blocks.length === 1 ? '' : 's'}
              </span>
            </div>
            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
              {([
                ['flow', 'Flow'],
                ['events', 'Events'],
                ['llm', 'LLM']
              ] as Array<[HieInspectorView, string]>).map(([viewKey, label]) => (
                <button
                  key={viewKey}
                  type="button"
                  onClick={() => setHieInspectorView(viewKey)}
                  style={inspectorToggleStyle(hieInspectorView === viewKey)}
                >
                  {label}
                </button>
              ))}
            </div>
          </div>
          {hieInspectorView === 'flow' && (
            <>
              {/* State Compass — 2×2 grid */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                {compassCards.map((card) => {
                  const accent = compassAccent[card.key];
                  const isExpanded = expandedCompassKey === card.key && !card.isEmpty;
                  return (
                    <button
                      key={card.key}
                      type="button"
                      onClick={() => {
                        if (card.isEmpty) return;
                        setExpandedCompassKey(isExpanded ? undefined : card.key);
                      }}
                      aria-label={`${card.label}: ${card.value}`}
                      style={{
                        ...inspectorMiniCardStyle,
                        padding: '8px 10px',
                        cursor: card.isEmpty ? 'default' : 'pointer',
                        opacity: card.isEmpty ? 0.6 : 1,
                        borderColor: isExpanded ? withAlpha(accent, 0.5) : detailBorderColor,
                        borderLeft: `3px solid ${withAlpha(accent, card.isEmpty ? 0.25 : 0.85)}`,
                        background: isExpanded
                          ? withAlpha(accent, spThemeColors.isDark ? 0.1 : 0.04)
                          : entryBackground,
                        textAlign: 'left'
                      }}
                    >
                      <div style={{ color: accent, fontSize: 9, fontWeight: 700, letterSpacing: 0.4, textTransform: 'uppercase' }}>
                        {card.label}
                      </div>
                      <div style={{ color: textColor, fontSize: 11, lineHeight: 1.35, marginTop: 3, wordBreak: 'break-word' }}>
                        {card.value}
                      </div>
                    </button>
                  );
                })}
              </div>

              {/* Expanded detail panel for selected compass card */}
              {expandedCompassKey && (() => {
                const card = compassCards.find((c) => c.key === expandedCompassKey);
                if (!card || card.isEmpty) return undefined;
                const accent = compassAccent[card.key];

                // Content detail: show visible blocks grid
                if (card.key === 'content' && hieSnapshot.blocks.length > 0) {
                  return (
                    <div style={{
                      ...inspectorSectionStyle,
                      borderColor: withAlpha(accent, 0.35),
                      background: withAlpha(accent, spThemeColors.isDark ? 0.06 : 0.02)
                    }}>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                        {hieSnapshot.blocks.slice(0, 4).map((block) => (
                          <div
                            key={block.blockId}
                            style={{
                              ...inspectorMiniCardStyle,
                              padding: '6px 8px',
                              minWidth: 110,
                              maxWidth: '100%'
                            }}
                          >
                            <div style={{ color: subtleTextColor, fontSize: 9, textTransform: 'uppercase' }}>
                              {block.blockType}
                            </div>
                            <div style={{ color: textColor, fontSize: 11, lineHeight: 1.4, marginTop: 2, wordBreak: 'break-word' }}>
                              {block.title}
                            </div>
                            <div style={{ color: subtleTextColor, fontSize: 10, marginTop: 2 }}>
                              {block.itemCount} item{block.itemCount === 1 ? '' : 's'}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                }

                // Artifacts detail: show artifact chain
                if (card.key === 'artifacts' && hieArtifacts.length > 0) {
                  return (
                    <div style={{
                      ...inspectorSectionStyle,
                      borderColor: withAlpha(accent, 0.35),
                      background: withAlpha(accent, spThemeColors.isDark ? 0.06 : 0.02)
                    }}>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                        {hieArtifacts.slice(0, 4).map((artifact, index) => {
                          const artAccent = getArtifactAccent(artifact);
                          return (
                            <div key={artifact.artifactId} style={{ display: 'flex', gap: 8, alignItems: 'stretch' }}>
                              <div style={{ width: 10, display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                                <div style={{ width: 6, height: 6, borderRadius: 999, marginTop: 6, background: artAccent }} />
                                {index < Math.min(hieArtifacts.length, 4) - 1 && (
                                  <div style={{ width: 1, flex: 1, marginTop: 4, background: withAlpha(artAccent, 0.28) }} />
                                )}
                              </div>
                              <div style={{ ...inspectorMiniCardStyle, flex: 1, borderColor: withAlpha(artAccent, 0.28) }}>
                                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'center' }}>
                                  <span style={{ color: artAccent, fontSize: 10, fontWeight: 700, letterSpacing: 0.3, textTransform: 'uppercase' }}>
                                    {getArtifactLinkageLabel(artifact)}
                                  </span>
                                  <span style={{ color: subtleTextColor, fontSize: 10 }}>
                                    {artifact.status}
                                  </span>
                                </div>
                                <div style={{ color: textColor, fontSize: 11, lineHeight: 1.45, marginTop: 4 }}>
                                  {describeHieArtifact(artifact)}
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                }

                // Thread / Task detail: text detail + meta
                return (
                  <div style={{
                    ...inspectorSectionStyle,
                    borderColor: withAlpha(accent, 0.35),
                    background: withAlpha(accent, spThemeColors.isDark ? 0.06 : 0.02)
                  }}>
                    {card.detail && (
                      <div style={{ color: textColor, fontSize: 11, lineHeight: 1.45, wordBreak: 'break-word' }}>
                        {card.detail}
                      </div>
                    )}
                    {card.meta && (
                      <div style={{ color: subtleTextColor, fontSize: 10, lineHeight: 1.4, wordBreak: 'break-word' }}>
                        {card.meta}
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* Situation Bar */}
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5 }}>
                {situationPills.map((pill) => (
                  <span
                    key={pill.key}
                    style={{
                      color: pill.accent,
                      background: withAlpha(pill.accent, spThemeColors.isDark ? 0.14 : 0.08),
                      border: `1px solid ${withAlpha(pill.accent, 0.25)}`,
                      borderRadius: 999,
                      padding: '3px 8px',
                      fontSize: 10,
                      fontWeight: 600
                    }}
                  >
                    {pill.label}
                  </span>
                ))}
              </div>

              {/* Latest Execution — collapsible, default collapsed */}
              {latestMcpTrace && (
                <div style={inspectorSectionStyle}>
                  <button
                    type="button"
                    onClick={() => setExecutionTraceExpanded(!executionTraceExpanded)}
                    style={{
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                      gap: 8,
                      background: 'none',
                      border: 'none',
                      padding: 0,
                      cursor: 'pointer',
                      width: '100%',
                      textAlign: 'left'
                    }}
                  >
                    <span style={{ color: textColor, fontSize: 11, fontWeight: 700 }}>
                      {executionTraceExpanded ? '\u25BC' : '\u25B6'}{' '}
                      {latestMcpTrace.toolLabel}{latestMcpTrace.serverLabel ? ` via ${latestMcpTrace.serverLabel}` : ''}
                    </span>
                  </button>
                  {executionTraceExpanded && (
                    <ExecutionTracePanel
                      trace={latestMcpTrace}
                      title="Latest Execution Run"
                      textColor={textColor}
                      subtleTextColor={subtleTextColor}
                      borderColor={detailBorderColor}
                      background={entryBackground}
                    />
                  )}
                </div>
              )}
            </>
          )}
          {hieInspectorView === 'events' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <div style={inspectorSectionStyle}>
                <span style={{ color: subtleTextColor, fontSize: 11, lineHeight: 1.45 }}>
                  Raw HIE events. The Flow tab condenses them into the current path.
                </span>
              </div>
              {hieRecentEvents.length === 0 ? (
                <div style={inspectorSectionStyle}>
                  <span style={{ color: subtleTextColor, fontSize: 11 }}>
                    No HIE events recorded yet.
                  </span>
                </div>
              ) : (
                hieRecentEvents.map((event: IHieEvent) => {
                  const exposureColor = getExposureModeColor(event.exposurePolicy.mode);
                  return (
                    <div
                      key={event.eventId}
                      style={{
                        ...inspectorMiniCardStyle,
                        borderLeft: `3px solid ${exposureColor}`
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'flex-start' }}>
                        <div style={{ minWidth: 0 }}>
                          <div style={{ color: textColor, fontSize: 11, fontWeight: 700, wordBreak: 'break-word' }}>
                            {event.eventName}
                          </div>
                          <div style={{ color: subtleTextColor, fontSize: 10, marginTop: 2 }}>
                            {describeHieEvent(event)}
                          </div>
                          <div style={{ color: subtleTextColor, fontSize: 10, marginTop: 4 }}>
                            {[
                              `corr ${shortTurnId(event.correlationId)}`,
                              event.rootTurnId ? `root ${shortTurnId(event.rootTurnId)}` : undefined
                            ].filter((part): part is string => !!part).join(' · ')}
                          </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4, flexShrink: 0 }}>
                          <span
                            style={{
                              color: exposureColor,
                              background: withAlpha(exposureColor, 0.12),
                              border: `1px solid ${withAlpha(exposureColor, 0.2)}`,
                              borderRadius: 999,
                              padding: '2px 8px',
                              fontSize: 10,
                              fontWeight: 700
                            }}
                          >
                            {event.exposurePolicy.mode}
                          </span>
                          <span style={{ color: subtleTextColor, fontSize: 10 }}>
                            {event.turnId ? shortTurnId(event.turnId) : formatRelativeTime(event.timestamp)}
                          </span>
                        </div>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          )}
          {hieInspectorView === 'llm' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <div style={inspectorSectionStyle}>
                <span style={{ color: subtleTextColor, fontSize: 11, lineHeight: 1.45 }}>
                  Prompt-ready HIE context actually sent toward the model.
                </span>
              </div>
              {hieProjectedSummary && (
                <div style={inspectorSectionStyle}>
                  <div style={{ color: textColor, fontSize: 11, fontWeight: 700 }}>
                    Current LLM-Ready Summary
                  </div>
                  <span style={{ color: subtleTextColor, fontSize: 11, lineHeight: 1.45 }}>
                    {hieProjectedSummary}
                  </span>
                </div>
              )}
              {hieContextHistory.length === 0 ? (
                <div style={inspectorSectionStyle}>
                  <span style={{ color: subtleTextColor, fontSize: 11 }}>
                    No prompt context injected yet.
                  </span>
                </div>
              ) : (
                hieContextHistory.map((message) => {
                  const contextColor = getContextTypeColor(message.contextType);
                  return (
                    <div
                      key={`${message.sentAt}-${message.contextType}`}
                      style={{
                        ...inspectorMiniCardStyle,
                        borderLeft: `3px solid ${contextColor}`
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'center' }}>
                        <span
                          style={{
                            color: contextColor,
                            fontSize: 10,
                            fontWeight: 700,
                            letterSpacing: 0.3,
                            textTransform: 'uppercase'
                          }}
                        >
                          {message.contextType}
                        </span>
                        <span style={{ color: subtleTextColor, fontSize: 10 }}>
                          {formatRelativeTime(message.sentAt)}
                        </span>
                      </div>
                      <div style={{ color: textColor, fontSize: 11, lineHeight: 1.45, marginTop: 4 }}>
                        {formatHieContextPreview(message)}
                      </div>
                      {message.blockIds.length > 0 && (
                        <div style={{ color: subtleTextColor, fontSize: 10, marginTop: 4 }}>
                          {message.blockIds.length} related block{message.blockIds.length === 1 ? '' : 's'}
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          )}
        </div>
        <Pivot
          selectedKey={logFilter}
          onLinkClick={(item) => {
            if (item?.props.itemKey) {
              setLogFilter(item.props.itemKey as LogCategory | 'all');
            }
          }}
          styles={{
            root: { display: 'flex' },
            link: {
              color: subtleTextColor,
              fontSize: 11,
              padding: '6px 10px',
              height: 32,
              lineHeight: '32px',
              borderRadius: 999,
              marginRight: 4
            },
            linkIsSelected: {
              color: SHAREPOINT_BLUE,
              backgroundColor: withAlpha(SHAREPOINT_BLUE, spThemeColors.isDark ? 0.2 : 0.12)
            }
          }}
        >
          {FILTER_TABS.map((tab) => (
            <PivotItem key={tab.key} itemKey={tab.key} headerText={tab.label} />
          ))}
        </Pivot>
      </div>

      <div
        ref={scrollRef}
        style={{
          flex: 1,
          overflowY: 'auto',
          padding: 12,
          display: 'flex',
          flexDirection: 'column',
          gap: 10,
          minHeight: 0
        }}
      >
        {filteredEntries.length === 0 ? (
          <div
            style={{
              margin: 'auto 0',
              padding: 20,
              borderRadius: 14,
              border: `1px solid ${detailBorderColor}`,
              background: entryBackground,
              color: subtleTextColor,
              textAlign: 'center',
              fontSize: 13,
              lineHeight: 1.5
            }}
          >
            No log messages in this view yet.
          </div>
        ) : (
          filteredEntries.map((entry) => (
            <LogEntryRow
              key={entry.id}
              borderColor={borderColor}
              entryBackground={entryBackground}
              entryShadow={entryShadow}
              detailBackground={detailBackground}
              detailBorderColor={detailBorderColor}
              entry={entry}
              subtleTextColor={subtleTextColor}
              textColor={textColor}
            />
          ))
        )}
      </div>
    </aside>
  );
};
