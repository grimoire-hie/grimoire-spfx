/**
 * ChartBlock
 * Renders simple bar, pie, or line charts using pure CSS/HTML.
 * No charting library dependency — keeps the SPFx bundle lean.
 */

import * as React from 'react';
import type { IChartData } from '../../../models/IBlock';
import { injectHoverStyles, renderHoverActions, ACTIONS_LC } from './shared/HoverActions';

const CHART_COLORS = [
  '#6cb4ee', '#f0a030', '#40c040', '#e06060', '#b080e0',
  '#40c0c0', '#e0e040', '#f07030', '#8090c0', '#c060a0'
];

// ─── Bar Chart ────────────────────────────────────────────────

const BarChart: React.FC<{ labels: string[]; values: number[] }> = ({ labels, values }) => {
  const max = Math.max(...values, 1);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
      {labels.map((label, idx) => {
        const pct = Math.round((values[idx] / max) * 100);
        const color = CHART_COLORS[idx % CHART_COLORS.length];
        return (
          <div key={idx}>
            <div style={{
              display: 'flex', justifyContent: 'space-between',
              fontSize: 11, color: '#605e5c', marginBottom: 2
            }}>
              <span>{label}</span>
              <span style={{ color: '#a19f9d' }}>{values[idx]}</span>
            </div>
            <div style={{
              height: 14, borderRadius: 7,
              background: 'rgba(0,0,0,0.06)', overflow: 'hidden'
            }}>
              <div style={{
                width: `${pct}%`, height: '100%', borderRadius: 7,
                background: color, transition: 'width 0.4s ease'
              }} />
            </div>
          </div>
        );
      })}
    </div>
  );
};

// ─── Pie Chart ────────────────────────────────────────────────

const PieChart: React.FC<{ labels: string[]; values: number[] }> = ({ labels, values }) => {
  const total = values.reduce((sum, v) => sum + v, 0) || 1;
  let cumulative = 0;

  // Build conic-gradient stops
  const stops: string[] = [];
  values.forEach((v, idx) => {
    const color = CHART_COLORS[idx % CHART_COLORS.length];
    const start = cumulative;
    cumulative += (v / total) * 360;
    stops.push(`${color} ${start}deg ${cumulative}deg`);
  });

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
      <div style={{
        width: 80, height: 80, borderRadius: '50%', flexShrink: 0,
        background: `conic-gradient(${stops.join(', ')})`
      }} />
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        {labels.map((label, idx) => (
          <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11 }}>
            <div style={{
              width: 8, height: 8, borderRadius: 2, flexShrink: 0,
              background: CHART_COLORS[idx % CHART_COLORS.length]
            }} />
            <span style={{ color: '#605e5c' }}>{label}</span>
            <span style={{ color: '#c8c6c4', marginLeft: 'auto' }}>
              {Math.round((values[idx] / total) * 100)}%
            </span>
          </div>
        ))}
      </div>
    </div>
  );
};

// ─── Line Chart (sparkline) ───────────────────────────────────

const LineChart: React.FC<{ labels: string[]; values: number[] }> = ({ labels, values }) => {
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const range = max - min || 1;
  const width = 280;
  const height = 60;
  const padding = 4;

  const points = values.map((v, idx) => {
    const x = padding + (idx / Math.max(values.length - 1, 1)) * (width - padding * 2);
    const y = height - padding - ((v - min) / range) * (height - padding * 2);
    return `${x},${y}`;
  }).join(' ');

  return (
    <div>
      <svg viewBox={`0 0 ${width} ${height}`} style={{ width: '100%', maxWidth: width, height }}>
        <polyline
          points={points}
          fill="none"
          stroke="#0064b4"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
        {values.map((v, idx) => {
          const x = padding + (idx / Math.max(values.length - 1, 1)) * (width - padding * 2);
          const y = height - padding - ((v - min) / range) * (height - padding * 2);
          return <circle key={idx} cx={x} cy={y} r="3" fill="#0064b4" />;
        })}
      </svg>
      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 10, color: '#c8c6c4', marginTop: 2 }}>
        {labels.length <= 8 ? labels.map((l, idx) => <span key={idx}>{l}</span>) : (
          <>
            <span>{labels[0]}</span>
            <span>{labels[labels.length - 1]}</span>
          </>
        )}
      </div>
    </div>
  );
};

// ─── Main Component ──────────────────────────────────────────

export const ChartBlock: React.FC<{ data: IChartData; blockId?: string }> = ({ data, blockId }) => {
  const { chartType, labels, values } = data;
  React.useEffect(() => { injectHoverStyles('ch'); }, []);

  if (!labels || !values || labels.length === 0) {
    return (
      <div style={{ fontSize: 12, color: '#a19f9d', fontStyle: 'italic' }}>
        No chart data.
      </div>
    );
  }

  switch (chartType) {
    case 'pie':
      return (
        <div className="grim-ch-row">
          <PieChart labels={labels} values={values} />
          <div style={{ marginTop: 8 }}>
            {renderHoverActions(ACTIONS_LC, blockId, 'chart', { title: data.title, chartType, labels, values }, 'grim-ch-actions')}
          </div>
        </div>
      );
    case 'line':
      return (
        <div className="grim-ch-row">
          <LineChart labels={labels} values={values} />
          <div style={{ marginTop: 8 }}>
            {renderHoverActions(ACTIONS_LC, blockId, 'chart', { title: data.title, chartType, labels, values }, 'grim-ch-actions')}
          </div>
        </div>
      );
    case 'bar':
    default:
      return (
        <div className="grim-ch-row">
          <BarChart labels={labels} values={values} />
          <div style={{ marginTop: 8 }}>
            {renderHoverActions(ACTIONS_LC, blockId, 'chart', { title: data.title, chartType, labels, values }, 'grim-ch-actions')}
          </div>
        </div>
      );
  }
};
