/**
 * ProgressTrackerBlock
 * Renders a progress bar with label and percentage.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type {
  IProgressTrackerData,
  IProgressTrackerStep,
  ProgressTrackerStepStatus
} from '../../../models/IBlock';
import { injectHoverStyles, renderHoverActions, ACTIONS_LC } from './shared/HoverActions';

const barContainerStyle: React.CSSProperties = {
  width: '100%',
  height: 6,
  borderRadius: 3,
  backgroundColor: 'rgba(0, 0, 0, 0.06)',
  overflow: 'hidden',
  marginTop: 8
};

const statusRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  marginBottom: 4
};

const labelStyle: React.CSSProperties = {
  fontSize: 13,
  color: '#323130',
  display: 'flex',
  alignItems: 'center',
  gap: 6
};

const percentStyle: React.CSSProperties = {
  fontSize: 12,
  fontWeight: 600,
  color: '#605e5c'
};

const detailStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#a19f9d',
  marginTop: 6
};

const STATUS_STYLE: Record<string, { color: string; icon: string }> = {
  complete: { color: '#4caf50', icon: 'CompletedSolid' },
  error: { color: '#ef5350', icon: 'ErrorBadge' }
};
const DEFAULT_STATUS_STYLE = { color: '#0064b4', icon: 'ProgressLoopInner' };

function getStatusStyle(status: string): { color: string; icon: string } {
  return STATUS_STYLE[status] || DEFAULT_STATUS_STYLE;
}

const stepsContainerStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 6,
  marginTop: 10
};

const stepRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  gap: 8
};

const stepTextStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 2,
  minWidth: 0
};

const stepLabelStyle: React.CSSProperties = {
  fontSize: 12,
  color: '#323130',
  lineHeight: 1.3
};

const stepDetailStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#8a8886',
  lineHeight: 1.3
};

function getStepIcon(status: ProgressTrackerStepStatus): string {
  switch (status) {
    case 'complete': return 'CompletedSolid';
    case 'error': return 'StatusErrorFull';
    case 'running': return 'ProgressLoopInner';
    default: return 'CircleRing';
  }
}

function getStepColor(status: ProgressTrackerStepStatus): string {
  switch (status) {
    case 'complete': return '#4caf50';
    case 'error': return '#ef5350';
    case 'running': return '#0064b4';
    default: return '#a19f9d';
  }
}

function renderStep(step: IProgressTrackerStep): React.ReactNode {
  return (
    <div key={step.id} style={stepRowStyle}>
      <Icon
        iconName={getStepIcon(step.status)}
        styles={{ root: { fontSize: 12, color: getStepColor(step.status), marginTop: 2 } }}
      />
      <div style={stepTextStyle}>
        <div style={stepLabelStyle}>{step.label}</div>
        {step.detail && <div style={stepDetailStyle}>{step.detail}</div>}
      </div>
    </div>
  );
}

export const ProgressTrackerBlock: React.FC<{ data: IProgressTrackerData; blockId?: string }> = ({ data, blockId }) => {
  const { label, progress, status, detail, steps } = data;
  const clampedProgress = Math.min(100, Math.max(0, progress));
  React.useEffect(() => { injectHoverStyles('pt'); }, []);

  return (
    <div>
      <div style={statusRowStyle}>
        <div style={labelStyle}>
          <Icon
            iconName={getStatusStyle(status).icon}
            styles={{ root: { fontSize: 14, color: getStatusStyle(status).color } }}
          />
          {label}
        </div>
        <div className="grim-pt-row" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={percentStyle}>{clampedProgress}%</span>
          {renderHoverActions(
            ACTIONS_LC,
            blockId,
            'progress-tracker',
            { label, progress: clampedProgress, status, detail, steps },
            'grim-pt-actions'
          )}
        </div>
      </div>

      <div style={barContainerStyle}>
        <div
          style={{
            width: `${clampedProgress}%`,
            height: '100%',
            borderRadius: 3,
            backgroundColor: getStatusStyle(status).color,
            transition: 'width 0.3s ease'
          }}
        />
      </div>

      {detail && <div style={detailStyle}>{detail}</div>}
      {steps && steps.length > 0 && (
        <div style={stepsContainerStyle}>
          {steps.map((step) => renderStep(step))}
        </div>
      )}
    </div>
  );
};
