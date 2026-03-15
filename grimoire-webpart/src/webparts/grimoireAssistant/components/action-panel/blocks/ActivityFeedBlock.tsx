/**
 * ActivityFeedBlock
 * Renders a timeline of recent activities (edits, shares, comments, etc.).
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { shallow } from 'zustand/shallow';
import type { IActivityFeedData } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';
import { useGrimoireStore } from '../../../store/useGrimoireStore';

const activityRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'stretch',
  padding: '8px 0',
  borderBottom: '1px solid rgba(0, 0, 0, 0.06)'
};

const dotStyle: React.CSSProperties = {
  width: 8,
  height: 8,
  borderRadius: '50%',
  marginTop: 5,
  flexShrink: 0
};

const actorStyle: React.CSSProperties = {
  fontWeight: 600,
  color: '#323130'
};

const actionTextStyle: React.CSSProperties = {
  fontSize: 12,
  color: '#605e5c',
  lineHeight: '1.4'
};

const timestampStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#c8c6c4',
  marginTop: 2
};

const checkboxStyle: React.CSSProperties = {
  width: 14,
  height: 14,
  borderRadius: 3,
  border: '1px solid rgba(0,0,0,0.25)',
  flexShrink: 0
};

const selectionColumnStyle: React.CSSProperties = {
  width: 28,
  minWidth: 28,
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'flex-start',
  paddingTop: 1,
  flexShrink: 0
};

const ACTION_ICONS: Record<string, { icon: string; color: string }> = {
  'created': { icon: 'Add', color: '#40c040' },
  'edited': { icon: 'Edit', color: '#6cb4ee' },
  'deleted': { icon: 'Delete', color: '#e04040' },
  'shared': { icon: 'Share', color: '#f0a030' },
  'commented': { icon: 'Comment', color: '#b080e0' },
  'uploaded': { icon: 'Upload', color: '#40c0c0' },
  'moved': { icon: 'MoveToFolder', color: '#f0a030' },
  'renamed': { icon: 'Rename', color: '#6cb4ee' }
};

function getActionInfo(action: string): { icon: string; color: string } {
  const lower = action.toLowerCase();
  const keys = Object.keys(ACTION_ICONS);
  for (let i = 0; i < keys.length; i++) {
    if (lower.includes(keys[i])) return ACTION_ICONS[keys[i]];
  }
  return { icon: 'InfoSolid', color: '#888888' };
}

export const ActivityFeedBlock: React.FC<{ data: IActivityFeedData; blockId?: string }> = ({ data, blockId }) => {
  const { activities } = data;
  const { activeActionBlockId, selectedActionIndices, toggleActionSelection } = useGrimoireStore((s) => ({
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    toggleActionSelection: s.toggleActionSelection
  }), shallow);
  const selectedSet = React.useMemo(() => new Set<number>(selectedActionIndices), [selectedActionIndices]);
  const canSelect = !!blockId && activeActionBlockId === blockId;

  if (activities.length === 0) {
    return (
      <div style={{ fontSize: 12, color: '#a19f9d', fontStyle: 'italic' }}>
        No recent activity.
      </div>
    );
  }

  return (
    <div>
      {activities.map((activity, idx) => {
        const info = getActionInfo(activity.action);
        const itemNum = idx + 1;
        const isSelected = canSelect && selectedSet.has(itemNum);
        return (
          <div
            key={idx}
            style={{ ...activityRowStyle, ...(isSelected ? { backgroundColor: 'rgba(0,100,180,0.08)', borderRadius: 4 } : {}) }}
            onClick={() => {
              emitBlockInteraction({
                blockId,
                blockType: 'activity-feed',
                action: 'click-activity',
                payload: { index: itemNum, target: activity.target, action: activity.action, actor: activity.actor, timestamp: activity.timestamp },
                schemaId: 'activity-feed.click-activity',
                timestamp: Date.now()
              });
            }}
          >
            <div style={selectionColumnStyle}>
              <input
                type="checkbox"
                checked={isSelected}
                disabled={!canSelect}
                style={{ ...checkboxStyle, opacity: canSelect ? 1 : 0.45, cursor: canSelect ? 'pointer' : 'default' }}
                onClick={(e) => { e.stopPropagation(); }}
                onChange={(e) => {
                  e.stopPropagation();
                  if (!blockId) return;
                  toggleActionSelection(blockId, itemNum);
                }}
              />
            </div>
            <div style={{ flex: 1, minWidth: 0, display: 'flex', alignItems: 'flex-start', gap: 10, paddingRight: 2 }}>
              <div style={{ ...dotStyle, background: info.color }} />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={actionTextStyle}>
                  <span style={actorStyle}>{activity.actor}</span>{' '}
                  {activity.action}{' '}
                  <Icon
                    iconName={info.icon}
                    styles={{ root: { fontSize: 11, color: info.color, verticalAlign: 'middle' } }}
                  />{' '}
                  <span style={{ color: '#0064b4' }}>{activity.target}</span>
                </div>
                <div style={timestampStyle}>{activity.timestamp}</div>
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
};
