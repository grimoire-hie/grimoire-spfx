/**
 * PermissionsViewBlock
 * Renders a permissions table for a SharePoint site/library/file.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { shallow } from 'zustand/shallow';
import type { IPermissionsViewData } from '../../../models/IBlock';
import { toWebViewerUrl } from '../../../utils/urlHelpers';
import { useGrimoireStore } from '../../../store/useGrimoireStore';
import { emitBlockInteraction } from '../interactionSchemas';

const headerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  marginBottom: 10,
  fontSize: 12,
  color: '#605e5c'
};

const rowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'stretch',
  padding: '6px 0',
  borderBottom: '1px solid rgba(0, 0, 0, 0.06)'
};

const principalStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  fontSize: 13,
  color: '#323130',
  flex: 1,
  minWidth: 0,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const roleBadgeBase: React.CSSProperties = {
  fontSize: 11,
  padding: '2px 8px',
  borderRadius: 10,
  fontWeight: 600,
  flexShrink: 0
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

const ROLE_COLORS: Record<string, { bg: string; fg: string }> = {
  'Full Control': { bg: 'rgba(224, 64, 64, 0.2)', fg: '#e06060' },
  'Owner': { bg: 'rgba(224, 64, 64, 0.2)', fg: '#e06060' },
  'Edit': { bg: 'rgba(240, 160, 48, 0.2)', fg: '#d48200' },
  'Contribute': { bg: 'rgba(240, 160, 48, 0.2)', fg: '#d48200' },
  'Read': { bg: 'rgba(64, 192, 64, 0.2)', fg: '#107c10' },
  'View Only': { bg: 'rgba(64, 192, 64, 0.2)', fg: '#107c10' }
};

export const PermissionsViewBlock: React.FC<{ data: IPermissionsViewData; blockId?: string }> = ({ data, blockId }) => {
  const { targetName, targetUrl, permissions } = data;
  const { activeActionBlockId, selectedActionIndices, toggleActionSelection } = useGrimoireStore((s) => ({
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    toggleActionSelection: s.toggleActionSelection
  }), shallow);
  const selectedSet = React.useMemo(() => new Set<number>(selectedActionIndices), [selectedActionIndices]);
  const canSelect = !!blockId && activeActionBlockId === blockId;

  if (permissions.length === 0) {
    return (
      <div style={{ fontSize: 12, color: '#a19f9d', fontStyle: 'italic' }}>
        No permissions data available.
      </div>
    );
  }

  return (
    <div>
      <div style={headerStyle}>
        <Icon iconName="Lock" styles={{ root: { fontSize: 14 } }} />
        <a
          href={toWebViewerUrl(targetUrl)}
          target="_blank"
          rel="noopener noreferrer"
          data-interception="off"
          style={{ color: '#0064b4', textDecoration: 'none' }}
          onClick={() => {
            emitBlockInteraction({
              blockId,
              blockType: 'permissions-view',
              action: 'open-external',
              payload: { name: targetName, url: targetUrl },
              timestamp: Date.now()
            });
          }}
        >
          {targetName}
        </a>
      </div>

      {permissions.map((p, idx) => {
        const colors = ROLE_COLORS[p.role] || { bg: 'rgba(128,128,128,0.2)', fg: '#888' };
        const itemNum = idx + 1;
        const isSelected = canSelect && selectedSet.has(itemNum);
        return (
          <div
            key={idx}
            style={{ ...rowStyle, ...(isSelected ? { backgroundColor: 'rgba(0,100,180,0.08)', borderRadius: 4 } : {}) }}
            onClick={() => {
              emitBlockInteraction({
                blockId,
                blockType: 'permissions-view',
                action: 'click-permission',
                schemaId: 'permissions-view.click-permission',
                payload: {
                  index: itemNum,
                  principal: p.principal,
                  role: p.role,
                  inherited: p.inherited,
                  targetName,
                  targetUrl
                },
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
            <div style={{ flex: 1, minWidth: 0, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, paddingRight: 2 }}>
              <div style={principalStyle}>
                <Icon
                  iconName={p.inherited ? 'People' : 'Contact'}
                  styles={{ root: { fontSize: 14, color: '#a19f9d', flexShrink: 0 } }}
                />
                <span>{p.principal}</span>
                {p.inherited && (
                  <span style={{ fontSize: 10, color: '#c8c6c4', marginLeft: 4 }}>
                    (inherited)
                  </span>
                )}
              </div>
              <span style={{ ...roleBadgeBase, background: colors.bg, color: colors.fg }}>
                {p.role}
              </span>
            </div>
          </div>
        );
      })}
    </div>
  );
};
