/**
 * ListItemsBlock
 * Renders SharePoint list items in a table layout with columns.
 */

import * as React from 'react';
import { shallow } from 'zustand/shallow';
import type { IListItemsData, IRenderHints } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';
import { useGrimoireStore } from '../../../store/useGrimoireStore';

const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  fontSize: 12
};

const thStyle: React.CSSProperties = {
  textAlign: 'left',
  padding: '6px 8px',
  borderBottom: '1px solid rgba(0, 0, 0, 0.12)',
  color: '#605e5c',
  fontWeight: 600,
  fontSize: 11,
  textTransform: 'uppercase',
  letterSpacing: '0.3px',
  whiteSpace: 'nowrap'
};

const tdStyle: React.CSSProperties = {
  padding: '6px 8px',
  borderBottom: '1px solid rgba(0, 0, 0, 0.06)',
  color: '#323130',
  maxWidth: 200,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const linkCellStyle: React.CSSProperties = {
  color: '#0f6cbd',
  textDecoration: 'underline',
  wordBreak: 'break-all' as const,
  whiteSpace: 'normal' as const
};

const footerStyle: React.CSSProperties = {
  paddingTop: 8,
  fontSize: 11,
  color: '#a19f9d'
};

const checkboxStyle: React.CSSProperties = {
  width: 14,
  height: 14,
  borderRadius: 3,
  border: '1px solid rgba(0,0,0,0.25)'
};

export const ListItemsBlock: React.FC<{ data: IListItemsData; blockId?: string; renderHints?: IRenderHints }> = ({ data, blockId, renderHints }) => {
  const { listName, columns, items, totalCount } = data;
  const { activeActionBlockId, selectedActionIndices, toggleActionSelection } = useGrimoireStore((s) => ({
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    toggleActionSelection: s.toggleActionSelection
  }), shallow);
  const selectedSet = React.useMemo(() => new Set<number>(selectedActionIndices), [selectedActionIndices]);
  const canSelect = !!blockId && activeActionBlockId === blockId;

  if (items.length === 0) {
    return (
      <div style={{ fontSize: 12, color: '#a19f9d', fontStyle: 'italic' }}>
        No items in &ldquo;{listName}&rdquo;.
      </div>
    );
  }

  const isUrlColumn = (columnName: string, value: string | undefined): boolean => {
    if (!value) return false;
    const lowered = columnName.toLowerCase();
    return lowered.includes('url') || lowered.includes('link') || /^https?:\/\//i.test(value);
  };

  return (
    <div>
      <div style={{ overflowX: 'auto' }}>
        <table style={tableStyle}>
          <thead>
            <tr>
              <th style={{ ...thStyle, width: 28, padding: '6px 4px' }} />
              {columns.map((col) => (
                <th key={col} style={thStyle}>{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {items.map((item, idx) => {
              const itemNum = idx + 1;
              const isSelected = canSelect && selectedSet.has(itemNum);
              const isHighlighted = renderHints?.highlight?.indexOf(itemNum) !== -1 && renderHints?.highlight !== undefined;
              const bgColor = idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.03)';
              const trStyle: React.CSSProperties = isHighlighted
                ? { background: bgColor, cursor: 'pointer', borderLeft: '3px solid #4fc3f7' }
                : { background: bgColor, cursor: 'pointer' };
              if (isSelected) trStyle.background = 'rgba(0,100,180,0.08)';

              return (
              <tr
                key={idx}
                className="grim-li-row"
                style={trStyle}
                onClick={() => {
                  emitBlockInteraction({
                    blockId,
                    blockType: 'list-items',
                    action: 'click-list-row',
                    payload: { rowData: item, index: itemNum },
                    schemaId: 'list-items.click-list-row',
                    timestamp: Date.now()
                  });
                }}
              >
                <td style={{ ...tdStyle, width: 28, padding: '6px 4px', textAlign: 'center' }}>
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
                </td>
                {columns.map((col) => (
                  <td key={col} style={tdStyle} title={item[col] || ''}>
                    {isUrlColumn(col, item[col]) ? (
                      <a
                        href={item[col]}
                        target="_blank"
                        rel="noreferrer"
                        data-interception="off"
                        style={linkCellStyle}
                        onClick={(e) => e.stopPropagation()}
                      >
                        {item[col]}
                      </a>
                    ) : (
                      item[col] || '—'
                    )}
                  </td>
                ))}
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={footerStyle}>
        {items.length} of {totalCount} item{totalCount !== 1 ? 's' : ''}
      </div>
    </div>
  );
};
