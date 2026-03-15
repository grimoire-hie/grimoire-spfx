/**
 * SelectionListBlock
 * Renders a list of selectable items (single or multi-select).
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { ISelectionListData } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';

const promptStyle: React.CSSProperties = {
  fontSize: 13,
  color: '#323130',
  marginBottom: 10
};

const itemStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  padding: '8px 10px',
  borderRadius: 6,
  cursor: 'pointer',
  marginBottom: 4,
  transition: 'background 0.15s ease'
};

const checkboxStyle: React.CSSProperties = {
  width: 18,
  height: 18,
  borderRadius: 4,
  border: '2px solid rgba(0, 0, 0, 0.2)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  flexShrink: 0,
  transition: 'all 0.15s ease'
};

const checkedBoxStyle: React.CSSProperties = {
  ...checkboxStyle,
  background: 'rgba(0, 100, 180, 0.15)',
  borderColor: '#0064b4'
};

const radioStyle: React.CSSProperties = {
  ...checkboxStyle,
  borderRadius: '50%'
};

const checkedRadioStyle: React.CSSProperties = {
  ...radioStyle,
  background: 'rgba(0, 100, 180, 0.15)',
  borderColor: '#0064b4'
};

const labelStyle: React.CSSProperties = {
  fontSize: 13,
  color: '#323130'
};

const descStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#a19f9d',
  marginTop: 1
};

const submitBtnStyle: React.CSSProperties = {
  marginTop: 10,
  padding: '8px 20px',
  borderRadius: 6,
  border: 'none',
  background: 'rgba(0, 100, 180, 0.12)',
  color: '#0064b4',
  fontSize: 13,
  fontWeight: 600,
  cursor: 'pointer',
  width: '100%'
};

export const SelectionListBlock: React.FC<{ data: ISelectionListData; blockId?: string }> = ({ data, blockId }) => {
  const { prompt, items: initialItems, multiSelect } = data;
  const [selected, setSelected] = React.useState<Set<string>>(() => {
    const init = new Set<string>();
    initialItems.forEach((item) => {
      if (item.selected) init.add(item.id);
    });
    return init;
  });
  const [submitted, setSubmitted] = React.useState(false);

  const sendSelection = React.useCallback((selectedIds: Set<string>) => {
    const selectedLabels: string[] = [];
    const selectedItems: Array<Record<string, unknown>> = [];
    initialItems.forEach((item) => {
      if (selectedIds.has(item.id)) {
        selectedLabels.push(item.label);
        selectedItems.push({
          id: item.id,
          label: item.label,
          ...(item.itemType ? { itemType: item.itemType } : {}),
          ...(item.description ? { description: item.description } : {}),
          ...(item.targetContext ? { targetContext: item.targetContext } : {})
        });
      }
    });
    const label = selectedLabels.length > 0 ? selectedLabels.join(', ') : '(none)';
    emitBlockInteraction({
      blockId,
      blockType: 'selection-list',
      action: 'select',
      payload: { label, prompt, selectedIds: Array.from(selectedIds), selectedItems },
      schemaId: 'selection.select',
      timestamp: Date.now()
    });
  }, [initialItems, prompt, blockId]);

  const handleToggle = React.useCallback((itemId: string) => {
    if (submitted) return;
    setSelected((prev) => {
      const next = new Set(prev);
      if (multiSelect) {
        if (next.has(itemId)) {
          next.delete(itemId);
        } else {
          next.add(itemId);
        }
      } else {
        next.clear();
        next.add(itemId);
        // For single-select, submit immediately
        setSubmitted(true);
        const singleSet = new Set<string>();
        singleSet.add(itemId);
        sendSelection(singleSet);
      }
      return next;
    });
  }, [multiSelect, submitted, sendSelection]);

  const handleSubmit = React.useCallback(() => {
    setSubmitted(true);
    sendSelection(selected);
  }, [selected, sendSelection]);

  return (
    <div>
      <div style={promptStyle}>{prompt}</div>
      {initialItems.map((item) => {
        const isSelected = selected.has(item.id);
        const box = multiSelect
          ? (isSelected ? checkedBoxStyle : checkboxStyle)
          : (isSelected ? checkedRadioStyle : radioStyle);

        return (
          <div
            key={item.id}
            style={{
              ...itemStyle,
              background: isSelected ? 'rgba(0, 100, 180, 0.06)' : 'transparent',
              opacity: submitted ? 0.6 : 1,
              pointerEvents: submitted ? 'none' : 'auto'
            }}
            onClick={() => handleToggle(item.id)}
          >
            <div style={box}>
              {isSelected && (
                <Icon
                  iconName="CheckMark"
                  styles={{ root: { fontSize: 11, color: '#0064b4' } }}
                />
              )}
            </div>
            <div>
              <div style={labelStyle}>{item.label}</div>
              {item.description && <div style={descStyle}>{item.description}</div>}
            </div>
          </div>
        );
      })}
      {multiSelect && !submitted && (
        <button style={submitBtnStyle} onClick={handleSubmit}>
          Confirm Selection
        </button>
      )}
      {submitted && (
        <div style={{ textAlign: 'center', marginTop: 8, fontSize: 12, color: '#a19f9d' }}>
          Selection sent
        </div>
      )}
    </div>
  );
};
