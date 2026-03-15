/**
 * ConfirmationDialogBlock
 * Renders a yes/no confirmation prompt with action buttons.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import * as strings from 'GrimoireAssistantWebPartStrings';
import type { IConfirmationDialogData } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';

const wrapperStyle: React.CSSProperties = {
  textAlign: 'center',
  padding: '8px 0'
};

const iconWrapperStyle: React.CSSProperties = {
  marginBottom: 12
};

const messageStyle: React.CSSProperties = {
  fontSize: 14,
  color: '#323130',
  lineHeight: '1.5',
  marginBottom: 16
};

const buttonRowStyle: React.CSSProperties = {
  display: 'flex',
  gap: 10,
  justifyContent: 'center'
};

const confirmButtonStyle: React.CSSProperties = {
  padding: '8px 20px',
  borderRadius: 6,
  border: 'none',
  background: 'rgba(64, 192, 64, 0.2)',
  color: '#107c10',
  fontSize: 13,
  fontWeight: 600,
  cursor: 'pointer'
};

const cancelButtonStyle: React.CSSProperties = {
  padding: '8px 20px',
  borderRadius: 6,
  border: '1px solid rgba(0, 0, 0, 0.15)',
  background: 'transparent',
  color: '#605e5c',
  fontSize: 13,
  cursor: 'pointer'
};

export const ConfirmationDialogBlock: React.FC<{ data: IConfirmationDialogData; blockId?: string }> = ({ data, blockId }) => {
  const { message, confirmLabel, cancelLabel } = data;
  const [responded, setResponded] = React.useState<'confirmed' | 'cancelled' | undefined>(undefined);

  if (responded) {
    return (
      <div style={{ textAlign: 'center', padding: '8px 0', fontSize: 13, color: '#605e5c' }}>
        <Icon
          iconName={responded === 'confirmed' ? 'CheckMark' : 'Cancel'}
          styles={{ root: { fontSize: 16, color: responded === 'confirmed' ? '#107c10' : '#e06060', marginRight: 6 } }}
        />
        {responded === 'confirmed' ? strings.ConfirmedLabel : strings.CancelledLabel}
      </div>
    );
  }

  return (
    <div style={wrapperStyle}>
      <div style={iconWrapperStyle}>
        <Icon
          iconName="Warning"
          styles={{ root: { fontSize: 28, color: '#f0a030' } }}
        />
      </div>
      <div style={messageStyle}>{message}</div>
      <div style={buttonRowStyle}>
        <button
          style={confirmButtonStyle}
          onClick={() => {
            setResponded('confirmed');
            emitBlockInteraction({
              blockId,
              blockType: 'confirmation-dialog',
              action: 'confirm',
              payload: { message },
              schemaId: 'confirmation.confirm',
              timestamp: Date.now()
            });
          }}
        >
          {confirmLabel}
        </button>
        <button
          style={cancelButtonStyle}
          onClick={() => {
            setResponded('cancelled');
            emitBlockInteraction({
              blockId,
              blockType: 'confirmation-dialog',
              action: 'cancel',
              payload: { message },
              schemaId: 'confirmation.cancel',
              timestamp: Date.now()
            });
          }}
        >
          {cancelLabel}
        </button>
      </div>
    </div>
  );
};
