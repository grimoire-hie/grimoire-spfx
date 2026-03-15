/**
 * ErrorBlock
 * Renders an error message with optional detail and retry action.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import * as strings from 'GrimoireAssistantWebPartStrings';
import type { IErrorData } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';

const containerStyle: React.CSSProperties = {
  display: 'flex',
  gap: 10,
  alignItems: 'flex-start'
};

const iconContainerStyle: React.CSSProperties = {
  flexShrink: 0,
  width: 32,
  height: 32,
  borderRadius: '50%',
  backgroundColor: 'rgba(244, 67, 54, 0.15)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center'
};

const messageStyle: React.CSSProperties = {
  fontSize: 13,
  fontWeight: 600,
  color: '#ef5350',
  marginBottom: 4
};

const detailStyle: React.CSSProperties = {
  fontSize: 12,
  color: '#605e5c',
  lineHeight: '1.4',
  whiteSpace: 'pre-wrap',
  fontFamily: 'Consolas, "Courier New", monospace',
  backgroundColor: '#f3f2f1',
  padding: '6px 8px',
  borderRadius: 4,
  marginTop: 6,
  maxHeight: 120,
  overflowY: 'auto'
};

export const ErrorBlock: React.FC<{ data: IErrorData; blockId?: string }> = ({ data, blockId }) => {
  const { message, detail, retryAction } = data;

  return (
    <div style={containerStyle}>
      <div style={iconContainerStyle}>
        <Icon
          iconName="ErrorBadge"
          styles={{ root: { fontSize: 16, color: '#ef5350' } }}
        />
      </div>
      <div style={{ flex: 1 }}>
        <div style={messageStyle}>{message}</div>
        {detail && <div style={detailStyle}>{detail}</div>}
        {retryAction && (
          <button
            style={{
              marginTop: 10,
              padding: '4px 12px',
              fontSize: 12,
              color: '#ef5350',
              background: 'rgba(244, 67, 54, 0.1)',
              border: '1px solid rgba(244, 67, 54, 0.35)',
              borderRadius: 4,
              cursor: 'pointer'
            }}
            onClick={() => {
              emitBlockInteraction({
                blockId,
                blockType: 'error',
                action: 'retry',
                payload: { retryAction: data.retryAction },
                schemaId: 'error.retry',
                timestamp: Date.now()
              });
            }}
          >
            {strings.RetryButton}
          </button>
        )}
      </div>
    </div>
  );
};
