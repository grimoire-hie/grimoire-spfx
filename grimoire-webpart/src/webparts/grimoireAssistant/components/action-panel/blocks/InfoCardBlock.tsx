/**
 * InfoCardBlock
 * Renders a simple informational card with heading, body, and optional icon.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IInfoCardData } from '../../../models/IBlock';
import { injectHoverStyles, renderHoverActions, ACTIONS_LC } from './shared/HoverActions';

const URL_SPLIT_PATTERN = /(https?:\/\/[^\s]+)/gi;
const URL_MATCH_PATTERN = /^https?:\/\/[^\s]+$/i;

const linkStyle: React.CSSProperties = {
  color: '#0f6cbd',
  textDecoration: 'underline',
  wordBreak: 'break-all'
};

export const InfoCardBlock: React.FC<{ data: IInfoCardData; blockId?: string }> = ({ data, blockId }) => {
  const { heading, body, icon } = data;
  React.useEffect(() => { injectHoverStyles('ic'); }, []);

  const renderedBody = React.useMemo(() => {
    const parts = body.split(URL_SPLIT_PATTERN);
    return parts.map((part, index) => {
      if (!URL_MATCH_PATTERN.test(part)) {
        return <React.Fragment key={`ic-text-${index}`}>{part}</React.Fragment>;
      }

      return (
        <a
          key={`ic-link-${index}`}
          href={part}
          target="_blank"
          rel="noopener noreferrer"
          data-interception="off"
          style={linkStyle}
        >
          {part}
        </a>
      );
    });
  }, [body]);

  return (
    <div className="grim-ic-row">
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        {icon && (
          <Icon
            iconName={icon}
            styles={{ root: { fontSize: 18, color: '#605e5c' } }}
          />
        )}
        <span style={{ marginLeft: 'auto' }}>
          {renderHoverActions(ACTIONS_LC, blockId, 'info-card', { heading, body, icon }, 'grim-ic-actions')}
        </span>
      </div>
      <div style={{ fontSize: 13, color: '#605e5c', lineHeight: '1.5', whiteSpace: 'pre-wrap' }}>
        {renderedBody}
      </div>
    </div>
  );
};
