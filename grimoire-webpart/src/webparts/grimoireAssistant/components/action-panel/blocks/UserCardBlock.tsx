/**
 * UserCardBlock
 * Renders a user profile card with photo, name, email, and metadata.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IUserCardData } from '../../../models/IBlock';
import { injectHoverStyles, renderHoverActions, ACTIONS_LC } from './shared/HoverActions';
import { emitBlockInteraction } from '../interactionSchemas';

const cardStyle: React.CSSProperties = {
  display: 'flex',
  gap: 14,
  alignItems: 'flex-start'
};

const avatarStyle: React.CSSProperties = {
  width: 56,
  height: 56,
  borderRadius: '50%',
  objectFit: 'cover',
  flexShrink: 0,
  border: '2px solid rgba(0, 0, 0, 0.1)'
};

const avatarPlaceholderStyle: React.CSSProperties = {
  ...avatarStyle,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  background: 'rgba(0, 100, 180, 0.1)',
  color: '#605e5c',
  fontSize: 22
};

const nameStyle: React.CSSProperties = {
  fontSize: 15,
  fontWeight: 600,
  color: '#323130',
  marginBottom: 2
};

const emailStyle: React.CSSProperties = {
  fontSize: 12,
  color: '#0064b4',
  textDecoration: 'none',
  marginBottom: 8
};

const metaRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  fontSize: 12,
  color: '#605e5c',
  marginBottom: 3
};

export const UserCardBlock: React.FC<{ data: IUserCardData; blockId?: string }> = ({ data, blockId }) => {
  const { displayName, email, jobTitle, department, officeLocation, phone, photoUrl } = data;
  const [photoFailed, setPhotoFailed] = React.useState(false);

  React.useEffect(() => { injectHoverStyles('uc'); }, []);
  React.useEffect(() => { setPhotoFailed(false); }, [photoUrl]);

  return (
    <div
      className="grim-uc-row"
      style={{ ...cardStyle, cursor: 'pointer' }}
      onClick={() => {
        emitBlockInteraction({
          blockId,
          blockType: 'user-card',
          action: 'click-user',
          schemaId: 'user-card.click-user',
          payload: { displayName, email, jobTitle, department, officeLocation, phone },
          timestamp: Date.now()
        });
      }}
    >
      {photoUrl && !photoFailed ? (
        <img src={photoUrl} alt={displayName} style={avatarStyle} onError={() => setPhotoFailed(true)} />
      ) : (
        <div style={avatarPlaceholderStyle}>
          <Icon iconName="Contact" />
        </div>
      )}
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={nameStyle}>{displayName}</div>
          {renderHoverActions(
            ACTIONS_LC,
            blockId,
            'user-card',
            { displayName, email, jobTitle, department },
            'grim-uc-actions'
          )}
        </div>
        <a href={`mailto:${email}`} style={emailStyle}>{email}</a>
        {jobTitle && (
          <div style={metaRowStyle}>
            <Icon iconName="Work" styles={{ root: { fontSize: 12, color: '#a19f9d' } }} />
            <span>{jobTitle}</span>
          </div>
        )}
        {department && (
          <div style={metaRowStyle}>
            <Icon iconName="Org" styles={{ root: { fontSize: 12, color: '#a19f9d' } }} />
            <span>{department}</span>
          </div>
        )}
        {officeLocation && (
          <div style={metaRowStyle}>
            <Icon iconName="MapPin" styles={{ root: { fontSize: 12, color: '#a19f9d' } }} />
            <span>{officeLocation}</span>
          </div>
        )}
        {phone && (
          <div style={metaRowStyle}>
            <Icon iconName="Phone" styles={{ root: { fontSize: 12, color: '#a19f9d' } }} />
            <span>{phone}</span>
          </div>
        )}
      </div>
    </div>
  );
};
