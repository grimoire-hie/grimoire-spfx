/**
 * SiteInfoBlock
 * Renders detailed information about a SharePoint site.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { ISiteInfoData } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';

const sectionStyle: React.CSSProperties = {
  marginTop: 12
};

const sectionLabelStyle: React.CSSProperties = {
  fontSize: 11,
  fontWeight: 600,
  color: '#a19f9d',
  textTransform: 'uppercase',
  letterSpacing: '0.5px',
  marginBottom: 4
};

const metaRowStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  padding: '3px 0',
  fontSize: 12
};

const labelStyle: React.CSSProperties = {
  color: '#a19f9d'
};

const valueStyle: React.CSSProperties = {
  color: '#323130',
  textAlign: 'right'
};

const tagStyle: React.CSSProperties = {
  display: 'inline-block',
  padding: '2px 8px',
  borderRadius: 4,
  backgroundColor: 'rgba(0, 0, 0, 0.04)',
  border: '1px solid rgba(0, 0, 0, 0.08)',
  fontSize: 11,
  color: '#605e5c',
  margin: '2px 4px 2px 0'
};

const linkStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginTop: 12,
  fontSize: 12,
  color: '#0064b4',
  textDecoration: 'none'
};

export const SiteInfoBlock: React.FC<{ data: ISiteInfoData; blockId?: string }> = ({ data, blockId }) => {
  const { siteName, siteUrl, description, owner, created, lastModified, storageUsed, libraries, lists } = data;

  return (
    <div>
      {/* Site header */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <Icon
          iconName="SharepointLogo"
          styles={{ root: { fontSize: 20, color: '#038387' } }}
        />
        <div style={{ fontSize: 14, fontWeight: 600, color: '#323130' }}>
          {siteName}
        </div>
      </div>

      {description && (
        <div style={{ fontSize: 12, color: '#605e5c', lineHeight: '1.4', marginBottom: 8 }}>
          {description}
        </div>
      )}

      {/* Metadata */}
      {owner && (
        <div style={metaRowStyle}>
          <span style={labelStyle}>Owner</span>
          <span style={valueStyle}>{owner}</span>
        </div>
      )}
      {created && (
        <div style={metaRowStyle}>
          <span style={labelStyle}>Created</span>
          <span style={valueStyle}>{created}</span>
        </div>
      )}
      {lastModified && (
        <div style={metaRowStyle}>
          <span style={labelStyle}>Modified</span>
          <span style={valueStyle}>{lastModified}</span>
        </div>
      )}
      {storageUsed && (
        <div style={metaRowStyle}>
          <span style={labelStyle}>Storage</span>
          <span style={valueStyle}>{storageUsed}</span>
        </div>
      )}

      {/* Libraries */}
      {libraries && libraries.length > 0 && (
        <div style={sectionStyle}>
          <div style={sectionLabelStyle}>Libraries</div>
          <div>{libraries.map((lib, idx) => <span key={idx} style={tagStyle}>{lib}</span>)}</div>
        </div>
      )}

      {/* Lists */}
      {lists && lists.length > 0 && (
        <div style={sectionStyle}>
          <div style={sectionLabelStyle}>Lists</div>
          <div>{lists.map((list, idx) => <span key={idx} style={tagStyle}>{list}</span>)}</div>
        </div>
      )}

      <a
        href={siteUrl}
        target="_blank"
        rel="noopener noreferrer"
        data-interception="off"
        style={linkStyle}
        onClick={() => {
          emitBlockInteraction({
            blockId,
            blockType: 'site-info',
            action: 'open-external',
            payload: { name: siteName, url: siteUrl },
            schemaId: 'site-info.open-external',
            timestamp: Date.now()
          });
        }}
      >
        <Icon iconName="OpenInNewTab" styles={{ root: { fontSize: 12 } }} />
        Open site
      </a>
    </div>
  );
};
