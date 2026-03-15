/**
 * FilePreviewBlock
 * Renders file metadata and an optional preview iframe.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IFilePreviewData } from '../../../models/IBlock';
import { toWebViewerUrl } from '../../../utils/urlHelpers';
import { getFileTypeIcon } from '../../../utils/fileTypeIcons';
import { formatBytes } from '../../../utils/formatBytes';
import { injectHoverStyles, renderHoverActions, ACTIONS_SC } from './shared/HoverActions';
import { emitBlockInteraction } from '../interactionSchemas';

const metaRowStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  padding: '4px 0',
  fontSize: 12,
  borderBottom: '1px solid rgba(0, 0, 0, 0.05)'
};

const labelStyle: React.CSSProperties = {
  color: '#a19f9d'
};

const valueStyle: React.CSSProperties = {
  color: '#323130',
  textAlign: 'right',
  maxWidth: '60%',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const previewContainerStyle: React.CSSProperties = {
  marginTop: 12,
  borderRadius: 6,
  overflow: 'hidden',
  border: '1px solid rgba(0, 0, 0, 0.08)',
  backgroundColor: 'rgba(0, 0, 0, 0.02)'
};

const linkStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginTop: 12,
  fontSize: 12,
  color: '#0064b4',
  textDecoration: 'none',
  cursor: 'pointer'
};

export const FilePreviewBlock: React.FC<{ data: IFilePreviewData; blockId?: string }> = ({ data, blockId }) => {
  const { fileName, fileUrl, fileType, size, lastModified, author, previewUrl, metadata } = data;
  const iconName = getFileTypeIcon(fileType);

  React.useEffect(() => { injectHoverStyles('fp'); }, []);

  const metaEntries: Array<{ label: string; value: string }> = [
    { label: 'Type', value: fileType?.toUpperCase() || 'Unknown' },
    { label: 'Size', value: formatBytes(size, 'Unknown') }
  ];

  if (lastModified) metaEntries.push({ label: 'Modified', value: lastModified });
  if (author) metaEntries.push({ label: 'Author', value: author });

  if (metadata) {
    Object.keys(metadata).forEach((key) => {
      metaEntries.push({ label: key, value: metadata[key] });
    });
  }

  return (
    <div>
      {/* File header */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
        <Icon
          iconName={iconName}
          styles={{ root: { fontSize: 24, color: '#605e5c' } }}
        />
        <div style={{ flex: 1, overflow: 'hidden' }}>
          <div style={{ fontSize: 14, fontWeight: 600, color: '#323130', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {fileName}
          </div>
        </div>
      </div>

      {/* Metadata rows */}
      {metaEntries.map((entry, idx) => (
        <div key={idx} style={metaRowStyle}>
          <span style={labelStyle}>{entry.label}</span>
          <span style={valueStyle}>{entry.value}</span>
        </div>
      ))}

      {/* Preview iframe */}
      {previewUrl && (
        <div style={previewContainerStyle}>
          <iframe
            src={previewUrl}
            title={`Preview: ${fileName}`}
            style={{ width: '100%', height: 200, border: 'none' }}
            sandbox="allow-scripts allow-same-origin"
          />
        </div>
      )}

      {/* Open link + hover actions */}
      <div className="grim-fp-row" style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 12 }}>
        <a
          href={toWebViewerUrl(fileUrl)}
          target="_blank"
          rel="noopener noreferrer"
          data-interception="off"
          style={{ ...linkStyle, marginTop: 0 }}
          onClick={() => {
            emitBlockInteraction({
              blockId,
              blockType: 'file-preview',
              action: 'open-external',
              payload: { fileName, url: fileUrl },
              schemaId: 'file-preview.open-external',
              timestamp: Date.now()
            });
          }}
        >
          <Icon iconName="OpenInNewTab" styles={{ root: { fontSize: 12 } }} />
          Open in browser
        </a>
        {renderHoverActions(
          ACTIONS_SC,
          blockId,
          'file-preview',
          { fileName, url: fileUrl, fileType },
          'grim-fp-actions'
        )}
      </div>
    </div>
  );
};
