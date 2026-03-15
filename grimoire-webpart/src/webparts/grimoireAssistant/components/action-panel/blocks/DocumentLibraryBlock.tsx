/**
 * DocumentLibraryBlock
 * Renders a browsable file/folder list from a SharePoint document library.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { shallow } from 'zustand/shallow';
import type { IDocumentLibraryData, IDocumentItem } from '../../../models/IBlock';
import { getFileTypeIcon } from '../../../utils/fileTypeIcons';
import { formatBytes } from '../../../utils/formatBytes';
import {
  extractSharePointSiteUrl,
  inferDocumentLibraryBaseUrl,
  resolveDocumentLibraryItemUrl
} from '../../../utils/sharePointUrlUtils';
import { emitBlockInteraction } from '../interactionSchemas';
import { useGrimoireStore } from '../../../store/useGrimoireStore';

const breadcrumbStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#a19f9d',
  marginBottom: 6
};

const columnHeaderStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  padding: '3px 0',
  borderBottom: '1px solid rgba(0, 0, 0, 0.08)',
  marginBottom: 2,
  fontSize: 10,
  color: '#a19f9d',
  fontWeight: 600,
  textTransform: 'uppercase',
  letterSpacing: '0.5px'
};

const itemRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'stretch',
  padding: '4px 0',
  borderBottom: '1px solid rgba(0, 0, 0, 0.05)',
  cursor: 'pointer'
};

const itemNameStyle: React.CSSProperties = {
  flex: 1,
  fontSize: 12,
  color: '#323130',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const itemMetaStyle: React.CSSProperties = {
  fontSize: 10,
  color: '#c8c6c4',
  whiteSpace: 'nowrap'
};

const checkboxStyle: React.CSSProperties = {
  width: 14,
  height: 14,
  borderRadius: 3,
  border: '1px solid rgba(0,0,0,0.25)'
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

function formatRelativeDate(dateStr?: string): string {
  if (!dateStr) return '';
  try {
    const date = new Date(dateStr);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    if (diffMins < 1) return 'just now';
    if (diffMins < 60) return `${diffMins}m ago`;
    const diffHours = Math.floor(diffMins / 60);
    if (diffHours < 24) return `${diffHours}h ago`;
    const diffDays = Math.floor(diffHours / 24);
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays}d ago`;
    if (diffDays < 30) return `${Math.floor(diffDays / 7)}w ago`;
    if (diffDays < 365) return `${Math.floor(diffDays / 30)}mo ago`;
    return `${Math.floor(diffDays / 365)}y ago`;
  } catch {
    return dateStr;
  }
}

function getItemIcon(item: IDocumentItem): string {
  if (item.type === 'folder') return 'FabricFolder';
  return getFileTypeIcon(item.fileType);
}

function getItemIconColor(item: IDocumentItem): string {
  if (item.type === 'folder') return '#ffc107';
  return '#a19f9d';
}

export const DocumentLibraryBlock: React.FC<{ data: IDocumentLibraryData; blockId?: string }> = ({ data, blockId }) => {
  const { siteName, libraryName, items, breadcrumb } = data;
  const { activeActionBlockId, selectedActionIndices, toggleActionSelection } = useGrimoireStore((s) => ({
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    toggleActionSelection: s.toggleActionSelection
  }), shallow);
  const selectedSet = React.useMemo(() => new Set<number>(selectedActionIndices), [selectedActionIndices]);
  const canSelect = !!blockId && activeActionBlockId === blockId;
  const siblingUrls = React.useMemo(() => items.map((item) => item.url), [items]);
  const documentLibraryUrl = React.useMemo(
    () => inferDocumentLibraryBaseUrl(siblingUrls),
    [siblingUrls]
  );
  const siteUrl = React.useMemo(
    () => extractSharePointSiteUrl(documentLibraryUrl || siblingUrls[0]),
    [documentLibraryUrl, siblingUrls]
  );
  const normalizedSiteName = React.useMemo(
    () => (/^https?:\/\//i.test(siteName || '') ? undefined : siteName),
    [siteName]
  );

  const hasSize = items.some((i) => i.type === 'file' && i.size !== undefined);
  const hasDate = items.some((i) => !!i.lastModified);

  return (
    <div>
      <div style={breadcrumbStyle}>
        {siteName} / {libraryName}
        {breadcrumb.length > 0 && ` / ${breadcrumb.join(' / ')}`}
      </div>

      {items.length === 0 ? (
        <div style={{ fontSize: 12, color: '#c8c6c4', fontStyle: 'italic' }}>
          Empty folder
        </div>
      ) : (
        <>
          <div style={columnHeaderStyle}>
            <span style={{ width: 18 }} />
            <span style={{ width: 14 }} />
            <span style={{ flex: 1 }}>Name</span>
            {hasSize && <span style={{ width: 55, textAlign: 'right' }}>Size</span>}
            {hasDate && <span style={{ width: 65, textAlign: 'right' }}>Modified</span>}
          </div>
          {items.map((item, idx) => {
            const itemNum = idx + 1;
            const isSelected = canSelect && selectedSet.has(itemNum);
            const resolvedUrl = resolveDocumentLibraryItemUrl(item.url, item.name, siblingUrls) || item.url;
            return (
              <div
                key={idx}
                className="grim-dl-row"
                style={{ ...itemRowStyle, ...(isSelected ? { backgroundColor: 'rgba(0,100,180,0.08)', borderRadius: 4 } : {}) }}
                title={resolvedUrl}
                onClick={() => {
                  emitBlockInteraction({
                    blockId,
                    blockType: 'document-library',
                    action: item.type === 'folder' ? 'click-folder' : 'click-file',
                    schemaId: item.type === 'folder' ? 'document-library.click-folder' : 'document-library.click-file',
                    payload: {
                      index: itemNum,
                      name: item.name,
                      title: item.name,
                      url: resolvedUrl,
                      fileOrFolderUrl: resolvedUrl,
                      fileOrFolderId: item.fileOrFolderId,
                      documentLibraryId: item.documentLibraryId,
                      documentLibraryUrl,
                      documentLibraryName: libraryName,
                      siteUrl,
                      siteName: normalizedSiteName,
                      fileType: item.fileType,
                      type: item.type
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
                <div style={{ flex: 1, minWidth: 0, display: 'flex', alignItems: 'center', gap: 6, paddingRight: 2 }}>
                  <Icon
                    iconName={getItemIcon(item)}
                    styles={{ root: { fontSize: 14, color: getItemIconColor(item) } }}
                  />
                  <span style={itemNameStyle}>{item.name}</span>
                  {hasSize && (
                    <span style={{ ...itemMetaStyle, width: 55, textAlign: 'right' }}>
                      {item.type === 'file' ? formatBytes(item.size) : ''}
                    </span>
                  )}
                  {hasDate && (
                    <span style={{ ...itemMetaStyle, width: 65, textAlign: 'right' }}>
                      {formatRelativeDate(item.lastModified)}
                    </span>
                  )}
                </div>
              </div>
            );
          })}
        </>
      )}

      <div style={{ fontSize: 10, color: '#c8c6c4', marginTop: 6 }}>
        {items.length} item{items.length !== 1 ? 's' : ''}
      </div>
    </div>
  );
};
