/**
 * SearchResultsBlock
 * Renders search results from Copilot Search, Copilot Retrieval, and SharePoint Search.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { shallow } from 'zustand/shallow';
import type { ISearchResultsData, IRenderHints } from '../../../models/IBlock';
import { toWebViewerUrl } from '../../../utils/urlHelpers';
import { getFileTypeIcon } from '../../../utils/fileTypeIcons';
import { emitBlockInteraction } from '../interactionSchemas';
import { useGrimoireStore } from '../../../store/useGrimoireStore';
import {
  describeSearchQueryBreadth,
  formatSearchQueryVariantLabel,
  getUserFacingSearchQueryVariants
} from '../../../services/search/SearchQueryVariantPresentation';

// Source badge colors (work on light backgrounds)
const SOURCE_BADGE_COLORS: Record<string, { color: string; bg: string }> = {
  'copilot-search': { color: '#1a73e8', bg: 'rgba(26, 115, 232, 0.1)' },
  'copilot-retrieval': { color: '#2e7d32', bg: 'rgba(46, 125, 50, 0.1)' },
  'sharepoint-search': { color: '#e65100', bg: 'rgba(230, 81, 0, 0.1)' }
};

const SOURCE_LABELS: Record<string, string> = {
  'copilot-search': 'Copilot Search',
  'copilot-retrieval': 'Copilot Retrieval',
  'sharepoint-search': 'SharePoint Search'
};

function getFileIcon(fileType?: string): string {
  return getFileTypeIcon(fileType);
}

/** SVG shape paths: circle = Copilot Search, triangle = Copilot Retrieval, square = SharePoint Search */
const SOURCE_SHAPES: Record<string, React.ReactNode> = {
  'copilot-search': <circle cx="4" cy="4" r="3.5" />,
  'copilot-retrieval': <polygon points="4,0.5 7.5,7.5 0.5,7.5" />,
  'sharepoint-search': <rect x="0.5" y="0.5" width="7" height="7" />
};

function renderSourceShape(sourceKey: string, size: number): React.ReactNode {
  const colors = SOURCE_BADGE_COLORS[sourceKey];
  const shape = SOURCE_SHAPES[sourceKey];
  if (!colors || !shape) return null;
  return (
    <svg width={size} height={size} viewBox="0 0 8 8" style={{ flexShrink: 0 }} fill={colors.color}>
      {shape}
    </svg>
  );
}

/** Render colored source shapes for a single result row */
function renderSourceShapes(resultSources?: string[]): React.ReactNode {
  if (!resultSources || resultSources.length === 0) return null;
  return (
    <span style={{ display: 'inline-flex', gap: 3, marginLeft: 4, alignItems: 'center' }}>
      {resultSources.map((s) => {
        const label = SOURCE_LABELS[s] || s;
        return (
          <span key={s} title={label}>
            {renderSourceShape(s, 9)}
          </span>
        );
      })}
    </span>
  );
}

/** Extract a readable path from a SharePoint URL for display as fallback context */
function extractUrlPath(url: string): string {
  try {
    const u = new URL(url);
    const path = decodeURIComponent(u.pathname);
    const cleaned = path
      .replace(/^\/sites\/[^/]+/, '')
      .replace(/^\/personal\/[^/]+/, '')
      .replace(/^\/_layouts\/.*$/, '')
      .replace(/^\/Shared Documents/, '/Docs')
      .replace(/^\//, '');
    return cleaned || u.hostname;
  } catch {
    return url.substring(0, 60);
  }
}

// Light-theme colors
const subtleText = '#a19f9d';
const dimText = 'rgba(0,0,0,0.45)';
const veryDimText = 'rgba(0,0,0,0.4)';
const borderColor = 'rgba(0,0,0,0.08)';
const iconColor = 'rgba(0,0,0,0.4)';
const hoverBg = 'rgba(0,0,0,0.03)';
const summaryColor = '#605e5c';
const badgeBg = 'rgba(0,0,0,0.06)';
const titleLinkColor = '#0064b4';
const fileTypeBadgeColor = 'rgba(0,0,0,0.5)';
const checkboxStyle: React.CSSProperties = {
  width: 15,
  height: 15,
  borderRadius: 3,
  border: '1px solid rgba(0,0,0,0.25)'
};

const selectionColumnStyle: React.CSSProperties = {
  width: 28,
  minWidth: 28,
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'flex-start',
  paddingTop: 2,
  flexShrink: 0
};

export const SearchResultsBlock: React.FC<{ data: ISearchResultsData; blockId?: string; renderHints?: IRenderHints }> = ({ data, blockId, renderHints }) => {
  const { query, results, totalCount, source, queryVariants } = data;
  const [showVariantDetails, setShowVariantDetails] = React.useState(false);
  const { activeActionBlockId, selectedActionIndices, toggleActionSelection } = useGrimoireStore((s) => ({
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    toggleActionSelection: s.toggleActionSelection
  }), shallow);
  const selectedSet = React.useMemo(() => new Set<number>(selectedActionIndices), [selectedActionIndices]);
  const userFacingVariants = React.useMemo(
    () => getUserFacingSearchQueryVariants(queryVariants),
    [queryVariants]
  );
  const searchBreadthSummary = React.useMemo(
    () => describeSearchQueryBreadth(queryVariants),
    [queryVariants]
  );
  const canSelect = !!blockId && activeActionBlockId === blockId;
  const sources = source.split('+').filter(Boolean);

  if (results.length === 0) {
    return (
      <div>
        <div style={{ fontSize: 12, color: subtleText, marginBottom: 8 }}>
          Searching for &ldquo;{query}&rdquo;...
        </div>
        {source === 'pending' ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <Spinner size={SpinnerSize.xSmall} />
            <span style={{ fontSize: 12, color: dimText }}>
              Searching across sources...
            </span>
          </div>
        ) : (
          <div style={{ fontSize: 12, color: dimText, fontStyle: 'italic' }}>
            No results found.
          </div>
        )}
      </div>
    );
  }

  return (
    <div>
      <div style={{ fontSize: 12, color: subtleText, marginBottom: 4 }}>
        &ldquo;{query}&rdquo;
      </div>

      {userFacingVariants.length > 0 && (
        <div style={{ marginBottom: 8 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'center' }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: subtleText }}>
              Search breadth
            </div>
            <button
              type="button"
              onClick={() => setShowVariantDetails((current) => !current)}
              style={{
                border: `1px solid ${borderColor}`,
                background: 'rgba(0,100,180,0.06)',
                color: '#005a9c',
                borderRadius: 999,
                padding: '2px 8px',
                fontSize: 10,
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              {showVariantDetails ? 'Hide details' : 'Show details'}
            </button>
          </div>
          {searchBreadthSummary && (
            <div style={{ fontSize: 11, color: dimText, marginTop: 4, lineHeight: 1.45 }}>
              {searchBreadthSummary}
            </div>
          )}
          {showVariantDetails && (
            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap', marginTop: 6 }}>
              {userFacingVariants.map((variant) => (
                <span
                  key={`${variant.kind}:${variant.query}`}
                  style={{
                    fontSize: 10,
                    fontWeight: 600,
                    color: '#005a9c',
                    background: 'rgba(0,100,180,0.08)',
                    borderRadius: 3,
                    padding: '2px 6px',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: 3
                  }}
                >
                  <Icon iconName="Search" styles={{ root: { fontSize: 9 } }} />
                  {formatSearchQueryVariantLabel(variant)}
                </span>
              ))}
            </div>
          )}
          </div>
      )}

      {results.map((result, idx) => {
        const itemNum = idx + 1;
        const isHighlighted = renderHints?.highlight?.indexOf(itemNum) !== -1 && renderHints?.highlight !== undefined;
        const annotation = renderHints?.annotate ? renderHints.annotate[itemNum] : undefined;
        const isSelected = canSelect && selectedSet.has(itemNum);

        const rowStyle: React.CSSProperties = {
          display: 'flex',
          alignItems: 'stretch',
          padding: '6px 0 5px',
          borderBottom: `1px solid ${borderColor}`,
          cursor: 'pointer',
          borderRadius: 4,
          ...(isSelected ? { backgroundColor: 'rgba(0,100,180,0.06)' } : {}),
          ...(isHighlighted ? { borderLeft: '3px solid #4fc3f7', paddingLeft: 8 } : {})
        };

        return (
        <div
          key={idx}
          className="grim-sr-row"
          style={rowStyle}
          onMouseEnter={(e) => { (e.currentTarget as HTMLDivElement).style.backgroundColor = hoverBg; }}
          onMouseLeave={(e) => { (e.currentTarget as HTMLDivElement).style.backgroundColor = ''; }}
          onClick={() => {
            emitBlockInteraction({
              blockId,
              blockType: 'search-results',
              action: 'click-result',
              schemaId: 'search-results.click-result',
              payload: {
                index: itemNum,
                title: result.title,
                url: result.url,
                fileType: result.fileType,
                author: result.author,
                siteName: result.siteName
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
              onClick={(e) => {
                e.stopPropagation();
              }}
              onChange={(e) => {
                e.stopPropagation();
                if (!blockId) return;
                toggleActionSelection(blockId, itemNum);
              }}
            />
          </div>
          <div style={{ flex: 1, minWidth: 0, paddingRight: 8 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 2 }}>
              <Icon
                iconName={getFileIcon(result.fileType)}
                styles={{ root: { fontSize: 14, color: iconColor, flexShrink: 0 } }}
              />
              <a
                href={toWebViewerUrl(result.url)}
                target="_blank"
                rel="noopener noreferrer"
                data-interception="off"
                style={{
                  fontSize: 13,
                  fontWeight: 600,
                  color: titleLinkColor,
                  textDecoration: 'none',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                  flex: 1
                }}
                title={result.title}
              >
                {result.title}
              </a>
              {renderSourceShapes(result.sources)}
              {annotation && <span style={{ fontSize: 10, fontWeight: 600, color: '#4fc3f7', background: 'rgba(79,195,247,0.12)', borderRadius: 3, padding: '1px 5px', marginLeft: 6, whiteSpace: 'nowrap' }}>{annotation}</span>}
            </div>
            {/* Location breadcrumb */}
            <div style={{ fontSize: 11, color: veryDimText, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', marginBottom: 1 }}>
              {result.fileType && (
                <span style={{ color: fileTypeBadgeColor, fontWeight: 600, marginRight: 4, textTransform: 'uppercase' }}>
                  {result.fileType}
                </span>
              )}
              {result.siteName ? `${result.siteName} \u00b7 ` : ''}{extractUrlPath(result.url)}
            </div>
            {result.summary && result.summary.trim() && (
              <div style={{
                fontSize: 12,
                color: summaryColor,
                lineHeight: '1.4',
                display: '-webkit-box',
                WebkitLineClamp: 2,
                WebkitBoxOrient: 'vertical',
                overflow: 'hidden'
              }}>{result.summary}</div>
            )}
            <div style={{ fontSize: 11, color: dimText, marginTop: 3, display: 'flex', gap: 10, alignItems: 'center' }}>
              {result.author && <span>{result.author}</span>}
              {result.lastModified && <span>{new Date(result.lastModified).toLocaleDateString()}</span>}
            </div>
          </div>
        </div>
        );
      })}

      <div style={{ paddingTop: 6, fontSize: 10, color: dimText, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span>{totalCount} result{totalCount !== 1 ? 's' : ''}</span>
        <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
          {sources.map((s) => {
            const label = SOURCE_LABELS[s] || s;
            const colors = SOURCE_BADGE_COLORS[s] || { color: dimText, bg: badgeBg };
            return (
              <span
                key={s}
                style={{
                  fontSize: 10,
                  fontWeight: 600,
                  color: colors.color,
                  background: colors.bg,
                  borderRadius: 3,
                  padding: '2px 6px',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 4
                }}
              >
                {renderSourceShape(s, 10)}
                {label}
              </span>
            );
          })}
        </div>
      </div>
    </div>
  );
};
