/**
 * MarkdownBlock
 * Renders markdown-like content with basic formatting.
 * Detects numbered item groups (emails, events, messages) and renders them as cards.
 * Uses simple regex-based parsing (no external markdown library).
 */

import * as React from 'react';
import { shallow } from 'zustand/shallow';
import { MARKDOWN_LENGTH_LIMITS } from '../../../config/assistantLengthLimits';
import type { IMarkdownData } from '../../../models/IBlock';
import { useGrimoireStore } from '../../../store/useGrimoireStore';
import { extractGroupFields } from '../selectionHelpers';

// ─── Shared styles ──────────────────────────────────────────────

const containerStyle: React.CSSProperties = {
  fontSize: 13,
  color: '#605e5c',
  lineHeight: '1.6'
};

const codeBlockStyle: React.CSSProperties = {
  display: 'block',
  backgroundColor: '#f3f2f1',
  border: '1px solid rgba(0, 0, 0, 0.08)',
  borderRadius: 4,
  padding: '8px 12px',
  fontFamily: 'Consolas, "Courier New", monospace',
  fontSize: 12,
  color: '#323130',
  overflowX: 'auto',
  whiteSpace: 'pre',
  margin: '8px 0'
};

const inlineCodeStyle: React.CSSProperties = {
  backgroundColor: 'rgba(0, 0, 0, 0.06)',
  borderRadius: 3,
  padding: '1px 4px',
  fontFamily: 'Consolas, "Courier New", monospace',
  fontSize: 12
};

const h1Style: React.CSSProperties = { fontSize: 16, fontWeight: 700, color: '#323130', margin: '12px 0 6px' };
const h2Style: React.CSSProperties = { fontSize: 14, fontWeight: 600, color: '#323130', margin: '10px 0 4px' };
const h3Style: React.CSSProperties = { fontSize: 13, fontWeight: 600, color: '#605e5c', margin: '8px 0 4px' };

const listStyle: React.CSSProperties = { paddingLeft: 20, margin: '4px 0' };
const hrStyle: React.CSSProperties = { border: 'none', borderTop: '1px solid rgba(0, 0, 0, 0.08)', margin: '12px 0' };
const linkStyle: React.CSSProperties = { color: '#0064b4', textDecoration: 'none' };

// ─── Card styles for numbered item groups ────────────────────────

const cardStyle: React.CSSProperties = {
  backgroundColor: 'rgba(0, 0, 0, 0.02)',
  borderLeft: '3px solid rgba(0, 120, 212, 0.3)',
  borderRadius: '0 6px 6px 0',
  padding: '10px 14px',
  margin: '8px 0'
};

const numBadgeStyle: React.CSSProperties = {
  display: 'inline-block',
  backgroundColor: 'rgba(0, 120, 212, 0.08)',
  color: '#0064b4',
  fontWeight: 600,
  fontSize: 11,
  borderRadius: 10,
  padding: '1px 8px',
  marginBottom: 6
};

const fieldRowStyle: React.CSSProperties = {
  margin: '2px 0',
  fontSize: 13,
  lineHeight: '1.5'
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
  paddingTop: 2,
  flexShrink: 0
};

// ─── Numbered item group detection ──────────────────────────────

interface INumberedGroup {
  num: number;
  lines: string[];
}

interface IGroupedContent {
  preamble: string[];
  groups: INumberedGroup[];
}

/**
 * Try to split markdown into preamble + numbered item groups.
 * Returns null if fewer than 2 groups found (not worth card rendering).
 */
function splitIntoNumberedGroups(content: string): IGroupedContent | undefined {
  const lines = content.split('\n');
  const preamble: string[] = [];
  const groups: INumberedGroup[] = [];
  let currentGroup: INumberedGroup | null = null;
  let foundFirstNumber = false;

  for (let li = 0; li < lines.length; li++) {
    const trimmed = lines[li].trim();

    // Numbered item start: "1." or "1)" alone or "1. Text after"
    const numMatch = /^(\d+)[.)]\s*(.*)$/.exec(trimmed);
    if (numMatch) {
      foundFirstNumber = true;
      if (currentGroup) groups.push(currentGroup);
      currentGroup = { num: parseInt(numMatch[1], 10), lines: [] };
      if (numMatch[2]) currentGroup.lines.push(numMatch[2]);
      continue;
    }

    // Skip --- separators everywhere
    if (/^[-*_]{3,}$/.test(trimmed)) continue;

    if (!foundFirstNumber) {
      preamble.push(lines[li]);
    } else if (currentGroup) {
      currentGroup.lines.push(lines[li]);
    }
  }

  if (currentGroup) groups.push(currentGroup);
  return groups.length >= 2 ? { preamble, groups } : undefined;
}

// ─── Card styles (interactive) ──────────────────────────────────

const cardHoverStyle: React.CSSProperties = {
  ...cardStyle,
  cursor: 'pointer',
  transition: 'background-color 0.15s ease, border-left-color 0.15s ease'
};

// ─── Card-based renderer ────────────────────────────────────────

function renderGroupedContent(
  grouped: IGroupedContent,
  blockId?: string,
  selection?: {
    activeBlockId?: string;
    selectedSet: Set<number>;
    onToggle: (index: number) => void;
  }
): React.ReactNode[] {
  const elements: React.ReactNode[] = [];
  let key = 0;

  // Render preamble lines
  grouped.preamble.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    elements.push(<div key={key++} style={{ marginBottom: 4 }}>{formatInline(trimmed)}</div>);
  });

  // Render each numbered group as a card with hover actions
  grouped.groups.forEach((group) => {
    const fields = extractGroupFields(group.lines);
    const title = fields.Subject || fields.Event || fields.Message || fields.From
      || fields.Title || Object.values(fields)[0] || 'Item';
    const canSelect = !!blockId && selection?.activeBlockId === blockId;
    const isSelected = !!canSelect && selection?.selectedSet.has(group.num);
    elements.push(
      <div
        key={key++}
        className="grim-md-card"
        title="Use checkbox to select items"
        style={{ ...cardHoverStyle, ...(isSelected ? { backgroundColor: 'rgba(0, 100, 180, 0.1)' } : {}) }}
        onMouseEnter={(e) => { (e.currentTarget as HTMLDivElement).style.backgroundColor = 'rgba(0, 120, 212, 0.06)'; }}
        onMouseLeave={(e) => {
          (e.currentTarget as HTMLDivElement).style.backgroundColor = isSelected ? 'rgba(0, 100, 180, 0.1)' : 'rgba(0, 0, 0, 0.02)';
        }}
      >
        <div style={{ display: 'flex', alignItems: 'stretch' }}>
          <div style={selectionColumnStyle}>
            <input
              type="checkbox"
              checked={isSelected}
              disabled={!canSelect}
              style={{ ...checkboxStyle, opacity: canSelect ? 1 : 0.45, cursor: canSelect ? 'pointer' : 'default' }}
              onClick={(e) => { e.stopPropagation(); }}
              onChange={(e) => {
                e.stopPropagation();
                selection?.onToggle(group.num);
              }}
            />
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 4 }}>
              <div style={numBadgeStyle}>{group.num}</div>
              <span style={{ fontSize: 11, color: '#a19f9d', fontWeight: 600, marginLeft: 8, whiteSpace: 'nowrap' }}>
                {title}
              </span>
            </div>
            {group.lines.map((line, i) => {
              const trimmed = line.trim();
              if (!trimmed) return null;
              return <div key={i} style={fieldRowStyle}>{formatInline(trimmed)}</div>;
            })}
          </div>
        </div>
      </div>
    );
  });

  return elements;
}

// ─── Line-by-line fallback renderer ─────────────────────────────

/**
 * Simple line-by-line markdown renderer.
 * Handles: headers, bold, italic, inline code, code blocks, lists, links, HR.
 */
function renderMarkdown(content: string): React.ReactNode[] {
  const lines = content.split('\n');
  const elements: React.ReactNode[] = [];
  let inCodeBlock = false;
  let codeBuffer: string[] = [];
  let key = 0;

  for (const line of lines) {
    // Code blocks
    if (line.trim().startsWith('```')) {
      if (inCodeBlock) {
        elements.push(<pre key={key++} style={codeBlockStyle}>{codeBuffer.join('\n')}</pre>);
        codeBuffer = [];
        inCodeBlock = false;
      } else {
        inCodeBlock = true;
      }
      continue;
    }
    if (inCodeBlock) {
      codeBuffer.push(line);
      continue;
    }

    const trimmed = line.trim();

    // Empty line
    if (!trimmed) {
      elements.push(<div key={key++} style={{ height: 6 }} />);
      continue;
    }

    // Horizontal rule
    if (/^[-*_]{3,}$/.test(trimmed)) {
      elements.push(<hr key={key++} style={hrStyle} />);
      continue;
    }

    // Headers
    if (trimmed.startsWith('### ')) {
      elements.push(<div key={key++} style={h3Style}>{formatInline(trimmed.slice(4))}</div>);
      continue;
    }
    if (trimmed.startsWith('## ')) {
      elements.push(<div key={key++} style={h2Style}>{formatInline(trimmed.slice(3))}</div>);
      continue;
    }
    if (trimmed.startsWith('# ')) {
      elements.push(<div key={key++} style={h1Style}>{formatInline(trimmed.slice(2))}</div>);
      continue;
    }

    // List items
    if (/^[-*+]\s/.test(trimmed)) {
      elements.push(
        <div key={key++} style={listStyle}>
          <span style={{ marginRight: 4 }}>&#x2022;</span>
          {formatInline(trimmed.slice(2))}
        </div>
      );
      continue;
    }

    // Numbered list
    if (/^\d+\.\s/.test(trimmed)) {
      const match = trimmed.match(/^(\d+)\.\s(.*)$/);
      if (match) {
        elements.push(
          <div key={key++} style={listStyle}>
            <span style={{ marginRight: 4 }}>{match[1]}.</span>
            {formatInline(match[2])}
          </div>
        );
        continue;
      }
    }

    // Regular paragraph
    elements.push(<div key={key++}>{formatInline(trimmed)}</div>);
  }

  // Flush any remaining code block
  if (inCodeBlock && codeBuffer.length > 0) {
    elements.push(<pre key={key++} style={codeBlockStyle}>{codeBuffer.join('\n')}</pre>);
  }

  return elements;
}

// ─── Inline formatting ──────────────────────────────────────────

/** Format inline markdown: bold, italic, inline code, links */
function formatInline(text: string): React.ReactNode {
  // Split by inline code first to protect code content
  const parts = text.split(/(`[^`]+`)/g);
  const result: React.ReactNode[] = [];

  parts.forEach((part, idx) => {
    if (part.startsWith('`') && part.endsWith('`')) {
      result.push(<code key={idx} style={inlineCodeStyle}>{part.slice(1, -1)}</code>);
    } else {
      // Process bold, italic, links in non-code parts
      let processed = part;
      // Bold
      processed = processed.replace(/\*\*([^*]+)\*\*/g, '|||BOLD_START|||$1|||BOLD_END|||');
      // Italic
      processed = processed.replace(/\*([^*]+)\*/g, '|||ITALIC_START|||$1|||ITALIC_END|||');
      // Links [text](url)
      processed = processed.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '|||LINK_START|||$1|||LINK_SEP|||$2|||LINK_END|||');

      // Convert to React nodes
      const fragments = processed.split(/(\|\|\|(?:BOLD|ITALIC|LINK)_(?:START|END|SEP)\|\|\|)/g);
      let i = 0;
      while (i < fragments.length) {
        const frag = fragments[i];
        if (frag === '|||BOLD_START|||') {
          result.push(<strong key={`${idx}-${i}`}>{fragments[i + 1]}</strong>);
          i += 3; // skip content + END
        } else if (frag === '|||ITALIC_START|||') {
          result.push(<em key={`${idx}-${i}`}>{fragments[i + 1]}</em>);
          i += 3;
        } else if (frag === '|||LINK_START|||') {
          const linkText = fragments[i + 1];
          const linkUrl = fragments[i + 3]; // skip SEP marker
          result.push(
            <a key={`${idx}-${i}`} href={linkUrl} target="_blank" rel="noopener noreferrer" data-interception="off" style={linkStyle}>
              {linkText}
            </a>
          );
          i += 5; // text + SEP marker + url + END
        } else if (frag && frag.indexOf('|||') === -1) {
          result.push(frag);
          i++;
        } else {
          i++;
        }
      }
    }
  });

  return <>{result}</>;
}

// ─── Component ──────────────────────────────────────────────────

const MAX_RENDER_CHARS = MARKDOWN_LENGTH_LIMITS.renderMaxChars;

export const MarkdownBlock: React.FC<{ data: IMarkdownData; blockId?: string }> = ({ data, blockId }) => {
  const { activeActionBlockId, selectedActionIndices, toggleActionSelection } = useGrimoireStore((s) => ({
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    toggleActionSelection: s.toggleActionSelection
  }), shallow);
  const selectedSet = React.useMemo(() => new Set<number>(selectedActionIndices), [selectedActionIndices]);

  // Cap content to prevent browser freeze on massive MCP replies
  const rawLen = data.content.length;
  const content = rawLen > MAX_RENDER_CHARS
    ? data.content.substring(0, MAX_RENDER_CHARS)
    : data.content;
  const truncated = rawLen > MAX_RENDER_CHARS;

  // Try card layout for numbered item groups (emails, events, messages)
  const grouped = React.useMemo(() => splitIntoNumberedGroups(content), [content]);

  return (
    <div style={containerStyle}>
      {grouped ? renderGroupedContent(grouped, blockId, {
        activeBlockId: activeActionBlockId,
        selectedSet,
        onToggle: (index) => {
          if (!blockId) return;
          toggleActionSelection(blockId, index);
        }
      }) : renderMarkdown(content)}
      {truncated && (
        <div style={{ fontSize: 11, color: '#a19f9d', fontStyle: 'italic', marginTop: 8 }}>
          Content truncated ({Math.round(rawLen / 1000)}K chars total)
        </div>
      )}
    </div>
  );
};
