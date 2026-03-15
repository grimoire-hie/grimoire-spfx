/**
 * TranscriptOverlay
 * Minimal floating transcript at the bottom of the virtual overlay.
 * Shows last few messages, fading older ones.
 */

import * as React from 'react';
import type { ITranscriptEntry } from '../../store/useGrimoireStore';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { TRANSCRIPT_OVERLAY_LENGTH_LIMITS } from '../../config/assistantLengthLimits';

export interface ITranscriptOverlayProps {
  entries: ITranscriptEntry[];
  /** Max entries to show (default 4) */
  maxVisible?: number;
  fadeOlderEntries?: boolean;
}

const MAX_OVERLAY_TEXT_CHARS = TRANSCRIPT_OVERLAY_LENGTH_LIMITS.textMaxChars;

function clipOverlayText(text: string): string {
  if (text.length <= MAX_OVERLAY_TEXT_CHARS) return text;
  return `${text.slice(0, MAX_OVERLAY_TEXT_CHARS - 1)}…`;
}

const containerStyle: React.CSSProperties = {
  width: '100%',
  display: 'flex',
  flexDirection: 'column',
  gap: 4,
  pointerEvents: 'none'
};

const bubbleBaseStyle: React.CSSProperties = {
  padding: '8px 14px',
  borderRadius: 12,
  fontSize: 13,
  lineHeight: '1.4',
  fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  maxWidth: '85%',
  backdropFilter: 'blur(8px)',
  transition: 'opacity 0.3s'
};

function withAlpha(color: string, alpha: number): string {
  const normalizedAlpha = Math.max(0, Math.min(1, alpha));
  const shortHex = /^#([0-9a-fA-F]{3})$/;
  const longHex = /^#([0-9a-fA-F]{6})$/;
  const shortMatch = color.match(shortHex);
  if (shortMatch) {
    const [r, g, b] = shortMatch[1].split('').map((part) => Number.parseInt(`${part}${part}`, 16));
    return `rgba(${r}, ${g}, ${b}, ${normalizedAlpha})`;
  }
  const longMatch = color.match(longHex);
  if (longMatch) {
    const raw = longMatch[1];
    const r = Number.parseInt(raw.slice(0, 2), 16);
    const g = Number.parseInt(raw.slice(2, 4), 16);
    const b = Number.parseInt(raw.slice(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${normalizedAlpha})`;
  }
  return color;
}

export const TranscriptOverlay: React.FC<ITranscriptOverlayProps> = ({
  entries,
  maxVisible = 4,
  fadeOlderEntries = true
}) => {
  const spThemeColors = useGrimoireStore((s) => s.spThemeColors);
  const visible = entries.slice(-maxVisible);
  const userBubbleStyle: React.CSSProperties = {
    ...bubbleBaseStyle,
    alignSelf: 'flex-end',
    backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.26 : 0.14),
    color: spThemeColors.bodyText,
    border: `1px solid ${spThemeColors.cardBorder}`,
    borderBottomRightRadius: 4
  };
  const assistantBubbleStyle: React.CSSProperties = {
    ...bubbleBaseStyle,
    alignSelf: 'flex-start',
    backgroundColor: withAlpha(spThemeColors.bodyBackground, spThemeColors.isDark ? 0.74 : 0.88),
    color: spThemeColors.bodyText,
    border: `1px solid ${spThemeColors.cardBorder}`,
    borderBottomLeftRadius: 4
  };
  const systemBubbleStyle: React.CSSProperties = {
    ...bubbleBaseStyle,
    alignSelf: 'center',
    backgroundColor: withAlpha('#40c040', spThemeColors.isDark ? 0.25 : 0.12),
    color: spThemeColors.bodyText,
    border: `1px solid ${withAlpha('#40c040', 0.4)}`,
    fontSize: 12,
    fontStyle: 'italic'
  };

  if (visible.length === 0) return null;

  return (
    <div style={containerStyle}>
      {visible.map((entry, index) => {
        // Fade older messages
        const age = visible.length - index;
        const opacity = !fadeOlderEntries
          ? 1
          : age <= 1 ? 1 : age <= 2 ? 0.85 : age <= 3 ? 0.7 : age <= 4 ? 0.55 : age <= 5 ? 0.4 : 0.3;

        let style: React.CSSProperties;
        switch (entry.role) {
          case 'user':
            style = { ...userBubbleStyle, opacity };
            break;
          case 'system':
            style = { ...systemBubbleStyle, opacity };
            break;
          default:
            style = { ...assistantBubbleStyle, opacity };
        }

        // Keep overlay lightweight during long streaming sessions.
        const displayText = clipOverlayText(entry.text);

        return (
          <div key={`${entry.timestamp.getTime()}-${index}`} style={style}>
            {displayText}
          </div>
        );
      })}
    </div>
  );
};
