/**
 * TextInput
 * Always-visible text input field below the avatar.
 * Sends text to the voice session or triggers function calls.
 */

import * as React from 'react';
import { TextField, IconButton } from '@fluentui/react';
import * as strings from 'GrimoireAssistantWebPartStrings';
import { useGrimoireStore } from '../../store/useGrimoireStore';

export interface ITextInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: (text: string) => void;
  disabled?: boolean;
  placeholder?: string;
}

const MAX_PROMPT_HISTORY = 100;

const containerBaseStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '8px 12px',
  borderRadius: 24,
  width: '100%'
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

export const TextInput: React.FC<ITextInputProps> = ({
  value,
  onChange,
  onSend,
  disabled = false,
  placeholder = strings.TypeAMessage
}) => {
  const spThemeColors = useGrimoireStore((s) => s.spThemeColors);
  const promptHistoryRef = React.useRef<string[]>([]);
  const historyIndexRef = React.useRef<number>(-1);
  const draftBeforeHistoryRef = React.useRef<string>('');

  const containerStyle: React.CSSProperties = {
    ...containerBaseStyle,
    backgroundColor: spThemeColors.bodyBackground,
    border: `1px solid ${spThemeColors.cardBorder}`
  };

  const resetHistoryNavigation = React.useCallback((): void => {
    historyIndexRef.current = -1;
    draftBeforeHistoryRef.current = '';
  }, []);

  const appendPromptToHistory = React.useCallback((prompt: string): void => {
    promptHistoryRef.current.push(prompt);
    if (promptHistoryRef.current.length > MAX_PROMPT_HISTORY) {
      promptHistoryRef.current.splice(0, promptHistoryRef.current.length - MAX_PROMPT_HISTORY);
    }
  }, []);

  const submitCurrentPrompt = React.useCallback((): void => {
    const trimmed = value.trim();
    if (!trimmed) return;

    onSend(trimmed);
    onChange('');
    appendPromptToHistory(trimmed);
    resetHistoryNavigation();
  }, [appendPromptToHistory, onChange, onSend, resetHistoryNavigation, value]);

  const handleKeyDown = (ev: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
    if (ev.key === 'Enter' && !ev.shiftKey) {
      ev.preventDefault();
      submitCurrentPrompt();
      return;
    }

    if (ev.key === 'ArrowUp') {
      const promptHistory = promptHistoryRef.current;
      if (promptHistory.length === 0) return;

      ev.preventDefault();

      if (historyIndexRef.current === -1) {
        draftBeforeHistoryRef.current = value;
        historyIndexRef.current = promptHistory.length - 1;
      } else if (historyIndexRef.current > 0) {
        historyIndexRef.current -= 1;
      }

      onChange(promptHistory[historyIndexRef.current]);
      return;
    }

    if (ev.key === 'ArrowDown' && historyIndexRef.current !== -1) {
      const promptHistory = promptHistoryRef.current;
      ev.preventDefault();

      const nextIndex = historyIndexRef.current + 1;
      if (nextIndex < promptHistory.length) {
        historyIndexRef.current = nextIndex;
        onChange(promptHistory[nextIndex]);
      } else {
        historyIndexRef.current = -1;
        onChange(draftBeforeHistoryRef.current);
      }
    }
  };

  const handleChange = (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    if (historyIndexRef.current !== -1) {
      historyIndexRef.current = -1;
      draftBeforeHistoryRef.current = '';
    }
    onChange(newValue || '');
  };

  const handleSend = (): void => {
    submitCurrentPrompt();
  };

  return (
    <div style={containerStyle}>
      <TextField
        value={value}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
        placeholder={placeholder}
        disabled={disabled}
        borderless
        autoComplete="off"
        styles={{
          root: { flex: 1 },
          fieldGroup: {
            backgroundColor: 'transparent',
            border: 'none'
          },
          field: {
            color: spThemeColors.bodyText,
            fontSize: 14,
            '::placeholder': {
              color: withAlpha(spThemeColors.bodySubtext, 0.75)
            }
          }
        }}
      />
      <IconButton
        iconProps={{ iconName: 'Send' }}
        ariaLabel={strings.SendMessage}
        disabled={disabled || !value.trim()}
        onClick={handleSend}
        styles={{
          root: {
            color: value.trim() ? spThemeColors.bodyText : withAlpha(spThemeColors.bodySubtext, 0.65),
            backgroundColor: 'transparent',
            width: 32,
            height: 32
          },
          rootHovered: {
            backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.18 : 0.08),
            color: spThemeColors.bodyText
          }
        }}
      />
    </div>
  );
};
