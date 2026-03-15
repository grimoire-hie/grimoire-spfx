/**
 * ColorThemeSelector
 * Circular color theme buttons — replaces ethnicity-based appearance presets.
 * 7 color themes: blue, red, green, purple, gold, cyan, white.
 */

import * as React from 'react';
import { ColorTheme, COLOR_THEMES } from '../../services/avatar/FaceTemplateData';

export interface IColorThemeSelectorProps {
  current: ColorTheme;
  onChange: (theme: ColorTheme) => void;
}

const containerStyle: React.CSSProperties = {
  display: 'flex',
  gap: 8,
  padding: '8px 12px',
  backgroundColor: 'rgba(0, 0, 0, 0.4)',
  borderRadius: 20,
  backdropFilter: 'blur(8px)',
  zIndex: 10
};

const BUTTON_SIZE = 28;

const THEME_ORDER: ColorTheme[] = ['blue', 'red', 'green', 'purple', 'gold', 'cyan', 'white'];

export const ColorThemeSelector: React.FC<IColorThemeSelectorProps> = ({
  current,
  onChange
}) => {
  const labelStyle: React.CSSProperties = {
    fontSize: 10,
    color: 'rgba(255, 255, 255, 0.5)',
    textTransform: 'uppercase',
    letterSpacing: 1,
    fontWeight: 600
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <span style={labelStyle}>Theme</span>
      <div style={containerStyle}>
      {THEME_ORDER.map((themeId) => {
        const theme = COLOR_THEMES[themeId];
        const isActive = themeId === current;
        return (
          <button
            key={themeId}
            onClick={() => onChange(themeId)}
            title={theme.label}
            style={{
              width: BUTTON_SIZE,
              height: BUTTON_SIZE,
              borderRadius: '50%',
              border: isActive ? '2px solid white' : '2px solid rgba(255,255,255,0.2)',
              backgroundColor: theme.primaryColor,
              cursor: 'pointer',
              transition: 'all 0.3s ease',
              boxShadow: isActive ? `0 0 10px ${theme.primaryColor}` : 'none',
              padding: 0,
              outline: 'none',
              opacity: isActive ? 1 : 0.7
            }}
          />
        );
      })}
      </div>
    </div>
  );
};
