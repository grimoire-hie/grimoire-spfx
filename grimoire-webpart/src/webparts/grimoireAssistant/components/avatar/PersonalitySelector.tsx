/**
 * PersonalitySelector
 * 4 circular personality mode buttons in the bottom-right corner.
 * Each button is color-coded to its personality.
 */

import * as React from 'react';
import { PersonalityMode, PersonalityEngine } from '../../services/avatar/PersonalityEngine';

export interface IPersonalitySelectorProps {
  current: PersonalityMode;
  onChange: (mode: PersonalityMode) => void;
}

const containerStyle: React.CSSProperties = {
  position: 'absolute',
  bottom: 24,
  right: 24,
  display: 'flex',
  flexDirection: 'column',
  gap: 10,
  zIndex: 10
};

const BUTTON_SIZE = 36;

export const PersonalitySelector: React.FC<IPersonalitySelectorProps> = ({
  current,
  onChange
}) => {
  const modes = React.useMemo(() => PersonalityEngine.getModes(), []);

  const labelStyle: React.CSSProperties = {
    fontSize: 10,
    color: 'rgba(255, 255, 255, 0.5)',
    textTransform: 'uppercase',
    letterSpacing: 1,
    fontWeight: 600,
    textAlign: 'center'
  };

  return (
    <div style={containerStyle}>
      <span style={labelStyle}>Mode</span>
      {modes.map((mode) => {
        const isActive = mode.mode === current;
        return (
          <button
            key={mode.mode}
            onClick={() => onChange(mode.mode)}
            title={mode.label}
            style={{
              width: BUTTON_SIZE,
              height: BUTTON_SIZE,
              borderRadius: '50%',
              border: isActive ? '2px solid white' : '2px solid rgba(255,255,255,0.3)',
              backgroundColor: isActive ? mode.color : 'rgba(0,0,0,0.4)',
              cursor: 'pointer',
              transition: 'all 0.3s ease',
              boxShadow: isActive ? `0 0 12px ${mode.color}` : 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: 0,
              outline: 'none'
            }}
          >
            <div
              style={{
                width: BUTTON_SIZE - 12,
                height: BUTTON_SIZE - 12,
                borderRadius: '50%',
                backgroundColor: mode.color,
                opacity: isActive ? 1 : 0.6,
                transition: 'opacity 0.3s'
              }}
            />
          </button>
        );
      })}
    </div>
  );
};
