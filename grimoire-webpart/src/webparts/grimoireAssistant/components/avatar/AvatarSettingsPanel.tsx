/**
 * AvatarSettingsPanel
 * In-pane settings drawer for avatar controls.
 */

import * as React from 'react';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { IconButton } from '@fluentui/react/lib/Button';
import * as strings from 'GrimoireAssistantWebPartStrings';
import type { IButtonStyles } from '@fluentui/react/lib/Button';
import type { IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import type { IToggleStyles } from '@fluentui/react/lib/Toggle';
import { shallow } from 'zustand/shallow';
import { useGrimoireStore, type ISpThemeColors } from '../../store/useGrimoireStore';
import { VISAGE_OPTIONS, VisageMode } from '../../services/avatar/FaceTemplateData';
import { getAvatarPersonaConfig } from '../../services/avatar/AvatarPersonaCatalog';
import { PERSONALITIES, PersonalityMode } from '../../services/avatar/PersonalityEngine';
import { SearchRecapMode } from '../../services/context/AssistantPreferenceUtils';
import { REALTIME_VOICE_OPTIONS } from '../../services/realtime/RealtimeVoiceCatalog';
const PERSONALITY_ORDER: PersonalityMode[] = ['normal', 'funny', 'harsh', 'devil'];

const personalityOptions: IDropdownOption[] = PERSONALITY_ORDER.map((id) => ({
  key: id,
  text: `${PERSONALITIES[id].label} — ${PERSONALITIES[id].description}`
}));

const visageOptions: IDropdownOption[] = (Object.keys(VISAGE_OPTIONS) as VisageMode[])
  .map((id) => ({
    key: id,
    text: VISAGE_OPTIONS[id].label
  }));

const searchRecapOptions: IDropdownOption[] = [
  { key: 'auto', text: strings.RecapOptionAuto },
  { key: 'always', text: strings.RecapOptionAlways },
  { key: 'off', text: strings.RecapOptionOff }
];

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

export function createAvatarSettingsDropdownStyles(spThemeColors: ISpThemeColors): Partial<IDropdownStyles> {
  return {
    root: { width: '100%' },
    title: {
      fontSize: 12,
      minHeight: 32,
      borderRadius: 4,
      backgroundColor: spThemeColors.bodyBackground,
      color: spThemeColors.bodyText,
      border: `1px solid ${spThemeColors.cardBorder}`,
      selectors: {
        ':hover': {
          backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.16 : 0.06)
        },
        ':focus': {
          backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.22 : 0.08)
        },
        '.ms-Dropdown-titleText': {
          color: `${spThemeColors.bodyText} !important`
        }
      }
    },
    caretDown: {
      color: spThemeColors.bodySubtext
    },
    callout: {
      backgroundColor: spThemeColors.cardBackground,
      border: `1px solid ${spThemeColors.cardBorder}`,
      color: spThemeColors.bodyText
    },
    dropdownItemsWrapper: {
      backgroundColor: spThemeColors.cardBackground,
      color: spThemeColors.bodyText
    },
    dropdownItem: {
      color: spThemeColors.bodyText,
      selectors: {
        ':hover': {
          color: spThemeColors.bodyText,
          backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.18 : 0.07)
        }
      }
    },
    dropdownItemSelected: {
      color: spThemeColors.bodyText,
      backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.24 : 0.12),
      selectors: {
        ':hover': {
          color: spThemeColors.bodyText,
          backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.3 : 0.16)
        }
      }
    }
  };
}

function createCloseButtonStyles(spThemeColors: ISpThemeColors): IButtonStyles {
  return {
    root: {
      width: 28,
      height: 28,
      minWidth: 28,
      color: spThemeColors.bodySubtext,
      backgroundColor: spThemeColors.bodyBackground,
      border: `1px solid ${spThemeColors.cardBorder}`,
      borderRadius: 4
    },
    rootHovered: {
      color: spThemeColors.bodyText,
      backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.18 : 0.08)
    }
  };
}

const sectionStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 6
};

function createToggleStyles(spThemeColors: ISpThemeColors): Partial<IToggleStyles> {
  return {
    root: { marginBottom: 0 },
    label: {
      color: spThemeColors.bodyText,
      fontSize: 12,
      fontWeight: 600,
      marginBottom: 8
    },
    text: {
      color: spThemeColors.bodySubtext,
      fontSize: 12
    }
  };
}

const hostStyle: React.CSSProperties = {
  position: 'absolute',
  top: 44,
  left: 0,
  right: 0,
  bottom: 0,
  zIndex: 20
};

const backdropBaseStyle: React.CSSProperties = {
  position: 'absolute',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0
};

const drawerBaseStyle: React.CSSProperties = {
  position: 'absolute',
  top: 0,
  right: 0,
  bottom: 0,
  width: 'min(340px, calc(100% - 12px))',
  display: 'flex',
  flexDirection: 'column'
};

const headerBaseStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  padding: '12px 12px 10px'
};

const bodyStyle: React.CSSProperties = {
  padding: '12px',
  overflowY: 'auto',
  display: 'flex',
  flexDirection: 'column',
  gap: 16
};

export interface IAvatarSettingsPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  voiceConnected?: boolean;
  onVoiceReconnect?: () => void;
}

export const AvatarSettingsPanel: React.FC<IAvatarSettingsPanelProps> = ({
  isOpen,
  onDismiss,
  voiceConnected = false,
  onVoiceReconnect
}) => {
  const {
    avatarEnabled,
    voiceId,
    publicWebSearchEnabled,
    publicWebSearchCapability,
    publicWebSearchCapabilityDetail,
    copilotWebGroundingEnabled,
    searchRecapMode,
    personality,
    visage,
    spThemeColors,
    markAvatarSettingsDirty,
    markAssistantSettingsDirty,
    setAvatarEnabled,
    setVoiceId,
    setPublicWebSearchEnabled,
    setCopilotWebGroundingEnabled,
    setSearchRecapMode,
    setPersonality,
    setVisage
  } = useGrimoireStore((s) => ({
    avatarEnabled: s.avatarEnabled,
    voiceId: s.voiceId,
    publicWebSearchEnabled: s.publicWebSearchEnabled,
    publicWebSearchCapability: s.publicWebSearchCapability,
    publicWebSearchCapabilityDetail: s.publicWebSearchCapabilityDetail,
    copilotWebGroundingEnabled: s.copilotWebGroundingEnabled,
    searchRecapMode: s.searchRecapMode,
    personality: s.personality,
    visage: s.visage,
    spThemeColors: s.spThemeColors,
    markAvatarSettingsDirty: s.markAvatarSettingsDirty,
    markAssistantSettingsDirty: s.markAssistantSettingsDirty,
    setAvatarEnabled: s.setAvatarEnabled,
    setVoiceId: s.setVoiceId,
    setPublicWebSearchEnabled: s.setPublicWebSearchEnabled,
    setCopilotWebGroundingEnabled: s.setCopilotWebGroundingEnabled,
    setSearchRecapMode: s.setSearchRecapMode,
    setPersonality: s.setPersonality,
    setVisage: s.setVisage
  }), shallow);

  const dropdownStyles = React.useMemo(
    () => createAvatarSettingsDropdownStyles(spThemeColors),
    [spThemeColors]
  );
  const closeButtonStyles = React.useMemo(
    () => createCloseButtonStyles(spThemeColors),
    [spThemeColors]
  );
  const toggleStyles = React.useMemo(
    () => createToggleStyles(spThemeColors),
    [spThemeColors]
  );
  const labelStyle: React.CSSProperties = {
    fontSize: 12,
    fontWeight: 600,
    color: spThemeColors.bodyText
  };
  const helperTextStyle: React.CSSProperties = {
    fontSize: 11,
    lineHeight: 1.4,
    color: spThemeColors.bodySubtext
  };
  const warningTextStyle: React.CSSProperties = {
    ...helperTextStyle,
    color: spThemeColors.isDark ? '#ffb3b3' : '#a4262c'
  };
  const backdropStyle: React.CSSProperties = {
    ...backdropBaseStyle,
    backgroundColor: withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.45 : 0.26)
  };
  const drawerStyle: React.CSSProperties = {
    ...drawerBaseStyle,
    backgroundColor: spThemeColors.cardBackground,
    borderLeft: `1px solid ${spThemeColors.cardBorder}`,
    boxShadow: `0 8px 28px ${withAlpha(spThemeColors.bodyText, spThemeColors.isDark ? 0.42 : 0.2)}`,
    color: spThemeColors.bodyText
  };
  const headerStyle: React.CSSProperties = {
    ...headerBaseStyle,
    borderBottom: `1px solid ${spThemeColors.cardBorder}`
  };
  const publicWebStatus = React.useMemo(() => {
    if (!publicWebSearchEnabled) {
      return {
        text: 'Research public websites and direct URLs through Azure OpenAI web search preview.',
        warning: false
      };
    }
    switch (publicWebSearchCapability) {
      case 'available':
        return {
          text: publicWebSearchCapabilityDetail || 'Public web research is available for this session.',
          warning: false
        };
      case 'blocked':
        return {
          text: publicWebSearchCapabilityDetail || 'Public web search is blocked for this tenant, subscription, or deployment.',
          warning: true
        };
      case 'unsupported':
        return {
          text: publicWebSearchCapabilityDetail || 'This Azure OpenAI deployment does not support web_search_preview.',
          warning: true
        };
      case 'error':
        return {
          text: publicWebSearchCapabilityDetail || 'The capability check failed. Public web search may be temporarily unavailable.',
          warning: true
        };
      default:
        return {
          text: 'Checking Azure public web search capability for this session...',
          warning: false
        };
    }
  }, [
    publicWebSearchCapability,
    publicWebSearchCapabilityDetail,
    publicWebSearchEnabled
  ]);
  const selectedVisage = VISAGE_OPTIONS[visage] ? visage : 'classic';
  const selectedVisagePersona = React.useMemo(
    () => getAvatarPersonaConfig(selectedVisage),
    [selectedVisage]
  );
  const visageHelperText = `${selectedVisagePersona.title} - ${selectedVisagePersona.uiDescription}`;

  if (!isOpen) return null;

  return (
    <div style={hostStyle}>
      <div style={backdropStyle} onClick={onDismiss} />
      <div style={drawerStyle}>
        <div style={headerStyle}>
          <span style={{ color: spThemeColors.bodyText, fontSize: 15, fontWeight: 600 }}>Avatar Settings</span>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            title={strings.CloseSettingsLabel}
            ariaLabel={strings.CloseSettingsLabel}
            onClick={onDismiss}
            styles={closeButtonStyles}
          />
        </div>

        <div style={bodyStyle}>
          <div style={sectionStyle}>
            <Toggle
              label={strings.AvatarToggleLabel}
              inlineLabel={false}
              checked={avatarEnabled}
              onText={strings.ToggleEnabled}
              offText={strings.ToggleDisabled}
              styles={toggleStyles}
              onChange={(_e, checked) => {
                markAvatarSettingsDirty();
                setAvatarEnabled(checked !== false);
              }}
            />
            <span style={helperTextStyle}>
              When disabled, Grimoire stops rendering the particle face and ignores expression and gaze cues. Voice, chat, and action results stay available.
            </span>
          </div>

          <div style={sectionStyle}>
            <span style={labelStyle}>Voice</span>
            <Dropdown
              selectedKey={voiceId}
              options={REALTIME_VOICE_OPTIONS}
              styles={dropdownStyles}
              onChange={(_e, option) => {
                if (!option) return;
                markAvatarSettingsDirty();
                setVoiceId(option.key as string);
                if (voiceConnected) {
                  onVoiceReconnect?.();
                }
              }}
            />
            <span style={helperTextStyle}>
              {voiceConnected
                ? strings.VoiceReconnectHint
                : strings.VoiceChangeNextConnect}
            </span>
          </div>

          <div style={sectionStyle}>
            <span style={labelStyle}>Personality</span>
            <Dropdown
              selectedKey={personality}
              options={personalityOptions}
              styles={dropdownStyles}
              onChange={(_e, option) => {
                if (!option) return;
                markAvatarSettingsDirty();
                setPersonality(option.key as PersonalityMode);
              }}
            />
          </div>

          <div style={sectionStyle}>
            <span style={labelStyle}>Visage</span>
            <Dropdown
              selectedKey={selectedVisage}
              options={visageOptions}
              styles={dropdownStyles}
              disabled={!avatarEnabled}
              onChange={(_e, option) => {
                if (!option) return;
                markAvatarSettingsDirty();
                setVisage(option.key as VisageMode);
              }}
            />
            <span style={helperTextStyle}>
              {visageHelperText}
            </span>
            {!avatarEnabled && (
              <span style={helperTextStyle}>
                Re-enable the avatar to change its visual style.
              </span>
            )}
          </div>

          <div style={sectionStyle}>
            <Toggle
              label={strings.PublicWebSearchLabel}
              inlineLabel={false}
              checked={publicWebSearchEnabled}
              onText={strings.ToggleEnabled}
              offText={strings.ToggleDisabled}
              styles={toggleStyles}
              onChange={(_e, checked) => {
                markAssistantSettingsDirty();
                setPublicWebSearchEnabled(!!checked);
              }}
            />
            <span style={publicWebStatus.warning ? warningTextStyle : helperTextStyle}>
              {publicWebStatus.text}
            </span>
          </div>

          <div style={sectionStyle}>
            <Toggle
              label={strings.CopilotWebGroundingLabel}
              inlineLabel={false}
              checked={copilotWebGroundingEnabled}
              onText={strings.ToggleEnabled}
              offText={strings.ToggleDisabled}
              styles={toggleStyles}
              onChange={(_e, checked) => {
                markAssistantSettingsDirty();
                setCopilotWebGroundingEnabled(!!checked);
              }}
            />
            <span style={helperTextStyle}>
              Lets M365 Copilot-backed reads use web grounding when the upstream service supports it.
            </span>
          </div>

          <div style={sectionStyle}>
            <span style={labelStyle}>Search Recap</span>
            <Dropdown
              selectedKey={searchRecapMode}
              options={searchRecapOptions}
              styles={dropdownStyles}
              onChange={(_e, option) => {
                if (!option) return;
                markAssistantSettingsDirty();
                setSearchRecapMode(option.key as SearchRecapMode);
              }}
            />
          </div>
        </div>
      </div>
    </div>
  );
};
