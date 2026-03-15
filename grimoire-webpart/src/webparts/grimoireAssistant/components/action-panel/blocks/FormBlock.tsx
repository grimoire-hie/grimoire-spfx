/**
 * FormBlock
 * Generic form renderer for write/create MCP operations.
 * Supports ~10 field types, validation, email tag chips, submit/cancel.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import * as strings from 'GrimoireAssistantWebPartStrings';
import type { ITag } from '@fluentui/react/lib/Pickers';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
import { TeamPicker } from '@pnp/spfx-controls-react/lib/TeamPicker';
import { TeamChannelPicker } from '@pnp/spfx-controls-react/lib/TeamChannelPicker';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import type { IFormData, IFormFieldDefinition } from '../../../models/IBlock';
import { emitBlockInteraction } from '../interactionSchemas';
import { getContext } from '../../../services/pnp/pnpContext';

// ─── Styles ──────────────────────────────────────────────────────

const descStyle: React.CSSProperties = {
  fontSize: 12,
  color: '#a19f9d',
  marginBottom: 12
};

const groupLabelStyle: React.CSSProperties = {
  fontSize: 12,
  fontWeight: 600,
  color: '#605e5c',
  marginTop: 12,
  marginBottom: 6,
  textTransform: 'uppercase' as const,
  letterSpacing: 0.5
};

const fieldLabelStyle: React.CSSProperties = {
  fontSize: 12,
  color: '#323130',
  marginBottom: 3,
  display: 'flex',
  alignItems: 'center',
  gap: 3
};

const requiredStarStyle: React.CSSProperties = {
  color: '#e06060',
  fontSize: 12
};

const inputStyle: React.CSSProperties = {
  width: '100%',
  padding: '7px 10px',
  borderRadius: 6,
  border: '1px solid rgba(0, 0, 0, 0.15)',
  fontSize: 13,
  color: '#323130',
  background: '#fff',
  outline: 'none',
  boxSizing: 'border-box' as const
};

const inputErrorStyle: React.CSSProperties = {
  ...inputStyle,
  borderColor: '#e06060'
};

const textareaStyle: React.CSSProperties = {
  ...inputStyle,
  resize: 'vertical' as const,
  fontFamily: 'inherit'
};

const errorTextStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#e06060',
  marginTop: 2
};

const tagContainerStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap' as const,
  gap: 4,
  marginBottom: 4
};

const tagStyle: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 4,
  padding: '2px 8px',
  borderRadius: 12,
  background: 'rgba(0, 100, 180, 0.1)',
  color: '#0064b4',
  fontSize: 12
};

const tagRemoveStyle: React.CSSProperties = {
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center'
};

const toggleContainerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  marginBottom: 4
};

const toggleTrackStyle: React.CSSProperties = {
  width: 36,
  height: 20,
  borderRadius: 10,
  background: 'rgba(0, 0, 0, 0.15)',
  position: 'relative' as const,
  cursor: 'pointer',
  transition: 'background 0.15s ease',
  flexShrink: 0
};

const toggleTrackOnStyle: React.CSSProperties = {
  ...toggleTrackStyle,
  background: 'rgba(0, 100, 180, 0.4)'
};

const toggleThumbStyle: React.CSSProperties = {
  width: 16,
  height: 16,
  borderRadius: '50%',
  background: '#fff',
  position: 'absolute' as const,
  top: 2,
  left: 2,
  transition: 'left 0.15s ease',
  boxShadow: '0 1px 3px rgba(0,0,0,0.2)'
};

const toggleThumbOnStyle: React.CSSProperties = {
  ...toggleThumbStyle,
  left: 18
};

const halfRowStyle: React.CSSProperties = {
  display: 'flex',
  gap: 10
};

const halfFieldStyle: React.CSSProperties = {
  flex: 1,
  minWidth: 0
};

const buttonRowStyle: React.CSSProperties = {
  display: 'flex',
  gap: 10,
  marginTop: 16
};

const submitButtonStyle: React.CSSProperties = {
  flex: 1,
  padding: '9px 20px',
  borderRadius: 6,
  border: 'none',
  background: 'rgba(0, 100, 180, 0.15)',
  color: '#0064b4',
  fontSize: 13,
  fontWeight: 600,
  cursor: 'pointer'
};

const cancelButtonStyle: React.CSSProperties = {
  padding: '9px 20px',
  borderRadius: 6,
  border: '1px solid rgba(0, 0, 0, 0.15)',
  background: 'transparent',
  color: '#605e5c',
  fontSize: 13,
  cursor: 'pointer'
};

const overlayStyle: React.CSSProperties = {
  textAlign: 'center' as const,
  padding: '16px 0'
};

const successStyle: React.CSSProperties = {
  ...overlayStyle,
  color: '#107c10'
};

const errorMsgStyle: React.CSSProperties = {
  ...overlayStyle,
  color: '#e06060'
};

const messageLinkStyle: React.CSSProperties = {
  color: '#0f6cbd',
  textDecoration: 'underline',
  wordBreak: 'break-all' as const
};

const attachmentSectionStyle: React.CSSProperties = {
  marginBottom: 12
};

const attachmentListStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column' as const,
  gap: 4
};

const attachmentRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8
};

const attachmentOpenButtonStyle: React.CSSProperties = {
  flex: 1,
  padding: 0,
  border: 'none',
  background: 'transparent',
  color: '#0f6cbd',
  textDecoration: 'underline',
  fontSize: 12,
  wordBreak: 'break-all' as const,
  textAlign: 'left' as const,
  cursor: 'pointer'
};

const attachmentRemoveButtonStyle: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  minHeight: 24,
  padding: '0 8px',
  borderRadius: 12,
  border: '1px solid rgba(0, 0, 0, 0.15)',
  background: '#fff',
  color: '#605e5c',
  cursor: 'pointer',
  flexShrink: 0
};

// ─── Email validation ────────────────────────────────────────────

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const URL_PATTERN = /(https?:\/\/[^\s]+)/gi;

function getAttachmentDisplayName(uri: string): string {
  const trimmed = uri.trim();
  if (!trimmed) {
    return '';
  }

  try {
    const parsed = new URL(trimmed);
    const pathSegments = parsed.pathname.split('/').filter(Boolean);
    return decodeURIComponent(pathSegments[pathSegments.length - 1] || trimmed);
  } catch {
    return trimmed;
  }
}

function openAttachmentInNewTab(uri: string): void {
  if (typeof window === 'undefined' || typeof window.open !== 'function') {
    return;
  }
  window.open(uri, '_blank', 'noopener,noreferrer');
}

function normalizePeoplePickerRecipients(items: IPersonaProps[] | undefined): string[] {
  if (!items || items.length === 0) {
    return [];
  }

  const seen = new Set<string>();
  const recipients: string[] = [];
  items.forEach((item) => {
    const candidate = [item.secondaryText, item.tertiaryText, item.optionalText, item.text]
      .find((value) => typeof value === 'string' && value.trim().length > 0)
      ?.trim();
    if (!candidate) {
      return;
    }
    const normalized = candidate.toLowerCase();
    if (seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    recipients.push(candidate);
  });
  return recipients;
}

// ─── Time slots for datetime picker ─────────────────────────────

function generateTimeSlots(): string[] {
  const slots: string[] = [];
  for (let h = 0; h < 24; h++) {
    for (let m = 0; m < 60; m += 30) {
      const hh = h < 10 ? '0' + h : '' + h;
      const mm = m < 10 ? '0' + m : '' + m;
      slots.push(hh + ':' + mm);
    }
  }
  return slots;
}

const TIME_SLOTS = generateTimeSlots();

function isFieldVisible(field: IFormFieldDefinition, values: Record<string, string>): boolean {
  if (field.type === 'hidden') {
    return false;
  }

  if (!field.visibleWhen) {
    return true;
  }

  const currentValue = values[field.visibleWhen.fieldKey] || '';
  if (field.visibleWhen.equals !== undefined) {
    return currentValue === field.visibleWhen.equals;
  }
  if (field.visibleWhen.anyOf && field.visibleWhen.anyOf.length > 0) {
    return field.visibleWhen.anyOf.indexOf(currentValue) !== -1;
  }

  return true;
}

// ─── Component ──────────────────────────────────────────────────

export interface IFormBlockProps {
  data: IFormData;
  blockId?: string;
  /** Execute the MCP tool from form data. Provided by ActionPanel. */
  onSubmit?: (formData: IFormData, fieldValues: Record<string, string>, emailTags: Record<string, string[]>) => Promise<{ success: boolean; message: string }>;
  /** Update the block in the store (e.g. status changes). Provided by ActionPanel. */
  onUpdateBlock?: (blockId: string, updates: { data: IFormData }) => void;
}

export const FormBlock: React.FC<IFormBlockProps> = ({ data, blockId, onSubmit, onUpdateBlock }) => {
  const { fields, description, status: initialStatus, errorMessage, successMessage } = data;

  // State
  const [values, setValues] = React.useState<Record<string, string>>(() => {
    const init: Record<string, string> = {};
    fields.forEach((f) => {
      init[f.key] = f.defaultValue || '';
    });
    return init;
  });
  const [emailTags, setEmailTags] = React.useState<Record<string, string[]>>(() => {
    const init: Record<string, string[]> = {};
    fields.forEach((f) => {
      if ((f.type === 'email-list' || f.type === 'people-picker') && f.defaultValue) {
        // Parse comma-separated default values into tags
        init[f.key] = f.defaultValue.split(',').map((e) => e.trim()).filter(Boolean);
      } else if (f.type === 'email-list' || f.type === 'people-picker') {
        init[f.key] = [];
      }
    });
    return init;
  });
  const [errors, setErrors] = React.useState<Record<string, string>>({});
  const [formStatus, setFormStatus] = React.useState<string>(initialStatus);
  const initialAttachmentUris = React.useMemo(() => (
    Array.isArray(data.submissionTarget.staticArgs?.attachmentUris)
      ? data.submissionTarget.staticArgs.attachmentUris.filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
      : []
  ), [data.submissionTarget.staticArgs]);
  const [attachmentUris, setAttachmentUris] = React.useState<string[]>(initialAttachmentUris);
  const pickerContext = React.useMemo(() => {
    try {
      return getContext();
    } catch {
      return undefined;
    }
  }, []);
  const pickerAppContext = pickerContext as unknown as React.ComponentProps<typeof TeamPicker>['appcontext'] | undefined;
  const peoplePickerContext = React.useMemo(() => {
    if (!pickerContext) {
      return undefined;
    }
    return {
      absoluteUrl: pickerContext.pageContext.web.absoluteUrl,
      msGraphClientFactory: pickerContext.msGraphClientFactory,
      spHttpClient: pickerContext.spHttpClient
    } as unknown as React.ComponentProps<typeof PeoplePicker>['context'];
  }, [pickerContext]);
  const selectedTeams = React.useMemo<ITag[]>(() => {
    const teamId = (values.teamId || '').trim();
    const teamName = (values.teamName || '').trim();
    if (!teamId || !teamName) {
      return [];
    }
    return [{ key: teamId, name: teamName }];
  }, [values.teamId, values.teamName]);
  const selectedChannels = React.useMemo<ITag[]>(() => {
    const channelId = (values.channelId || '').trim();
    const channelName = (values.channelName || '').trim();
    if (!channelId || !channelName) {
      return [];
    }
    return [{ key: channelId, name: channelName }];
  }, [values.channelId, values.channelName]);

  // Sync external status changes
  React.useEffect(() => {
    setFormStatus(initialStatus);
  }, [initialStatus]);

  React.useEffect(() => {
    setAttachmentUris(initialAttachmentUris);
  }, [initialAttachmentUris]);

  const isDisabled = formStatus === 'submitting' || formStatus === 'success';

  const buildSubmissionData = React.useCallback((status?: IFormData['status'], message?: string): IFormData => {
    const nextStaticArgs = { ...data.submissionTarget.staticArgs };
    if (attachmentUris.length > 0) {
      nextStaticArgs.attachmentUris = attachmentUris;
    } else {
      delete nextStaticArgs.attachmentUris;
    }

    return {
      ...data,
      status: status || data.status,
      errorMessage: status === 'error' ? message : data.errorMessage,
      successMessage: status === 'success' ? message : data.successMessage,
      submissionTarget: {
        ...data.submissionTarget,
        staticArgs: nextStaticArgs
      }
    };
  }, [attachmentUris, data]);

  const removeAttachment = React.useCallback((uri: string) => {
    if (isDisabled) return;
    setAttachmentUris((prev) => {
      const next = prev.filter((value) => value !== uri);
      if (onUpdateBlock && blockId) {
        onUpdateBlock(blockId, { data: {
          ...buildSubmissionData(),
          submissionTarget: {
            ...data.submissionTarget,
            staticArgs: next.length > 0
              ? { ...data.submissionTarget.staticArgs, attachmentUris: next }
              : Object.fromEntries(
                Object.entries(data.submissionTarget.staticArgs).filter(([key]) => key !== 'attachmentUris')
              )
          }
        } });
      }
      return next;
    });
  }, [blockId, buildSubmissionData, data.submissionTarget, isDisabled, onUpdateBlock]);

  // ─── Email Tag Handlers ─────────────────────────────────────

  const addEmailTag = React.useCallback((fieldKey: string, email: string) => {
    const trimmed = email.trim();
    if (!trimmed) return;
    if (!EMAIL_REGEX.test(trimmed)) {
      setErrors((prev) => ({ ...prev, [fieldKey]: 'Invalid email address' }));
      return;
    }
    setErrors((prev) => { const n = { ...prev }; delete n[fieldKey]; return n; });
    setEmailTags((prev) => {
      const existing = prev[fieldKey] || [];
      if (existing.indexOf(trimmed) !== -1) return prev;
      return { ...prev, [fieldKey]: [...existing, trimmed] };
    });
    setValues((prev) => ({ ...prev, [fieldKey]: '' }));
  }, []);

  const removeEmailTag = React.useCallback((fieldKey: string, email: string) => {
    setEmailTags((prev) => ({
      ...prev,
      [fieldKey]: (prev[fieldKey] || []).filter((e) => e !== email)
    }));
  }, []);

  // ─── Validation ─────────────────────────────────────────────

  const validate = React.useCallback((): boolean => {
    const newErrors: Record<string, string> = {};
    fields.forEach((f) => {
      if (!isFieldVisible(f, values)) return;
      if (f.required) {
        if (f.type === 'email-list' || f.type === 'people-picker') {
          if (!emailTags[f.key] || emailTags[f.key].length === 0) {
            newErrors[f.key] = f.type === 'people-picker'
              ? 'At least one person is required'
              : 'At least one email is required';
          }
        } else if (!values[f.key] || values[f.key].trim() === '') {
          newErrors[f.key] = 'This field is required';
        }
      }
      if (f.type === 'email' && values[f.key] && !EMAIL_REGEX.test(values[f.key])) {
        newErrors[f.key] = 'Invalid email address';
      }
    });
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }, [fields, values, emailTags]);

  // ─── Submit ─────────────────────────────────────────────────

  const handleSubmit = React.useCallback(() => {
    if (!validate() || !onSubmit) return;

    setFormStatus('submitting');
    const submissionData = buildSubmissionData('submitting');

    onSubmit(submissionData, values, emailTags).then((result) => {
      if (result.success) {
        setFormStatus('success');
        if (onUpdateBlock && blockId) {
          onUpdateBlock(blockId, { data: buildSubmissionData('success', result.message) });
        }
      } else {
        setFormStatus('error');
        if (onUpdateBlock && blockId) {
          onUpdateBlock(blockId, { data: buildSubmissionData('error', result.message) });
        }
      }

      // Notify HIE about the submission
      const fieldSummary = buildFieldSummary(submissionData, values, emailTags);
      emitBlockInteraction({
        blockId,
        blockType: 'form',
        action: 'submit-form',
        schemaId: 'form.submit',
        payload: {
          preset: submissionData.preset,
          toolName: submissionData.submissionTarget.toolName,
          success: result.success,
          fields: fieldSummary,
          message: result.message
        },
        timestamp: Date.now()
      });
    }).catch((err: Error) => {
      setFormStatus('error');
      if (onUpdateBlock && blockId) {
        onUpdateBlock(blockId, { data: buildSubmissionData('error', err.message || 'Form submission failed.') });
      }
      const fieldSummary = buildFieldSummary(submissionData, values, emailTags);
      emitBlockInteraction({
        blockId,
        blockType: 'form',
        action: 'submit-form',
        schemaId: 'form.submit',
        payload: {
          preset: submissionData.preset,
          toolName: submissionData.submissionTarget.toolName,
          success: false,
          fields: fieldSummary,
          message: err.message || 'Form submission failed.'
        },
        timestamp: Date.now()
      });
    });
  }, [blockId, buildSubmissionData, emailTags, onSubmit, onUpdateBlock, validate, values]);

  // ─── Cancel ─────────────────────────────────────────────────

  const handleCancel = React.useCallback(() => {
    emitBlockInteraction({
      blockId,
      blockType: 'form',
      action: 'cancel-form',
      schemaId: 'form.cancel',
      payload: { preset: data.preset },
      timestamp: Date.now()
    });
  }, [blockId, data.preset]);

  // ─── Render Helpers ─────────────────────────────────────────

  const renderTagInput = React.useCallback((field: IFormFieldDefinition, currentInputStyle: React.CSSProperties): React.ReactElement => {
    const tags = emailTags[field.key] || [];
    return (
      <div>
        {tags.length > 0 && (
          <div style={tagContainerStyle}>
            {tags.map((email) => (
              <span key={email} style={tagStyle}>
                {email}
                {!isDisabled && (
                  <span style={tagRemoveStyle} onClick={() => removeEmailTag(field.key, email)}>
                    <Icon iconName="Cancel" styles={{ root: { fontSize: 10 } }} />
                  </span>
                )}
              </span>
            ))}
          </div>
        )}
        <input
          style={currentInputStyle}
          placeholder={field.placeholder}
          value={values[field.key] || ''}
          disabled={isDisabled}
          onChange={(e) => {
            const val = e.target.value;
            if (val.indexOf(',') !== -1) {
              const parts = val.split(',');
              parts.forEach((part, idx) => {
                if (idx < parts.length - 1 && part.trim()) {
                  addEmailTag(field.key, part);
                }
              });
              setValues((prev) => ({ ...prev, [field.key]: parts[parts.length - 1] }));
            } else {
              setValues((prev) => ({ ...prev, [field.key]: val }));
            }
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              addEmailTag(field.key, values[field.key] || '');
            }
            if (e.key === 'Backspace' && !values[field.key]) {
              const tags2 = emailTags[field.key] || [];
              if (tags2.length > 0) {
                removeEmailTag(field.key, tags2[tags2.length - 1]);
              }
            }
          }}
          onBlur={() => {
            if (values[field.key]?.trim()) {
              addEmailTag(field.key, values[field.key]);
            }
          }}
        />
      </div>
    );
  }, [addEmailTag, emailTags, isDisabled, removeEmailTag, values]);

  const renderField = (field: IFormFieldDefinition): React.ReactElement | null => {
    if (field.type === 'hidden') return null;

    const hasError = !!errors[field.key];
    const currentInputStyle = hasError ? inputErrorStyle : inputStyle;

    switch (field.type) {
      case 'text':
      case 'email':
      case 'number':
        return (
          <input
            style={currentInputStyle}
            type={field.type === 'number' ? 'number' : 'text'}
            placeholder={field.placeholder}
            value={values[field.key] || ''}
            disabled={isDisabled}
            onChange={(e) => {
              setValues((prev) => ({ ...prev, [field.key]: e.target.value }));
              if (errors[field.key]) {
                setErrors((prev) => { const n = { ...prev }; delete n[field.key]; return n; });
              }
            }}
          />
        );

      case 'textarea':
        return (
          <textarea
            style={{ ...textareaStyle, borderColor: hasError ? '#e06060' : 'rgba(0, 0, 0, 0.15)' }}
            rows={field.rows || 4}
            placeholder={field.placeholder}
            value={values[field.key] || ''}
            disabled={isDisabled}
            onChange={(e) => {
              setValues((prev) => ({ ...prev, [field.key]: e.target.value }));
              if (errors[field.key]) {
                setErrors((prev) => { const n = { ...prev }; delete n[field.key]; return n; });
              }
            }}
          />
        );

      case 'email-list': {
        return renderTagInput(field, currentInputStyle);
      }

      case 'people-picker':
        if (!peoplePickerContext) {
          return renderTagInput(field, currentInputStyle);
        }
        return (
          <PeoplePicker
            context={peoplePickerContext}
            titleText=""
            placeholder={field.placeholder}
            personSelectionLimit={10}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={emailTags[field.key] || []}
            disabled={isDisabled}
            required={field.required}
            onChange={(items) => {
              setEmailTags((prev) => ({
                ...prev,
                [field.key]: normalizePeoplePickerRecipients(items || [])
              }));
              if (errors[field.key]) {
                setErrors((prev) => {
                  const next = { ...prev };
                  delete next[field.key];
                  return next;
                });
              }
            }}
          />
        );

      case 'datetime': {
        // Split stored ISO value into date + time parts
        const storedVal = values[field.key] || '';
        let datePart = '';
        let timePart = '09:00';
        if (storedVal) {
          const parts = storedVal.split('T');
          datePart = parts[0] || '';
          timePart = (parts[1] || '09:00').substring(0, 5);
        }
        return (
          <div style={{ display: 'flex', gap: 6 }}>
            <input
              style={{ ...currentInputStyle, flex: 2 }}
              type="date"
              value={datePart}
              disabled={isDisabled}
              onChange={(e) => {
                const newDate = e.target.value;
                setValues((prev) => ({
                  ...prev,
                  [field.key]: newDate ? newDate + 'T' + timePart + ':00' : ''
                }));
              }}
            />
            <select
              style={{ ...currentInputStyle, flex: 1 }}
              value={timePart}
              disabled={isDisabled}
              onChange={(e) => {
                const newTime = e.target.value;
                setValues((prev) => ({
                  ...prev,
                  [field.key]: datePart ? datePart + 'T' + newTime + ':00' : ''
                }));
              }}
            >
              {TIME_SLOTS.map((slot) => (
                <option key={slot} value={slot}>{slot}</option>
              ))}
            </select>
          </div>
        );
      }

      case 'date':
        return (
          <input
            style={currentInputStyle}
            type="date"
            value={values[field.key] || ''}
            disabled={isDisabled}
            onChange={(e) => {
              setValues((prev) => ({ ...prev, [field.key]: e.target.value }));
            }}
          />
        );

      case 'toggle': {
        const isOn = values[field.key] === 'true';
        return (
          <div style={toggleContainerStyle}>
            <div
              style={isOn ? toggleTrackOnStyle : toggleTrackStyle}
              onClick={() => {
                if (!isDisabled) {
                  setValues((prev) => ({ ...prev, [field.key]: isOn ? 'false' : 'true' }));
                }
              }}
            >
              <div style={isOn ? toggleThumbOnStyle : toggleThumbStyle} />
            </div>
            <span style={{ fontSize: 12, color: '#605e5c' }}>{isOn ? 'Yes' : 'No'}</span>
          </div>
        );
      }

      case 'dropdown':
        return (
          <select
            style={currentInputStyle}
            value={values[field.key] || ''}
            disabled={isDisabled}
            onChange={(e) => {
              setValues((prev) => ({ ...prev, [field.key]: e.target.value }));
            }}
          >
            <option value="">{field.placeholder || strings.SelectPlaceholder}</option>
            {(field.options || []).map((opt) => (
              <option key={opt.key} value={opt.key}>{opt.text}</option>
            ))}
          </select>
        );

      case 'team-picker':
        if (!pickerAppContext) {
          return (
            <input
              style={currentInputStyle}
              type="text"
              value={values.teamName || ''}
              disabled
              placeholder={strings.TeamPickerUnavailable}
            />
          );
        }
        return (
          <TeamPicker
            appcontext={pickerAppContext}
            selectedTeams={selectedTeams}
            itemLimit={1}
            onSelectedTeams={(tagsList) => {
              const selected = tagsList[0];
              setValues((prev) => ({
                ...prev,
                teamId: selected ? String(selected.key) : '',
                teamName: selected?.name || '',
                channelId: '',
                channelName: ''
              }));
              setErrors((prev) => {
                const next = { ...prev };
                delete next.teamId;
                delete next.channelId;
                return next;
              });
            }}
          />
        );

      case 'channel-picker':
        if (!pickerAppContext) {
          return (
            <input
              style={currentInputStyle}
              type="text"
              value={values.channelName || ''}
              disabled
              placeholder={strings.ChannelPickerUnavailable}
            />
          );
        }
        if (!(values.teamId || '').trim()) {
          return (
            <div
              style={{
                ...currentInputStyle,
                color: '#605e5c',
                background: 'rgba(0, 0, 0, 0.03)'
              }}
            >
              Select a team first.
            </div>
          );
        }
        return (
          <TeamChannelPicker
            appcontext={pickerAppContext}
            teamId={values.teamId}
            selectedChannels={selectedChannels}
            itemLimit={1}
            onSelectedChannels={(tagsList) => {
              const selected = tagsList[0];
              setValues((prev) => ({
                ...prev,
                channelId: selected ? String(selected.key) : '',
                channelName: selected?.name || ''
              }));
              setErrors((prev) => {
                const next = { ...prev };
                delete next.channelId;
                return next;
              });
            }}
          />
        );

      default:
        return null;
    }
  };

  // ─── Group fields ───────────────────────────────────────────

  const renderFields = (): React.ReactElement[] => {
    const elements: React.ReactElement[] = [];
    let currentGroup: string | undefined;
    let halfBuffer: IFormFieldDefinition[] = [];

    const renderFieldWrapper = (field: IFormFieldDefinition): React.ReactElement => (
      <div key={field.key} style={{ marginBottom: 10 }}>
        <div style={fieldLabelStyle}>
          {field.label}
          {field.required && <span style={requiredStarStyle}>*</span>}
        </div>
        {renderField(field)}
        {errors[field.key] && <div style={errorTextStyle}>{errors[field.key]}</div>}
      </div>
    );

    const flushHalves = (): void => {
      if (halfBuffer.length === 0) return;
      if (halfBuffer.length === 2) {
        elements.push(
          <div key={`half-${halfBuffer[0].key}-${halfBuffer[1].key}`} style={halfRowStyle}>
            <div style={halfFieldStyle}>{renderFieldWrapper(halfBuffer[0])}</div>
            <div style={halfFieldStyle}>{renderFieldWrapper(halfBuffer[1])}</div>
          </div>
        );
      } else {
        halfBuffer.forEach((f) => {
          elements.push(renderFieldWrapper(f));
        });
      }
      halfBuffer = [];
    };

    const visibleFields = fields.filter((f) => isFieldVisible(f, values));

    visibleFields.forEach((field) => {
      // Group header
      if (field.group !== currentGroup) {
        flushHalves();
        currentGroup = field.group;
        if (currentGroup) {
          elements.push(
            <div key={`group-${currentGroup}`} style={groupLabelStyle}>{currentGroup}</div>
          );
        }
      }

      // Handle half-width fields
      if (field.width === 'half') {
        halfBuffer.push(field);
        if (halfBuffer.length === 2) {
          flushHalves();
        }
      } else {
        flushHalves();
        elements.push(renderFieldWrapper(field));
      }
    });

    flushHalves();
    return elements;
  };

  const renderMessageWithLinks = React.useCallback((message: string): React.ReactNode => {
    const parts = message.split(URL_PATTERN);
    return parts.map((part, index) => {
      if (part.match(URL_PATTERN)) {
        return (
          <a
            key={`msg-link-${index}`}
            href={part}
            target="_blank"
            rel="noreferrer"
            style={messageLinkStyle}
          >
            {part}
          </a>
        );
      }

      return <React.Fragment key={`msg-text-${index}`}>{part}</React.Fragment>;
    });
  }, []);

  // ─── Status overlays ───────────────────────────────────────

  if (formStatus === 'success') {
    return (
      <div style={successStyle}>
        <Icon iconName="CheckMark" styles={{ root: { fontSize: 24, color: '#107c10', marginBottom: 8 } }} />
        <div style={{ fontSize: 13 }}>{renderMessageWithLinks(successMessage || 'Operation completed successfully.')}</div>
      </div>
    );
  }

  if (formStatus === 'error' && errorMessage) {
    return (
      <div>
        <div style={errorMsgStyle}>
          <Icon iconName="ErrorBadge" styles={{ root: { fontSize: 24, color: '#e06060', marginBottom: 8 } }} />
          <div style={{ fontSize: 13 }}>{errorMessage}</div>
        </div>
        <div style={buttonRowStyle}>
          <button style={submitButtonStyle} onClick={() => {
            setFormStatus('editing');
            // Clear the block error state too
            if (onUpdateBlock && blockId) {
              onUpdateBlock(blockId, { data: { ...data, status: 'editing', errorMessage: undefined } });
            }
          }}>
            Edit &amp; Retry
          </button>
          <button style={cancelButtonStyle} onClick={handleCancel}>Cancel</button>
        </div>
      </div>
    );
  }

  // ─── Main form ─────────────────────────────────────────────

  return (
    <div>
      {description && <div style={descStyle}>{renderMessageWithLinks(description)}</div>}
      {attachmentUris.length > 0 && (
        <div style={attachmentSectionStyle}>
          <div style={groupLabelStyle}>Attachments</div>
          <div style={attachmentListStyle}>
            {attachmentUris.map((uri, index) => (
              <div key={`attachment-${index}`} style={attachmentRowStyle}>
                <button
                  type="button"
                  style={attachmentOpenButtonStyle}
                  onClick={(event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    openAttachmentInNewTab(uri);
                  }}
                >
                  {getAttachmentDisplayName(uri)}
                </button>
                {!isDisabled && (
                  <button
                    type="button"
                    aria-label={`Remove attachment ${getAttachmentDisplayName(uri)}`}
                    style={attachmentRemoveButtonStyle}
                    onClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      removeAttachment(uri);
                    }}
                  >
                    Remove
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
      {renderFields()}
      {formStatus === 'submitting' ? (
        <div style={{ textAlign: 'center', marginTop: 16 }}>
          <Spinner size={SpinnerSize.small} label={strings.SubmittingLabel} />
        </div>
      ) : (
        <div style={buttonRowStyle}>
          <button style={submitButtonStyle} onClick={handleSubmit} disabled={isDisabled}>
            Submit
          </button>
          <button style={cancelButtonStyle} onClick={handleCancel} disabled={isDisabled}>
            Cancel
          </button>
        </div>
      )}
    </div>
  );
};

// ─── Helpers ────────────────────────────────────────────────────

function buildFieldSummary(
  formData: IFormData,
  values: Record<string, string>,
  emailTags: Record<string, string[]>
): string {
  const parts: string[] = [];
  formData.fields.forEach((f) => {
    if (!isFieldVisible(f, values)) return;
    if (f.type === 'email-list' || f.type === 'people-picker') {
      const tags = emailTags[f.key] || [];
      if (tags.length > 0) {
        parts.push(`${f.label}: ${tags.join(', ')}`);
      }
    } else {
      const val = values[f.key];
      if (val) {
        parts.push(`${f.label}: ${val.length > 50 ? val.substring(0, 50) + '...' : val}`);
      }
    }
  });
  return parts.join(', ');
}
