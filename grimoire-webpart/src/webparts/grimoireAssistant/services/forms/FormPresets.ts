/**
 * FormPresets — Preset field layouts and submission targets for form blocks.
 * Maps FormPresetId → default fields + MCP tool routing.
 */

import type { FormPresetId, IFormFieldDefinition, IFormSubmissionTarget } from '../../models/IBlock';

export interface IFormPreset {
  fields: IFormFieldDefinition[];
  submissionTarget: IFormSubmissionTarget;
}

// ─── Preset Definitions ─────────────────────────────────────────

const EMAIL_COMPOSE: IFormPreset = {
  fields: [
    { key: 'to', label: 'To', type: 'email-list', required: true, placeholder: 'name@example.com', group: 'Recipients' },
    { key: 'cc', label: 'Cc', type: 'email-list', required: false, placeholder: 'Optional', group: 'Recipients' },
    { key: 'bcc', label: 'Bcc', type: 'email-list', required: false, placeholder: 'Optional', group: 'Recipients' },
    { key: 'subject', label: 'Subject', type: 'text', required: true, placeholder: 'Email subject', group: 'Message' },
    { key: 'body', label: 'Body', type: 'textarea', required: true, placeholder: 'Write your message...', rows: 8, group: 'Message' }
  ],
  submissionTarget: {
    toolName: 'SendEmailWithAttachments',
    serverId: 'mcp_MailTools',
    staticArgs: {}
  }
};

const EMAIL_REPLY: IFormPreset = {
  fields: [
    { key: 'comment', label: 'Reply', type: 'textarea', required: true, placeholder: 'Write your reply...', rows: 6 }
  ],
  submissionTarget: {
    toolName: 'ReplyToMessage',
    serverId: 'mcp_MailTools',
    staticArgs: {}
  }
};

const EMAIL_FORWARD: IFormPreset = {
  fields: [
    { key: 'additionalTo', label: 'To', type: 'email-list', required: true, placeholder: 'name@example.com', group: 'Recipients' },
    { key: 'introComment', label: 'Message', type: 'textarea', required: false, placeholder: 'Add a message...', rows: 4, group: 'Message' }
  ],
  submissionTarget: {
    toolName: 'ForwardMessage',
    serverId: 'mcp_MailTools',
    staticArgs: {}
  }
};

const EMAIL_REPLY_ALL_THREAD: IFormPreset = {
  fields: [
    { key: 'introComment', label: 'Message', type: 'textarea', required: true, placeholder: 'Add your recap...', rows: 8 }
  ],
  submissionTarget: {
    toolName: 'ReplyAllWithFullThread',
    serverId: 'mcp_MailTools',
    staticArgs: {
      includeOriginalNonInlineAttachments: false
    }
  }
};

const EVENT_CREATE: IFormPreset = {
  fields: [
    { key: 'subject', label: 'Title', type: 'text', required: true, placeholder: 'Event title' },
    { key: 'startDateTime', label: 'Start', type: 'datetime', required: true, group: 'When', width: 'half' },
    { key: 'endDateTime', label: 'End', type: 'datetime', required: true, group: 'When', width: 'half' },
    { key: 'location', label: 'Location', type: 'text', required: false, placeholder: 'Room or address', group: 'Details' },
    { key: 'attendeeEmails', label: 'Attendees', type: 'email-list', required: false, placeholder: 'name@example.com', group: 'Details' },
    { key: 'bodyContent', label: 'Description', type: 'textarea', required: false, placeholder: 'Event description...', rows: 4, group: 'Details' },
    { key: 'isOnlineMeeting', label: 'Online meeting', type: 'toggle', required: false, group: 'Options' }
  ],
  submissionTarget: {
    toolName: 'CreateEvent',
    serverId: 'mcp_CalendarTools',
    staticArgs: {}
  }
};

const EVENT_UPDATE: IFormPreset = {
  fields: [
    { key: 'subject', label: 'Title', type: 'text', required: false, placeholder: 'Event title' },
    { key: 'startDateTime', label: 'Start', type: 'datetime', required: false, group: 'When', width: 'half' },
    { key: 'endDateTime', label: 'End', type: 'datetime', required: false, group: 'When', width: 'half' },
    { key: 'location', label: 'Location', type: 'text', required: false, placeholder: 'Room or address', group: 'Details' },
    { key: 'body', label: 'Description', type: 'textarea', required: false, placeholder: 'Event description...', rows: 4, group: 'Details' }
  ],
  submissionTarget: {
    toolName: 'UpdateEvent',
    serverId: 'mcp_CalendarTools',
    staticArgs: {}
  }
};

const TEAMS_MESSAGE: IFormPreset = {
  fields: [
    { key: 'content', label: 'Message', type: 'textarea', required: true, placeholder: 'Type your message...', rows: 4 }
  ],
  submissionTarget: {
    toolName: 'PostMessage',
    serverId: 'mcp_TeamsServer',
    staticArgs: {}
  }
};

const TEAMS_CHANNEL_MESSAGE: IFormPreset = {
  fields: [
    { key: 'content', label: 'Message', type: 'textarea', required: true, placeholder: 'Type your message...', rows: 4 }
  ],
  submissionTarget: {
    toolName: 'PostChannelMessage',
    serverId: 'mcp_TeamsServer',
    staticArgs: {}
  }
};

const SHARE_TEAMS_CHAT: IFormPreset = {
  fields: [
    { key: 'recipients', label: 'People', type: 'people-picker', required: true, placeholder: 'Search people in your organization', group: 'Destination' },
    { key: 'topic', label: 'Chat topic', type: 'text', required: false, placeholder: 'Optional group chat topic', group: 'Destination' },
    { key: 'content', label: 'Message', type: 'textarea', required: true, placeholder: 'Review the share message...', rows: 8, group: 'Message' }
  ],
  submissionTarget: {
    toolName: 'share_teams_chat',
    serverId: 'internal-share',
    staticArgs: {}
  }
};

const SHARE_TEAMS_CHANNEL: IFormPreset = {
  fields: [
    { key: 'teamId', label: 'Team', type: 'team-picker', required: true, group: 'Destination' },
    { key: 'teamName', label: 'Team name', type: 'hidden', required: false, group: 'Destination' },
    { key: 'channelId', label: 'Channel', type: 'channel-picker', required: true, group: 'Destination' },
    { key: 'channelName', label: 'Channel name', type: 'hidden', required: false, group: 'Destination' },
    { key: 'content', label: 'Message', type: 'textarea', required: true, placeholder: 'Review the share message...', rows: 8, group: 'Message' }
  ],
  submissionTarget: {
    toolName: 'share_teams_channel',
    serverId: 'internal-share',
    staticArgs: {}
  }
};

const FILE_CREATE: IFormPreset = {
  fields: [
    { key: 'filename', label: 'File name', type: 'text', required: true, placeholder: 'document.txt' },
    { key: 'contentText', label: 'Content', type: 'textarea', required: true, placeholder: 'File content...', rows: 8 }
  ],
  submissionTarget: {
    toolName: 'createSmallTextFile',
    serverId: 'mcp_ODSPRemoteServer',
    staticArgs: {}
  }
};

const WORD_DOCUMENT_CREATE: IFormPreset = {
  fields: [
    { key: 'filename', label: 'Document name', type: 'text', required: true, placeholder: 'Document.docx' },
    { key: 'contentText', label: 'Content', type: 'textarea', required: true, placeholder: 'Document content...', rows: 10 }
  ],
  submissionTarget: {
    toolName: 'createSmallBinaryFile',
    serverId: 'mcp_ODSPRemoteServer',
    staticArgs: {}
  }
};

const FOLDER_CREATE: IFormPreset = {
  fields: [
    { key: 'folderName', label: 'Folder name', type: 'text', required: true, placeholder: 'New folder' }
  ],
  submissionTarget: {
    toolName: 'createFolder',
    serverId: 'mcp_ODSPRemoteServer',
    staticArgs: {}
  }
};

const ITEM_RENAME: IFormPreset = {
  fields: [
    { key: 'newFileOrFolderName', label: 'New name', type: 'text', required: true, placeholder: 'Updated name' }
  ],
  submissionTarget: {
    toolName: 'renameFileOrFolder',
    serverId: 'mcp_ODSPRemoteServer',
    staticArgs: {}
  }
};

const CHANNEL_CREATE: IFormPreset = {
  fields: [
    { key: 'displayName', label: 'Channel name', type: 'text', required: true, placeholder: 'General' },
    { key: 'description', label: 'Description', type: 'textarea', required: false, placeholder: 'Channel description...', rows: 3 }
  ],
  submissionTarget: {
    toolName: 'CreateChannel',
    serverId: 'mcp_TeamsServer',
    staticArgs: {}
  }
};

const CHAT_CREATE: IFormPreset = {
  fields: [
    { key: 'members_upns', label: 'Members', type: 'email-list', required: true, placeholder: 'name@example.com' },
    { key: 'topic', label: 'Topic', type: 'text', required: false, placeholder: 'Chat topic (optional)' }
  ],
  submissionTarget: {
    toolName: 'CreateChat',
    serverId: 'mcp_TeamsServer',
    staticArgs: { chatType: 'group' }
  }
};

const LIST_ITEM_CREATE: IFormPreset = {
  fields: [],
  submissionTarget: {
    toolName: 'createListItem',
    serverId: 'mcp_SharePointListsTools',
    staticArgs: {}
  }
};

const LIST_ITEM_UPDATE: IFormPreset = {
  fields: [],
  submissionTarget: {
    toolName: 'updateListItem',
    serverId: 'mcp_SharePointListsTools',
    staticArgs: {}
  }
};

const GENERIC: IFormPreset = {
  fields: [],
  submissionTarget: {
    toolName: '',
    serverId: '',
    staticArgs: {}
  }
};

// ─── Preset Registry ─────────────────────────────────────────────

const PRESET_MAP: Record<FormPresetId, IFormPreset> = {
  'email-compose': EMAIL_COMPOSE,
  'email-reply': EMAIL_REPLY,
  'email-forward': EMAIL_FORWARD,
  'email-reply-all-thread': EMAIL_REPLY_ALL_THREAD,
  'event-create': EVENT_CREATE,
  'event-update': EVENT_UPDATE,
  'teams-message': TEAMS_MESSAGE,
  'teams-channel-message': TEAMS_CHANNEL_MESSAGE,
  'share-teams-chat': SHARE_TEAMS_CHAT,
  'share-teams-channel': SHARE_TEAMS_CHANNEL,
  'file-create': FILE_CREATE,
  'word-document-create': WORD_DOCUMENT_CREATE,
  'folder-create': FOLDER_CREATE,
  'item-rename': ITEM_RENAME,
  'list-item-create': LIST_ITEM_CREATE,
  'list-item-update': LIST_ITEM_UPDATE,
  'channel-create': CHANNEL_CREATE,
  'chat-create': CHAT_CREATE,
  'generic': GENERIC
};

/**
 * Get a form preset by ID.
 * Returns a deep copy so callers can safely modify fields.
 */
export function getFormPreset(presetId: FormPresetId): IFormPreset {
  const preset = PRESET_MAP[presetId] || GENERIC;
  return {
    fields: preset.fields.map((f) => ({ ...f })),
    submissionTarget: {
      ...preset.submissionTarget,
      staticArgs: { ...preset.submissionTarget.staticArgs },
      ...(preset.submissionTarget.targetContext
        ? { targetContext: { ...preset.submissionTarget.targetContext } }
        : {})
    }
  };
}
