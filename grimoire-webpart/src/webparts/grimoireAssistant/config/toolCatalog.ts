import * as strings from 'GrimoireAssistantWebPartStrings';
import type { BlockType } from '../models/IBlock';

export type ToolCatalogCategory =
  | 'search'
  | 'browse'
  | 'content-reading'
  | 'mcp'
  | 'ui-display'
  | 'status-expression'
  | 'selection-data'
  | 'm365-catalog'
  | 'form-composition'
  | 'personal-context';

export type ToolRuntimePartition = 'search' | 'content' | 'mcp' | 'ui-personal';

export interface IToolParameterProperty {
  type: string;
  description: string;
  enum?: readonly string[];
  items?: { type: string };
}

export interface IToolDefinition {
  type: 'function';
  name: string;
  description: string;
  parameters: {
    type: 'object';
    properties: Record<string, IToolParameterProperty>;
    required?: readonly string[];
  };
}

export interface IToolCatalogEntry extends IToolDefinition {
  category: ToolCatalogCategory;
  runtimePartition: ToolRuntimePartition;
  asyncBehavior: 'sync' | 'async';
  activityLabel?: string;
  defaultBlockType?: BlockType;
  systemPromptDescription?: string;
}

const SEARCH_SHAREPOINT_QUERY_DESCRIPTION =
  'Preserve the user request in natural language. Keep the original phrasing, misspellings, and language instead of compressing it to keywords.';

const BASE_TOOL_CATALOG = [
  {
    type: 'function',
    name: 'search_sharepoint',
    category: 'search',
    runtimePartition: 'search',
    asyncBehavior: 'async',
    activityLabel: strings.ActivitySearchingSharePoint,
    defaultBlockType: 'search-results',
    systemPromptDescription: 'Search for documents, files, and content across SharePoint and OneDrive. Use this for internal M365 content requests and as the default for generic enterprise search phrasing when the user did not name a workload or external source.',
    description: 'Search for documents, files, and content across SharePoint and OneDrive. Use this for INTERNAL M365 content discovery. Do not use this for external/public web URLs. Returns search results displayed in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: SEARCH_SHAREPOINT_QUERY_DESCRIPTION },
        max_results: { type: 'string', description: 'Maximum number of results (default: "10")' },
        render_hints: { type: 'string', description: 'Optional JSON render hints: {"highlight": [1,3], "annotate": {"1": "most relevant"}, "collapse": false}' }
      },
      required: ['query']
    }
  },
  {
    type: 'function',
    name: 'search_people',
    category: 'search',
    runtimePartition: 'search',
    asyncBehavior: 'async',
    activityLabel: strings.ActivitySearchingPeople,
    defaultBlockType: 'user-card',
    systemPromptDescription: 'Find people by name, title, or department in the organization.',
    description: 'Find people in the organization by name, title, department, or skills. Returns user cards in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Person name, title, department, or skills to search for' },
        max_results: { type: 'string', description: 'Maximum number of results (default: "5")' }
      },
      required: ['query']
    }
  },
  {
    type: 'function',
    name: 'search_sites',
    category: 'search',
    runtimePartition: 'search',
    asyncBehavior: 'async',
    activityLabel: strings.ActivitySearchingSites,
    defaultBlockType: 'site-info',
    systemPromptDescription: 'Discover SharePoint sites by name or topic.',
    description: 'Discover SharePoint sites by name or topic. Returns site info cards in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Site name or topic to search for' },
        max_results: { type: 'string', description: 'Maximum number of results (default: "5")' }
      },
      required: ['query']
    }
  },
  {
    type: 'function',
    name: 'search_emails',
    category: 'search',
    runtimePartition: 'search',
    asyncBehavior: 'async',
    activityLabel: strings.ActivitySearchingEmails,
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Search for emails by keyword, sender, subject, or date. Use this for any email-related request.',
    description: 'Search emails in Outlook by keyword, sender, subject, or date. Returns email results displayed in the action panel. Use this for any email-related query. For requests like "show my recent/latest emails", use query "*" or a broad term like "recent".',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search keywords (e.g., "budget report", "from:john", "meeting agenda"). For recent emails, use a broad term.' },
        folder: { type: 'string', description: 'Optional mail folder to search in (e.g., "Inbox", "Sent Items", "Drafts")' },
        max_results: { type: 'string', description: 'Maximum number of results (default: "10")' }
      },
      required: ['query']
    }
  },
  {
    type: 'function',
    name: 'research_public_web',
    category: 'search',
    runtimePartition: 'search',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityResearchingWeb,
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Research public websites, public web topics, or direct URLs through Azure OpenAI web search preview. Use this for explicit public-web, internet, website, or URL requests. Pass the user request in `query` and include `target_url` when a specific page is named.',
    description: 'Research the public web or summarize a public URL using Azure OpenAI web_search_preview. Use this for explicit public-web, website, internet, GitHub, Wikipedia, or direct-URL requests. Do not use M365 Copilot tools for this. Results appear in the action panel with cited sources.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Preserve the user request in natural language. Example: "Summarize this page and tell me what the repository is about".' },
        target_url: { type: 'string', description: 'Optional specific public URL to inspect or summarize.' }
      },
      required: ['query']
    }
  },
  {
    type: 'function',
    name: 'browse_document_library',
    category: 'browse',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityBrowsingLibrary,
    defaultBlockType: 'document-library',
    systemPromptDescription: 'List files and folders in a SharePoint document library.',
    description: 'List files and folders in a SharePoint document library. Displays a browsable file list in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        site_url: { type: 'string', description: 'SharePoint site URL' },
        library_name: { type: 'string', description: 'Document library name (default: "Documents")' },
        folder_path: { type: 'string', description: 'Optional subfolder path' }
      },
      required: ['site_url']
    }
  },
  {
    type: 'function',
    name: 'show_file_details',
    category: 'browse',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingFileDetails,
    defaultBlockType: 'file-preview',
    systemPromptDescription: 'Show metadata and preview for a specific file.',
    description: 'Show metadata and preview for a specific file. Displays a file preview card in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        file_url: { type: 'string', description: 'Full URL of the file' },
        file_name: { type: 'string', description: 'Display name of the file' },
        render_hints: { type: 'string', description: 'Optional JSON render hints: {"highlight": [1], "annotate": {"1": "requested"}}' }
      },
      required: ['file_url']
    }
  },
  {
    type: 'function',
    name: 'show_site_info',
    category: 'browse',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingSiteInfo,
    defaultBlockType: 'site-info',
    systemPromptDescription: 'Display detailed information about a SharePoint site.',
    description: 'Display detailed information about a SharePoint site including libraries, lists, and storage.',
    parameters: {
      type: 'object',
      properties: {
        site_url: { type: 'string', description: 'SharePoint site URL' }
      },
      required: ['site_url']
    }
  },
  {
    type: 'function',
    name: 'show_list_items',
    category: 'browse',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingListItems,
    defaultBlockType: 'list-items',
    systemPromptDescription: 'Show items from a SharePoint list.',
    description: 'Show items from a SharePoint list in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        site_url: { type: 'string', description: 'SharePoint site URL' },
        list_name: { type: 'string', description: 'Name of the SharePoint list' },
        max_items: { type: 'string', description: 'Maximum items to show (default: "20")' },
        render_hints: { type: 'string', description: 'Optional JSON render hints: {"highlight": [1,2], "annotate": {"1": "important"}, "collapse": false}' }
      },
      required: ['site_url', 'list_name']
    }
  },
  {
    type: 'function',
    name: 'read_file_content',
    category: 'browse',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityReadingFile,
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Process one or more SharePoint/OneDrive files through Copilot Chat API. Always set `mode`: `summarize`, `full`, or `answer`. Use `file_url` for one file or `file_urls` for multiple files.',
    description: 'Process one or more SharePoint/OneDrive files with Copilot Chat API. Use mode "summarize" for summaries, mode "full" to retrieve full content (single file), and mode "answer" to answer a specific question. Provide file_url for one file, or file_urls for multiple files.',
    parameters: {
      type: 'object',
      properties: {
        file_url: { type: 'string', description: 'Full URL of the file (from search results or document library)' },
        file_urls: {
          type: 'array',
          items: { type: 'string' },
          description: 'Optional list of file URLs for multi-document summarize/answer operations.'
        },
        file_name: { type: 'string', description: 'Display name of the file' },
        mode: {
          type: 'string',
          description: 'Processing mode: summarize (default for summarize requests), full (complete content), or answer (answer the provided question).',
          enum: ['summarize', 'full', 'answer']
        },
        question: { type: 'string', description: 'Optional question to answer when mode is "answer".' }
      }
    }
  },
  {
    type: 'function',
    name: 'read_email_content',
    category: 'content-reading',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityReadingEmail,
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Process email content via Copilot Chat API. Always set `mode`: `summarize`, `full`, or `answer`.',
    description: 'Process email content via Copilot Chat API. Use mode "summarize" for summaries, mode "full" for complete retrieval, and mode "answer" for question-focused responses.',
    parameters: {
      type: 'object',
      properties: {
        subject: { type: 'string', description: 'Preferred identifier: email subject line from search results or context' },
        sender: { type: 'string', description: 'Optional secondary identifier: sender name or email address' },
        date_hint: { type: 'string', description: 'Optional secondary identifier: approximate date (e.g., "last week", "2026-03-01")' },
        mode: {
          type: 'string',
          description: 'Processing mode: summarize, full, or answer.',
          enum: ['summarize', 'full', 'answer']
        },
        question: { type: 'string', description: 'Optional specific question about the email content (required for mode "answer").' }
      },
      required: []
    }
  },
  {
    type: 'function',
    name: 'read_teams_messages',
    category: 'content-reading',
    runtimePartition: 'content',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityReadingMessages,
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Process Teams chat/channel content via Copilot Chat API. Always set `mode`: `summarize`, `full`, or `answer`.',
    description: 'Process Teams chat/channel content via Copilot Chat API. Use mode "summarize" for summaries, mode "full" for complete retrieval, and mode "answer" for question-focused responses.',
    parameters: {
      type: 'object',
      properties: {
        chat_or_channel: { type: 'string', description: 'Name or description of the chat/channel (e.g., "General in Engineering team", "my chat with John")' },
        topic: { type: 'string', description: 'Optional topic or keyword to focus on' },
        mode: {
          type: 'string',
          description: 'Processing mode: summarize, full, or answer.',
          enum: ['summarize', 'full', 'answer']
        },
        question: { type: 'string', description: 'Optional specific question about the messages (required for mode "answer").' }
      },
      required: ['chat_or_channel']
    }
  },
  {
    type: 'function',
    name: 'connect_mcp_server',
    category: 'mcp',
    runtimePartition: 'mcp',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityConnectingToServer,
    systemPromptDescription: 'Connect to an MCP server to access its tools.',
    description: 'Connect to an MCP server to access its tools. Call this before using call_mcp_tool.',
    parameters: {
      type: 'object',
      properties: {
        server_url: { type: 'string', description: 'MCP server URL' },
        server_name: { type: 'string', description: 'Display name for the server' }
      },
      required: ['server_url']
    }
  },
  {
    type: 'function',
    name: 'call_mcp_tool',
    category: 'mcp',
    runtimePartition: 'mcp',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityRunningMcpTool,
    systemPromptDescription: 'Execute a tool on a connected MCP server.',
    description: 'Execute a tool on a connected MCP server. The server must be connected first via connect_mcp_server.',
    parameters: {
      type: 'object',
      properties: {
        session_id: { type: 'string', description: 'Session ID from connect_mcp_server' },
        tool_name: { type: 'string', description: 'Name of the MCP tool to call' },
        arguments_json: { type: 'string', description: 'Tool arguments as JSON string' }
      },
      required: ['tool_name']
    }
  },
  {
    type: 'function',
    name: 'list_mcp_tools',
    category: 'mcp',
    runtimePartition: 'mcp',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityListingTools,
    systemPromptDescription: 'List available tools on a connected MCP server.',
    description: 'List available tools on a connected MCP server.',
    parameters: {
      type: 'object',
      properties: {
        session_id: { type: 'string', description: 'Session ID from connect_mcp_server' }
      },
      required: []
    }
  },
  {
    type: 'function',
    name: 'show_info_card',
    category: 'ui-display',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'info-card',
    systemPromptDescription: 'Display an informational card with a heading and body.',
    description: 'Display an informational card with a heading and body text in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        heading: { type: 'string', description: 'Card heading' },
        body: { type: 'string', description: 'Card body text' },
        icon: { type: 'string', description: 'Optional Fluent UI icon name' }
      },
      required: ['heading', 'body']
    }
  },
  {
    type: 'function',
    name: 'show_markdown',
    category: 'ui-display',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Display formatted markdown content.',
    description: 'Display formatted markdown content in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Block title' },
        content: { type: 'string', description: 'Markdown content to display' }
      },
      required: ['title', 'content']
    }
  },
  {
    type: 'function',
    name: 'ask_confirmation',
    category: 'ui-display',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'confirmation-dialog',
    systemPromptDescription: 'Ask the user to confirm an action before proceeding.',
    description: 'Ask the user to confirm an action before proceeding. Shows a confirmation dialog in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        message: { type: 'string', description: 'Confirmation message' },
        confirm_label: { type: 'string', description: 'Label for the confirm button (default: "Confirm")' },
        cancel_label: { type: 'string', description: 'Label for the cancel button (default: "Cancel")' }
      },
      required: ['message']
    }
  },
  {
    type: 'function',
    name: 'clear_action_panel',
    category: 'ui-display',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    systemPromptDescription: 'Clear all displayed blocks from the action panel.',
    description: 'Clear all displayed blocks from the action panel.',
    parameters: {
      type: 'object',
      properties: {},
      required: []
    }
  },
  {
    type: 'function',
    name: 'set_expression',
    category: 'status-expression',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    systemPromptDescription: 'Change your facial expression (happy, thinking, surprised, confused, idle).',
    description: 'Change your facial expression on the avatar. Use this to visually communicate your state to the user.',
    parameters: {
      type: 'object',
      properties: {
        expression: {
          type: 'string',
          description: 'The expression to set',
          enum: ['idle', 'listening', 'thinking', 'speaking', 'surprised', 'happy', 'confused']
        }
      },
      required: ['expression']
    }
  },
  {
    type: 'function',
    name: 'show_progress',
    category: 'status-expression',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'progress-tracker',
    systemPromptDescription: 'Show a progress indicator for long-running operations.',
    description: 'Show a progress indicator for long-running operations in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        label: { type: 'string', description: 'Progress label' },
        progress: { type: 'string', description: 'Progress percentage (0-100)' }
      },
      required: ['label']
    }
  },
  {
    type: 'function',
    name: 'show_error',
    category: 'status-expression',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'error',
    systemPromptDescription: 'Display an error message to the user.',
    description: 'Display an error message to the user in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        message: { type: 'string', description: 'Error message' },
        detail: { type: 'string', description: 'Optional error detail' }
      },
      required: ['message']
    }
  },
  {
    type: 'function',
    name: 'use_m365_capability',
    category: 'm365-catalog',
    runtimePartition: 'mcp',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityRunningM365Tool,
    systemPromptDescription: 'Execute any M365 MCP tool by name. Auto-connects to the right server, acquires token, executes, and displays results. Use this for operations like listing files, sending email, creating events, managing Teams, etc.',
    description: 'Execute an M365 MCP tool by name. Automatically connects to the correct Agent 365 server, acquires a token, executes, and displays results. Use this for M365 operations like listing files, sending email, managing calendar, Teams, SharePoint lists, user profiles, and internal Copilot chat tasks. Do not use this for explicit public-web or URL research.',
    parameters: {
      type: 'object',
      properties: {
        tool_name: { type: 'string', description: 'The MCP tool name to execute (e.g., listFolderContents, sendMail, createEvent, listChats, getMyProfile)' },
        arguments_json: { type: 'string', description: 'Tool arguments as JSON string (e.g., \'{"searchQuery": "budget"}\')' },
        server_hint: { type: 'string', description: 'Optional server ID hint if the tool name is ambiguous (e.g., mcp_ODSPRemoteServer)' }
      },
      required: ['tool_name']
    }
  },
  {
    type: 'function',
    name: 'list_m365_servers',
    category: 'm365-catalog',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingServerCatalog,
    systemPromptDescription: 'List available M365 MCP capabilities. Use with no arguments for a general overview, or pass `focus` for workload-specific capability questions such as "what can you do for SharePoint?" or "what can you do in Teams?".',
    description: 'List available M365 MCP capabilities. Use this for general capability questions or pass `focus` to drill into one workload such as sharepoint, onedrive, outlook, mail, calendar, teams, word, profile, or copilot.',
    parameters: {
      type: 'object',
      properties: {
        focus: {
          type: 'string',
          description: 'Optional workload focus for detailed capability drill-down. Supported aliases: sharepoint, onedrive, outlook, mail, calendar, teams, word, profile, copilot.'
        }
      },
      required: []
    }
  },
  {
    type: 'function',
    name: 'show_selection_list',
    category: 'selection-data',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'selection-list',
    systemPromptDescription: 'Present a list of options for the user to choose from.',
    description: 'Present a list of options for the user to choose from in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        prompt: { type: 'string', description: 'Question or prompt for the user' },
        items_json: { type: 'string', description: 'JSON array of {id, label, description} objects' },
        multi_select: { type: 'string', description: 'Allow multiple selections ("true" or "false", default: "false")' }
      },
      required: ['prompt', 'items_json']
    }
  },
  {
    type: 'function',
    name: 'show_chart',
    category: 'selection-data',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'chart',
    systemPromptDescription: 'Display a simple chart (bar, pie, or line).',
    description: 'Display a simple chart (bar, pie, or line) in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Chart title' },
        chart_type: { type: 'string', description: 'Chart type', enum: ['bar', 'pie', 'line'] },
        labels_json: { type: 'string', description: 'JSON array of label strings' },
        values_json: { type: 'string', description: 'JSON array of numeric values' }
      },
      required: ['title', 'chart_type', 'labels_json', 'values_json']
    }
  },
  {
    type: 'function',
    name: 'show_permissions',
    category: 'selection-data',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'permissions-view',
    systemPromptDescription: 'Inspect permissions for a file, folder, or site through the connected SharePoint & OneDrive MCP tools when that capability is available.',
    description: 'Inspect permissions for a file, folder, or site in the action panel through MCP only. If the connected SharePoint & OneDrive MCP tools do not expose permission inspection, this returns an honest unsupported result instead of falling back.',
    parameters: {
      type: 'object',
      properties: {
        target_name: { type: 'string', description: 'Name of the file, folder, or site' },
        target_url: { type: 'string', description: 'URL of the target' }
      },
      required: ['target_name', 'target_url']
    }
  },
  {
    type: 'function',
    name: 'show_activity_feed',
    category: 'selection-data',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    defaultBlockType: 'activity-feed',
    description: 'Show recent file activity (edits, uploads) for a SharePoint site in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        site_url: { type: 'string', description: 'SharePoint site URL' },
        max_items: { type: 'string', description: 'Maximum activities to show (default: "20")' }
      },
      required: ['site_url']
    }
  },
  {
    type: 'function',
    name: 'show_compose_form',
    category: 'form-composition',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'sync',
    activityLabel: strings.ActivityPreparingForm,
    defaultBlockType: 'form',
    systemPromptDescription: 'Display an interactive form for composing content. Use the appropriate preset (email-compose, event-create, teams-message, etc.) and pre-fill all known info from conversation.',
    description: 'Display an interactive form for composing content (email, event, Teams message, file, etc.). Pre-fill fields from conversation context. The user fills in and submits to execute the operation. Use this for any WRITE/CREATE operation instead of collecting info conversationally.',
    parameters: {
      type: 'object',
      properties: {
        preset: {
          type: 'string',
          description: 'Form preset type',
          enum: [
            'email-compose', 'email-reply', 'email-forward', 'email-reply-all-thread',
            'event-create', 'event-update',
            'teams-message', 'teams-channel-message',
            'file-create', 'folder-create',
            'list-item-create', 'list-item-update',
            'channel-create', 'chat-create',
            'generic'
          ]
        },
        title: { type: 'string', description: 'Form title shown in the action panel header' },
        description: { type: 'string', description: 'Optional description shown above the form fields' },
        prefill_json: { type: 'string', description: 'JSON object with pre-filled values keyed by field key (e.g., \'{"to": "john@example.com", "subject": "Budget Report"}\')' },
        static_args_json: { type: 'string', description: 'JSON object with hidden args not shown in form but included in MCP submission (e.g., \'{"messageId": "AAMk..."}\'). For visible-result sharing, include explicit scope here too (for example `shareSelectionIndices`, `shareBlockId`, and `shareScopeMode`) so "first", "selected", and "all" stay exact.' },
        custom_fields_json: { type: 'string', description: 'For generic preset: JSON array of field definitions [{key, label, type, required, placeholder, ...}]' },
        custom_target_json: { type: 'string', description: 'For generic preset: JSON submission target {toolName, serverId, staticArgs, fieldToParamMap}' }
      },
      required: ['preset', 'title']
    }
  },
  {
    type: 'function',
    name: 'get_my_profile',
    category: 'personal-context',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingProfile,
    defaultBlockType: 'user-card',
    systemPromptDescription: `Get the current user's full profile (job title, department, manager, contact info).`,
    description: 'Get the current user\'s full profile with job title, department, office, manager chain, and contact info. Displays a user card in the action panel.',
    parameters: {
      type: 'object',
      properties: {},
      required: []
    }
  },
  {
    type: 'function',
    name: 'get_recent_documents',
    category: 'personal-context',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingRecentDocs,
    defaultBlockType: 'search-results',
    systemPromptDescription: `Documents recently opened or edited by the user. Use for "what was I working on?".`,
    description: 'Get documents recently opened or edited by the current user. Use when the user asks "what was I working on?", "my recent files", or similar. Displays search results in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        max_results: { type: 'string', description: 'Maximum number of results (default: "10")' }
      },
      required: []
    }
  },
  {
    type: 'function',
    name: 'get_trending_documents',
    category: 'personal-context',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityLoadingTrendingDocs,
    defaultBlockType: 'search-results',
    systemPromptDescription: 'Documents trending around the user in the organization.',
    description: 'Get documents trending around the current user in the organization. Shows what colleagues are working on. Displays search results in the action panel.',
    parameters: {
      type: 'object',
      properties: {
        max_results: { type: 'string', description: 'Maximum number of results (default: "10")' }
      },
      required: []
    }
  },
  {
    type: 'function',
    name: 'save_note',
    category: 'personal-context',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivitySavingNote,
    systemPromptDescription: 'Save a note to persistent memory. Use when the user says "remember", "note this", "save for later".',
    description: 'Save a note to persistent memory. Notes survive across sessions. Use when the user says "remember", "note this", "save for later".',
    parameters: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'The note text to save' },
        tags: { type: 'string', description: 'Comma-separated tags for organization (e.g., "budget,meeting,weekly")' }
      },
      required: ['text']
    }
  },
  {
    type: 'function',
    name: 'recall_notes',
    category: 'personal-context',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityRecallingNotes,
    defaultBlockType: 'markdown',
    systemPromptDescription: 'Recall saved notes by keyword or tag.',
    description: 'Recall saved notes by keyword or tag. Use when the user asks about previously saved information, or says "what did I save about...?".',
    parameters: {
      type: 'object',
      properties: {
        keyword: { type: 'string', description: 'Keyword to search for in saved notes and tags' }
      },
      required: []
    }
  },
  {
    type: 'function',
    name: 'delete_note',
    category: 'personal-context',
    runtimePartition: 'ui-personal',
    asyncBehavior: 'async',
    activityLabel: strings.ActivityDeletingNote,
    systemPromptDescription: 'Delete a saved note or all notes. Use when the user says "forget", "delete", "remove that note".',
    description: 'Delete a saved note by its ID, or delete all notes. Use when the user says "forget that", "delete my notes", "remove that note".',
    parameters: {
      type: 'object',
      properties: {
        note_id: { type: 'string', description: 'ID of the note to delete (from recall_notes results). Use "all" to delete all notes.' }
      },
      required: ['note_id']
    }
  }
] as const satisfies readonly IToolCatalogEntry[];

export type ToolCatalogName = typeof BASE_TOOL_CATALOG[number]['name'];
export type ToolCatalogNameByPartition<P extends ToolRuntimePartition> =
  Extract<typeof BASE_TOOL_CATALOG[number], { runtimePartition: P }>['name'];

function cloneToolParameters(parameters: IToolDefinition['parameters']): IToolDefinition['parameters'] {
  const properties = Object.keys(parameters.properties).reduce<Record<string, IToolParameterProperty>>((acc, key) => {
    const property = parameters.properties[key];
    acc[key] = {
      ...property,
      enum: property.enum ? [...property.enum] : undefined,
      items: property.items ? { ...property.items } : undefined
    };
    return acc;
  }, {});

  return {
    type: 'object',
    properties,
    required: parameters.required ? [...parameters.required] : undefined
  };
}

function cloneToolEntry(entry: IToolCatalogEntry): IToolCatalogEntry {
  return {
    ...entry,
    parameters: cloneToolParameters(entry.parameters)
  };
}

function applySearchLanguageRule(entry: IToolCatalogEntry, searchLanguageRule?: string): IToolCatalogEntry {
  if (entry.name !== 'search_sharepoint' || !searchLanguageRule) {
    return entry;
  }

  return {
    ...entry,
    parameters: {
      ...entry.parameters,
      properties: {
        ...entry.parameters.properties,
        query: {
          ...entry.parameters.properties.query,
          description: `${SEARCH_SHAREPOINT_QUERY_DESCRIPTION} ${searchLanguageRule}`
        }
      }
    }
  };
}

export function getToolCatalog(searchLanguageRule?: string): IToolCatalogEntry[] {
  return BASE_TOOL_CATALOG.map((entry) => applySearchLanguageRule(cloneToolEntry(entry), searchLanguageRule));
}

export function getToolCatalogEntry(name: string, searchLanguageRule?: string): IToolCatalogEntry | undefined {
  return getToolCatalog(searchLanguageRule).find((entry) => entry.name === name);
}

export function getToolCatalogCount(): number {
  return BASE_TOOL_CATALOG.length;
}

export function getToolCategoryCount(): number {
  return new Set(BASE_TOOL_CATALOG.map((entry) => entry.category)).size;
}
