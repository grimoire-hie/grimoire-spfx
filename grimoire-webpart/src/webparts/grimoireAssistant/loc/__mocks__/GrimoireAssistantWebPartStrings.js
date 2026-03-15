// Jest CJS mock for the SPFx AMD loc module.
// Mirrors en-us.js but in CommonJS format so tests can resolve the module.
"use strict";

module.exports = {
  PropertyPaneDescription: "Grimoire Assistant Configuration",
  BasicGroupName: "General",
  DescriptionFieldLabel: "Description",
  BackendGroupName: "Grimoire Backend (Proxy)",
  ProxyUrlFieldLabel: "Proxy URL",
  ProxyUrlFieldDescription: "",
  ProxyApiKeyFieldLabel: "Proxy API Key",
  ProxyApiKeyFieldDescription: "",
  BackendApiResourceFieldLabel: "Backend API Resource",
  BackendApiResourceFieldDescription: "",
  ProxyBackendFieldLabel: "Backend Model",
  ProxyBackendReasoningOptionLabel: "Reasoning (recommended)",
  ProxyBackendFastOptionLabel: "Fast (lightweight)",
  M365McpGroupName: "M365 MCP Servers",
  McpEnvironmentIdFieldLabel: "Environment ID",
  McpEnvironmentIdFieldDescription: "",
  AppLocalEnvironmentSharePoint: "",
  AppLocalEnvironmentTeams: "",
  AppLocalEnvironmentOffice: "",
  AppLocalEnvironmentOutlook: "",
  AppSharePointEnvironment: "",
  AppTeamsTabEnvironment: "",
  AppOfficeEnvironment: "",
  AppOutlookEnvironment: "",
  UnknownEnvironment: "",

  // TextInput
  TypeAMessage: "Type a message...",
  SendMessage: "Send message",

  // ActionPanel
  FocusButton: "Focus",
  SummarizeButton: "Summarize",
  ChatButton: "Chat",
  RecapButton: "Recap",
  RecapLoadingButton: "Recap\u2026",
  ShareButton: "Share",
  EmptyPanelHint: "Results and details will appear here",

  // AvatarPanel status
  NoApiStatus: "No API",
  ApiPendingStatus: "API...",
  ApiStatus: "API",
  VoicePendingStatus: "Voice...",
  VoiceStatus: "Voice",
  MicPendingStatus: "Mic...",
  MicStatus: "Mic",
  LogsButton: "Logs",
  AvatarSettingsTitle: "Avatar settings",
  PreparingAvatar: "Preparing avatar...",
  PreparingAvatarDetail: "Loading the selected SVG and motion bindings in the background.",
  ConnectVoice: "Connect voice",
  TurnAudioOff: "Turn audio off",
  NoBackendConfigured: "No backend configured",

  // AvatarPanel tooltips
  NoBackendTooltip: "No backend configured",
  CheckingApiTooltip: "Checking API...",
  BackendConnectedTooltip: "Backend connected",
  BackendUnavailableTooltip: "Backend unavailable",
  ConnectingVoiceTooltip: "Connecting voice...",
  VoiceDisconnectedTooltip: "Voice disconnected",
  ClickToDisconnectTooltip: "Click to disconnect",
  MicAllowedTooltip: "Mic allowed",
  MicRequestingTooltip: "Requesting mic...",
  MicBlockedTooltip: "Mic blocked",

  // StatusBar
  StatusConnected: "Connected",
  StatusListening: "Listening",
  StatusConnecting: "Connecting",
  StatusError: "Error",
  StatusIdle: "Idle",
  StatusSpeaking: "Speaking",
  StatusOffline: "Offline",
  BackendLabel: "Backend",
  VoiceLabel: "Voice",
  MicLabel: "Mic",
  McpLabel: "MCP",
  McpServersTooltip: "{0} MCP server(s)",
  NoMcpServersTooltip: "No MCP servers",
  MicAllowed: "Microphone allowed",
  MicBlocked: "Microphone blocked",

  // Header
  GrimoireTitle: "Grimoire",
  CloseLogsLabel: "Close logs",
  OpenLogsLabel: "Open logs",

  // AppLayout
  LoadingAvatarSurface: "Loading avatar surface...",
  LoadingAssistantTools: "Loading assistant tools...",
  LoadingLogs: "Loading logs...",
  ResizeLogsPaneLabel: "Resize logs pane",
  ResizeLogsPaneTitle: "Drag to resize logs pane",
  ResizeAvatarActionLabel: "Resize avatar and action panes",
  ResizeAvatarActionTitle: "Drag to resize avatar and action panes",

  // ErrorBoundary
  RenderError: "Grimoire render error",
  RetryButton: "Retry",

  // ConfirmationDialogBlock
  ConfirmedLabel: "Confirmed",
  CancelledLabel: "Cancelled",

  // FormBlock
  SelectPlaceholder: "Select...",
  SubmittingLabel: "Submitting...",
  TeamPickerUnavailable: "Team picker unavailable",
  ChannelPickerUnavailable: "Channel picker unavailable",

  // AvatarSettingsPanel
  RecapOptionAuto: "Auto \u2014 only when results need a recap",
  RecapOptionAlways: "Always \u2014 append a recap after every search",
  RecapOptionOff: "Off \u2014 show results only",
  CloseSettingsLabel: "Close settings",
  AvatarToggleLabel: "Avatar",
  ToggleEnabled: "Enabled",
  ToggleDisabled: "Disabled",
  PublicWebSearchLabel: "Public Web Search (Preview)",
  CopilotWebGroundingLabel: "Copilot Web Grounding (M365)",
  VoiceReconnectHint: "Changing the voice reconnects the live session so the new voice applies immediately.",
  VoiceChangeNextConnect: "Voice changes apply the next time you connect voice.",

  // Tool Activity Labels
  ActivitySearchingSharePoint: "Searching SharePoint...",
  ActivitySearchingPeople: "Searching people...",
  ActivitySearchingSites: "Searching sites...",
  ActivitySearchingEmails: "Searching emails...",
  ActivityResearchingWeb: "Researching the web...",
  ActivityBrowsingLibrary: "Browsing library...",
  ActivityLoadingFileDetails: "Loading file details...",
  ActivityLoadingSiteInfo: "Loading site info...",
  ActivityLoadingListItems: "Loading list items...",
  ActivityReadingFile: "Reading file...",
  ActivityReadingEmail: "Reading email...",
  ActivityReadingMessages: "Reading messages...",
  ActivityConnectingToServer: "Connecting to server...",
  ActivityRunningMcpTool: "Running MCP tool...",
  ActivityListingTools: "Listing tools...",
  ActivityRunningM365Tool: "Running M365 tool...",
  ActivityLoadingServerCatalog: "Loading server catalog...",
  ActivityPreparingForm: "Preparing form...",
  ActivityLoadingProfile: "Loading profile...",
  ActivityLoadingRecentDocs: "Loading recent documents...",
  ActivityLoadingTrendingDocs: "Loading trending documents...",
  ActivitySavingNote: "Saving note...",
  ActivityRecallingNotes: "Recalling notes...",
  ActivityDeletingNote: "Deleting note..."
};
