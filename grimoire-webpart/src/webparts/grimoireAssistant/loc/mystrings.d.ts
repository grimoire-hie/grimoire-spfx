declare interface IGrimoireAssistantWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  BackendGroupName: string;
  ProxyUrlFieldLabel: string;
  ProxyUrlFieldDescription: string;
  ProxyApiKeyFieldLabel: string;
  ProxyApiKeyFieldDescription: string;
  BackendApiResourceFieldLabel: string;
  BackendApiResourceFieldDescription: string;
  ProxyBackendFieldLabel: string;
  ProxyBackendReasoningOptionLabel: string;
  ProxyBackendFastOptionLabel: string;
  DeploymentPrefixFieldLabel: string;
  DeploymentPrefixFieldDescription: string;
  M365McpGroupName: string;
  McpEnvironmentIdFieldLabel: string;
  McpEnvironmentIdFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;

  // ─── TextInput ───────────────────────────────────────────────
  TypeAMessage: string;
  SendMessage: string;

  // ─── ActionPanel ─────────────────────────────────────────────
  FocusButton: string;
  SummarizeButton: string;
  ChatButton: string;
  RecapButton: string;
  RecapLoadingButton: string;
  ShareButton: string;
  EmptyPanelHint: string;

  // ─── AvatarPanel status ──────────────────────────────────────
  NoApiStatus: string;
  ApiPendingStatus: string;
  ApiStatus: string;
  VoicePendingStatus: string;
  VoiceStatus: string;
  MicPendingStatus: string;
  MicStatus: string;
  LogsButton: string;
  AvatarSettingsTitle: string;
  PreparingAvatar: string;
  PreparingAvatarDetail: string;
  ConnectVoice: string;
  TurnAudioOff: string;
  NoBackendConfigured: string;

  // ─── AvatarPanel tooltips ────────────────────────────────────
  NoBackendTooltip: string;
  CheckingApiTooltip: string;
  BackendConnectedTooltip: string;
  BackendUnavailableTooltip: string;
  ConnectingVoiceTooltip: string;
  VoiceDisconnectedTooltip: string;
  ClickToDisconnectTooltip: string;
  MicAllowedTooltip: string;
  MicRequestingTooltip: string;
  MicBlockedTooltip: string;

  // ─── StatusBar ───────────────────────────────────────────────
  StatusConnected: string;
  StatusListening: string;
  StatusConnecting: string;
  StatusError: string;
  StatusIdle: string;
  StatusSpeaking: string;
  StatusOffline: string;
  BackendLabel: string;
  VoiceLabel: string;
  MicLabel: string;
  McpLabel: string;
  McpServersTooltip: string;
  NoMcpServersTooltip: string;
  MicAllowed: string;
  MicBlocked: string;

  // ─── Header ──────────────────────────────────────────────────
  GrimoireTitle: string;
  CloseLogsLabel: string;
  OpenLogsLabel: string;

  // ─── AppLayout ───────────────────────────────────────────────
  LoadingAvatarSurface: string;
  LoadingAssistantTools: string;
  LoadingLogs: string;
  ResizeLogsPaneLabel: string;
  ResizeLogsPaneTitle: string;
  ResizeAvatarActionLabel: string;
  ResizeAvatarActionTitle: string;

  // ─── ErrorBoundary ───────────────────────────────────────────
  RenderError: string;
  RetryButton: string;

  // ─── ConfirmationDialogBlock ─────────────────────────────────
  ConfirmedLabel: string;
  CancelledLabel: string;

  // ─── FormBlock ───────────────────────────────────────────────
  SelectPlaceholder: string;
  SubmittingLabel: string;
  TeamPickerUnavailable: string;
  ChannelPickerUnavailable: string;

  // ─── AvatarSettingsPanel ─────────────────────────────────────
  RecapOptionAuto: string;
  RecapOptionAlways: string;
  RecapOptionOff: string;
  CloseSettingsLabel: string;
  AvatarToggleLabel: string;
  ToggleEnabled: string;
  ToggleDisabled: string;
  PublicWebSearchLabel: string;
  CopilotWebGroundingLabel: string;
  VoiceReconnectHint: string;
  VoiceChangeNextConnect: string;

  // ─── Tool Activity Labels ────────────────────────────────────
  ActivitySearchingSharePoint: string;
  ActivitySearchingPeople: string;
  ActivitySearchingSites: string;
  ActivitySearchingEmails: string;
  ActivityResearchingWeb: string;
  ActivityBrowsingLibrary: string;
  ActivityLoadingFileDetails: string;
  ActivityLoadingSiteInfo: string;
  ActivityLoadingListItems: string;
  ActivityReadingFile: string;
  ActivityReadingEmail: string;
  ActivityReadingMessages: string;
  ActivityConnectingToServer: string;
  ActivityRunningMcpTool: string;
  ActivityListingTools: string;
  ActivityRunningM365Tool: string;
  ActivityLoadingServerCatalog: string;
  ActivityPreparingForm: string;
  ActivityLoadingProfile: string;
  ActivityLoadingRecentDocs: string;
  ActivityLoadingTrendingDocs: string;
  ActivitySavingNote: string;
  ActivityRecallingNotes: string;
  ActivityDeletingNote: string;
}

declare module 'GrimoireAssistantWebPartStrings' {
  const strings: IGrimoireAssistantWebPartStrings;
  export = strings;
}
