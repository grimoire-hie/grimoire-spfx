/**
 * useGrimoireStore — Zustand store for Grimoire Assistant
 * ~15 focused fields for connection state, avatar, config, conversation,
 * action panel blocks, MCP connections, and logging.
 */

import { create } from 'zustand';
import type { Expression } from '../services/avatar/ExpressionEngine';
import type { PersonalityMode } from '../services/avatar/PersonalityEngine';
import type { VisageMode } from '../services/avatar/FaceTemplateData';
import type { BlockType, IBlock } from '../models/IBlock';
import type { IMcpConnection } from '../models/IMcpTypes';
import type { ILogEntry } from '../services/logging/LogTypes';
import type { LogCategory } from '../services/logging/LogTypes';
import type { AadHttpClient } from '@microsoft/sp-http';
import type { IUserContext } from '../services/context/ContextService';
import type { ConversationLanguage } from '../services/context/ConversationLanguage';
import type { SearchRecapMode } from '../services/context/AssistantPreferenceUtils';
import type { AvatarActionCueType, IAvatarActionCue } from '../services/avatar/AvatarActionCue';
import { createCorrelationId } from '../services/hie/HAEContracts';

// ─── Transcript ─────────────────────────────────────────────────

export interface ITranscriptEntry {
  entryId?: string;
  turnId?: string;
  rootTurnId?: string;
  parentTurnId?: string;
  text: string;
  role: 'user' | 'assistant' | 'system';
  timestamp: Date;
}

// ─── Action Focus ───────────────────────────────────────────────

export interface IFocusedItem {
  index: number;
  title: string;
  kind?: string;
  url?: string;
  itemType?: string;
  payload?: Record<string, unknown>;
}

export interface IFocusedContext {
  blockId: string;
  blockType: BlockType;
  blockTitle: string;
  itemCount: number;
  items: IFocusedItem[];
  updatedAt: number;
}

function isActionableBlockType(type: BlockType): boolean {
  return type === 'search-results'
    || type === 'document-library'
    || type === 'file-preview'
    || type === 'site-info'
    || type === 'user-card'
    || type === 'list-items'
    || type === 'selection-list'
    || type === 'permissions-view'
    || type === 'activity-feed'
    || type === 'chart'
    || type === 'info-card'
    || type === 'markdown'
    || type === 'progress-tracker';
}

// ─── Proxy Config ───────────────────────────────────────────────

export interface IProxyConfig {
  proxyUrl: string;
  proxyApiKey: string;
  backendApiResource?: string;
  backend: string;
  deployment: string;
  apiVersion: string;
}

// ─── SP Theme Colors ────────────────────────────────────────────

export interface ISpThemeColors {
  bodyBackground: string;
  bodyText: string;
  bodySubtext: string;
  cardBackground: string;
  cardBorder: string;
  isDark: boolean;
}

export const DEFAULT_SP_THEME_COLORS: ISpThemeColors = {
  bodyBackground: '#ffffff',
  bodyText: '#323130',
  bodySubtext: '#605e5c',
  cardBackground: '#ffffff',
  cardBorder: '#edebe9',
  isDark: false
};

export const DEFAULT_LOG_SIDEBAR_WIDTH = 420;

// ─── UI Mode ────────────────────────────────────────────────────

export type UiMode = 'widget' | 'active';
export type PublicWebSearchCapabilityStatus = 'unknown' | 'available' | 'blocked' | 'unsupported' | 'error';
export type StartupPhase = 'bootstrap-widget' | 'active-shell' | 'deferred-runtimes';
export type StartupCacheSource = 'none' | 'network' | 'session-cache';
export type AvatarRenderState = 'placeholder' | 'svg-loading' | 'bindings-ready' | 'animated-ready';
export type StartupMetricName =
  | 'widget-render'
  | 'start-click'
  | 'active-shell-paint'
  | 'deferred-runtimes-start'
  | 'health-check'
  | 'graph-enrichment'
  | 'preferences-hydration'
  | 'public-web-probe'
  | 'svg-fetch'
  | 'svg-bind'
  | 'first-animated-frame';

export interface IStartupMetric {
  status: 'recorded' | 'running' | 'completed';
  startedAtMs?: number;
  durationMs?: number;
  completedAtMs?: number;
  detail?: string;
}

// ─── Connection State ───────────────────────────────────────────

export type ConnectionState = 'idle' | 'connecting' | 'connected' | 'speaking' | 'error';
export type AssistantPlaybackState = 'idle' | 'buffering' | 'playing' | 'error';

// ─── Store State ────────────────────────────────────────────────

export interface IGrimoireState {
  // Connection & Session
  connectionState: ConnectionState;

  // UI Mode & Theme
  uiMode: UiMode;
  spThemeColors: ISpThemeColors;
  startupPhase: StartupPhase;
  startupMetrics: Partial<Record<StartupMetricName, IStartupMetric>>;

  // Avatar
  avatarEnabled: boolean;
  avatarRenderState: AvatarRenderState;
  personality: PersonalityMode;
  visage: VisageMode;
  expression: Expression;
  gazeTarget: 'none' | 'action-panel';
  avatarActionCue: IAvatarActionCue | undefined;
  remoteStream: MediaStream | undefined;
  micStream: MediaStream | undefined;
  isMuted: boolean;
  backendOk: boolean;
  micGranted: boolean;

  // Config
  proxyConfig: IProxyConfig | undefined;
  aadHttpClient: AadHttpClient | undefined;
  userApiClient: AadHttpClient | undefined;
  userApiClientReady: boolean;
  /** Async function to acquire a bearer token for a given resource URL */
  getToken: ((resource: string) => Promise<string>) | undefined;
  /** Power Platform environment GUID for Agent 365 MCP servers */
  mcpEnvironmentId: string | undefined;
  /** User identity + environment context (built from pageContext + Graph /me) */
  userContext: IUserContext | undefined;

  // Conversation (session-scoped, cleared on disconnect)
  transcript: ITranscriptEntry[];
  textInputValue: string;
  /** True while TextChatService is processing an HTTP chat response */
  textChatActive: boolean;
  /** Activity status shown below avatar (e.g. "Searching...", "Thinking...") */
  activityStatus: string;
  /** Session-level conversation language, updated when the user switches language. */
  conversationLanguage: ConversationLanguage | undefined;

  // Voice
  /** Selected voice ID for realtime sessions (user preference, persists across sessions) */
  voiceId: string;
  /** Assistant playback state for the active voice session. */
  assistantPlaybackState: AssistantPlaybackState;
  /** Whether explicit public-web / URL research is enabled (user preference). */
  publicWebSearchEnabled: boolean;
  /** Session-scoped capability status for Azure public web search. */
  publicWebSearchCapability: PublicWebSearchCapabilityStatus;
  /** Optional detail explaining why public web search is blocked/unsupported/error. */
  publicWebSearchCapabilityDetail: string | undefined;
  /** Timestamp for the most recent backend health evaluation. */
  lastHealthCheckAt: number | undefined;
  /** Source used for the latest backend health evaluation. */
  lastHealthCheckSource: StartupCacheSource;
  /** Timestamp for the latest public web search capability evaluation. */
  publicWebSearchCapabilityCheckedAt: number | undefined;
  /** Source used for the latest public web search capability evaluation. */
  publicWebSearchCapabilitySource: StartupCacheSource;
  /** Whether Copilot Chat API requests can use web grounding (user preference). */
  copilotWebGroundingEnabled: boolean;
  /** Whether search results should include an automatic recap block. */
  searchRecapMode: SearchRecapMode;
  /** True once avatar settings have been hydrated from backend preferences. */
  avatarPrefsHydrated: boolean;
  /** True once a user explicitly changed avatar-related settings locally. */
  avatarSettingsDirty: boolean;
  /** True once a user explicitly changed assistant-related settings locally. */
  assistantSettingsDirty: boolean;

  // Action Panel
  blocks: IBlock[];
  /** Block IDs that haven't been acknowledged (moused over) by the user */
  freshBlockIds: string[];
  /** Most recent actionable block (selection-capable list) */
  activeActionBlockId: string | undefined;
  /** 1-based selected item indices within activeActionBlockId */
  selectedActionIndices: number[];
  /** Local focus context set instantly from header Focus action */
  focusedContext: IFocusedContext | undefined;

  // MCP
  mcpConnections: IMcpConnection[];

  // Logging
  logEntries: ILogEntry[];
  logSidebarOpen: boolean;
  logSidebarWidth: number;
  logFilter: LogCategory | 'all';
}

// ─── Store Actions ──────────────────────────────────────────────

export interface IGrimoireActions {
  // Connection
  setConnectionState: (state: ConnectionState) => void;

  // UI Mode & Theme
  setUiMode: (mode: UiMode) => void;
  setSpThemeColors: (colors: ISpThemeColors) => void;
  setStartupPhase: (phase: StartupPhase) => void;
  updateStartupMetric: (name: StartupMetricName, metric: IStartupMetric) => void;

  // Avatar
  setAvatarEnabled: (enabled: boolean) => void;
  setAvatarRenderState: (state: AvatarRenderState) => void;
  setPersonality: (mode: PersonalityMode) => void;
  setVisage: (visage: VisageMode) => void;
  setExpression: (expression: Expression) => void;
  setGazeTarget: (target: 'none' | 'action-panel') => void;
  triggerAvatarActionCue: (type: AvatarActionCueType) => void;
  setRemoteStream: (stream: MediaStream | undefined) => void;
  setMicStream: (stream: MediaStream | undefined) => void;
  setMuted: (muted: boolean) => void;
  setBackendOk: (ok: boolean) => void;
  setHealthCheckState: (ok: boolean, checkedAt: number, source: StartupCacheSource) => void;
  setMicGranted: (granted: boolean) => void;

  // Config
  setProxyConfig: (config: IProxyConfig | undefined) => void;
  setAadHttpClient: (client: AadHttpClient | undefined) => void;
  setUserApiClient: (client: AadHttpClient | undefined) => void;
  setUserApiClientReady: (ready: boolean) => void;
  setGetToken: (fn: ((resource: string) => Promise<string>) | undefined) => void;
  setMcpEnvironmentId: (id: string | undefined) => void;
  setUserContext: (ctx: IUserContext | undefined) => void;

  // Conversation
  addTranscript: (entry: ITranscriptEntry) => void;
  updateLastTranscript: (text: string) => void;
  clearTranscript: () => void;
  setTextInputValue: (value: string) => void;
  setTextChatActive: (active: boolean) => void;
  setActivityStatus: (status: string) => void;
  setConversationLanguage: (language: ConversationLanguage | undefined) => void;

  // Voice
  setVoiceId: (id: string) => void;
  setAssistantPlaybackState: (state: AssistantPlaybackState) => void;
  setPublicWebSearchEnabled: (enabled: boolean) => void;
  setPublicWebSearchCapability: (
    status: PublicWebSearchCapabilityStatus,
    detail?: string,
    checkedAt?: number,
    source?: StartupCacheSource
  ) => void;
  setCopilotWebGroundingEnabled: (enabled: boolean) => void;
  setSearchRecapMode: (mode: SearchRecapMode) => void;
  setAvatarPrefsHydrated: (hydrated: boolean) => void;
  markAvatarSettingsDirty: () => void;
  clearAvatarSettingsDirty: () => void;
  markAssistantSettingsDirty: () => void;
  clearAssistantSettingsDirty: () => void;

  // Action Panel
  pushBlock: (block: IBlock) => void;
  insertBlockAfter: (referenceBlockId: string, block: IBlock) => void;
  updateBlock: (blockId: string, updates: Partial<IBlock>) => void;
  removeBlock: (blockId: string) => void;
  clearBlocks: () => void;
  acknowledgeBlock: (blockId: string) => void;
  setActiveActionBlock: (blockId: string | undefined) => void;
  toggleActionSelection: (blockId: string, index: number) => void;
  setActionSelection: (blockId: string, indices: number[]) => void;
  clearActionSelection: () => void;
  setFocusedContext: (context: IFocusedContext | undefined) => void;

  // MCP
  addMcpConnection: (connection: IMcpConnection) => void;
  updateMcpConnection: (sessionId: string, updates: Partial<IMcpConnection>) => void;
  removeMcpConnection: (sessionId: string) => void;

  // Logging
  addLogEntry: (entry: ILogEntry) => void;
  clearLogEntries: () => void;
  setLogSidebarOpen: (open: boolean) => void;
  setLogSidebarWidth: (width: number) => void;
  setLogFilter: (filter: LogCategory | 'all') => void;

  // Session reset
  resetSession: () => void;
}

// ─── Store ──────────────────────────────────────────────────────

let avatarActionCueCounter = 0;

export const useGrimoireStore = create<IGrimoireState & IGrimoireActions>((set) => ({
  // ─── Initial State ──────────────────────────────────────────
  connectionState: 'idle',

  uiMode: 'widget',
  spThemeColors: DEFAULT_SP_THEME_COLORS,
  startupPhase: 'bootstrap-widget',
  startupMetrics: {},

  avatarEnabled: true,
  avatarRenderState: 'placeholder',
  personality: 'normal',
  visage: 'classic',
  expression: 'idle',
  gazeTarget: 'none',
  avatarActionCue: undefined,
  remoteStream: undefined,
  micStream: undefined,
  isMuted: false,
  backendOk: false,
  micGranted: false,

  proxyConfig: undefined,
  aadHttpClient: undefined,
  userApiClient: undefined,
  userApiClientReady: false,
  getToken: undefined,
  mcpEnvironmentId: undefined,
  userContext: undefined,

  transcript: [],
  textInputValue: '',
  textChatActive: false,
  activityStatus: '',
  conversationLanguage: undefined,

  voiceId: 'alloy',
  assistantPlaybackState: 'idle',
  publicWebSearchEnabled: false,
  publicWebSearchCapability: 'unknown',
  publicWebSearchCapabilityDetail: undefined,
  lastHealthCheckAt: undefined,
  lastHealthCheckSource: 'none',
  publicWebSearchCapabilityCheckedAt: undefined,
  publicWebSearchCapabilitySource: 'none',
  copilotWebGroundingEnabled: false,
  searchRecapMode: 'off',
  avatarPrefsHydrated: false,
  avatarSettingsDirty: false,
  assistantSettingsDirty: false,

  blocks: [],
  freshBlockIds: [],
  activeActionBlockId: undefined,
  selectedActionIndices: [],
  focusedContext: undefined,

  mcpConnections: [],

  logEntries: [],
  logSidebarOpen: false,
  logSidebarWidth: DEFAULT_LOG_SIDEBAR_WIDTH,
  logFilter: 'all',

  // ─── Connection Actions ─────────────────────────────────────
  setConnectionState: (connectionState) => set({ connectionState }),

  // ─── UI Mode & Theme Actions ──────────────────────────────
  setUiMode: (uiMode) => set((state) => ({
    uiMode,
    startupPhase: uiMode === 'widget' ? 'bootstrap-widget' : state.startupPhase
  })),
  setSpThemeColors: (spThemeColors) => set({ spThemeColors }),
  setStartupPhase: (startupPhase) => set({ startupPhase }),
  updateStartupMetric: (name, metric) =>
    set((state) => ({
      startupMetrics: {
        ...state.startupMetrics,
        [name]: metric
      }
    })),

  // ─── Avatar Actions ─────────────────────────────────────────
  setAvatarEnabled: (avatarEnabled) =>
    set((state) => {
      if (state.avatarEnabled === avatarEnabled) {
        return state;
      }

      return avatarEnabled
        ? {
          avatarEnabled,
          avatarRenderState: 'placeholder'
        }
        : {
          avatarEnabled,
          avatarRenderState: 'placeholder',
          expression: 'idle',
          gazeTarget: 'none',
          avatarActionCue: undefined
        };
    }),
  setAvatarRenderState: (avatarRenderState) => set({ avatarRenderState }),
  setPersonality: (personality) => set({ personality }),
  setVisage: (visage) =>
    set((state) => (
      state.visage === visage
        ? state
        : {
          visage,
          avatarRenderState: 'placeholder'
        }
    )),
  setExpression: (expression) => set({ expression }),
  setGazeTarget: (gazeTarget) => set({ gazeTarget }),
  triggerAvatarActionCue: (type) =>
    set({
      avatarActionCue: {
        id: ++avatarActionCueCounter,
        type,
        at: Date.now()
      }
    }),
  setRemoteStream: (remoteStream) => set({ remoteStream }),
  setMicStream: (micStream) => set({ micStream }),
  setMuted: (isMuted) => set({ isMuted }),
  setBackendOk: (backendOk) => set({ backendOk }),
  setHealthCheckState: (backendOk, lastHealthCheckAt, lastHealthCheckSource) => set({
    backendOk,
    lastHealthCheckAt,
    lastHealthCheckSource
  }),
  setMicGranted: (micGranted) => set({ micGranted }),

  // ─── Config Actions ─────────────────────────────────────────
  setProxyConfig: (proxyConfig) => set({
    proxyConfig,
    avatarPrefsHydrated: false,
    avatarSettingsDirty: false,
    assistantSettingsDirty: false,
    publicWebSearchCapability: 'unknown',
    publicWebSearchCapabilityDetail: undefined,
    lastHealthCheckAt: undefined,
    lastHealthCheckSource: 'none',
    publicWebSearchCapabilityCheckedAt: undefined,
    publicWebSearchCapabilitySource: 'none'
  }),
  setAadHttpClient: (aadHttpClient) => set({ aadHttpClient }),
  setUserApiClient: (userApiClient) => set({ userApiClient }),
  setUserApiClientReady: (userApiClientReady) => set({ userApiClientReady }),
  setGetToken: (getToken) => set({ getToken }),
  setMcpEnvironmentId: (mcpEnvironmentId) => set({ mcpEnvironmentId }),
  setUserContext: (userContext) => set({ userContext }),

  // ─── Conversation Actions ───────────────────────────────────
  addTranscript: (entry) =>
    set((state) => {
      const normalizedEntry: ITranscriptEntry = {
        ...entry,
        entryId: entry.entryId || createCorrelationId('msg'),
        rootTurnId: entry.rootTurnId || entry.turnId
      };
      const next = [...state.transcript, normalizedEntry];
      return { transcript: next.length > 100 ? next.slice(-100) : next };
    }),
  updateLastTranscript: (text) =>
    set((state) => {
      const { transcript } = state;
      if (transcript.length === 0) return state;
      const last = transcript[transcript.length - 1];
      if (last.role !== 'assistant') return state;
      const updated = [...transcript];
      updated[updated.length - 1] = { ...last, text };
      return { transcript: updated };
    }),
  clearTranscript: () => set({ transcript: [] }),
  setTextInputValue: (textInputValue) => set({ textInputValue }),
  setTextChatActive: (textChatActive) => set({ textChatActive }),
  setActivityStatus: (activityStatus) => set({ activityStatus }),
  setConversationLanguage: (conversationLanguage) => set({ conversationLanguage }),

  // ─── Voice Actions ───────────────────────────────────────────
  setVoiceId: (voiceId) => set({ voiceId }),
  setAssistantPlaybackState: (assistantPlaybackState) => set({ assistantPlaybackState }),
  setPublicWebSearchEnabled: (publicWebSearchEnabled) =>
    set((state) => ({
      publicWebSearchEnabled,
      publicWebSearchCapability: publicWebSearchEnabled ? state.publicWebSearchCapability : 'unknown',
      publicWebSearchCapabilityDetail: publicWebSearchEnabled ? state.publicWebSearchCapabilityDetail : undefined,
      publicWebSearchCapabilityCheckedAt: publicWebSearchEnabled ? state.publicWebSearchCapabilityCheckedAt : undefined,
      publicWebSearchCapabilitySource: publicWebSearchEnabled ? state.publicWebSearchCapabilitySource : 'none'
    })),
  setPublicWebSearchCapability: (
    publicWebSearchCapability,
    publicWebSearchCapabilityDetail,
    publicWebSearchCapabilityCheckedAt,
    publicWebSearchCapabilitySource
  ) =>
    set((state) => ({
      publicWebSearchCapability,
      publicWebSearchCapabilityDetail,
      publicWebSearchCapabilityCheckedAt: typeof publicWebSearchCapabilityCheckedAt === 'number'
        ? publicWebSearchCapabilityCheckedAt
        : state.publicWebSearchCapabilityCheckedAt,
      publicWebSearchCapabilitySource: publicWebSearchCapabilitySource || state.publicWebSearchCapabilitySource
    })),
  setCopilotWebGroundingEnabled: (copilotWebGroundingEnabled) => set({ copilotWebGroundingEnabled }),
  setSearchRecapMode: (searchRecapMode) => set({ searchRecapMode }),
  setAvatarPrefsHydrated: (avatarPrefsHydrated) => set({ avatarPrefsHydrated }),
  markAvatarSettingsDirty: () => set({ avatarSettingsDirty: true }),
  clearAvatarSettingsDirty: () => set({ avatarSettingsDirty: false }),
  markAssistantSettingsDirty: () => set({ assistantSettingsDirty: true }),
  clearAssistantSettingsDirty: () => set({ assistantSettingsDirty: false }),

  // ─── Action Panel Actions ───────────────────────────────────
  pushBlock: (block) =>
    set((state) => {
      const isActionable = isActionableBlockType(block.type);
      return {
        blocks: [...state.blocks, block],
        freshBlockIds: [...state.freshBlockIds, block.id],
        activeActionBlockId: isActionable ? block.id : state.activeActionBlockId,
        selectedActionIndices: isActionable ? [] : state.selectedActionIndices,
        focusedContext: isActionable ? undefined : state.focusedContext
      };
    }),
  insertBlockAfter: (referenceBlockId, block) =>
    set((state) => {
      const preserveActionContext = block.originTool?.startsWith('block-recap:') === true;
      const referenceIndex = state.blocks.findIndex((existing) => existing.id === referenceBlockId);
      if (referenceIndex === -1) {
        const isActionable = isActionableBlockType(block.type);
        return {
          blocks: [...state.blocks, block],
          freshBlockIds: [...state.freshBlockIds, block.id],
          activeActionBlockId: isActionable && !preserveActionContext ? block.id : state.activeActionBlockId,
          selectedActionIndices: isActionable && !preserveActionContext ? [] : state.selectedActionIndices,
          focusedContext: isActionable && !preserveActionContext ? undefined : state.focusedContext
        };
      }

      const blocks = state.blocks.slice();
      blocks.splice(referenceIndex + 1, 0, block);
      const isActionable = isActionableBlockType(block.type);
      return {
        blocks,
        freshBlockIds: [...state.freshBlockIds, block.id],
        activeActionBlockId: isActionable && !preserveActionContext ? block.id : state.activeActionBlockId,
        selectedActionIndices: isActionable && !preserveActionContext ? [] : state.selectedActionIndices,
        focusedContext: isActionable && !preserveActionContext ? undefined : state.focusedContext
      };
    }),
  updateBlock: (blockId, updates) =>
    set((state) => ({
      blocks: state.blocks.map((b) =>
        b.id === blockId ? { ...b, ...updates } : b
      )
    })),
  removeBlock: (blockId) =>
    set((state) => {
      const blocks = state.blocks.filter((b) => b.id !== blockId);
      const freshBlockIds = state.freshBlockIds.filter((id) => id !== blockId);
      const removingActive = state.activeActionBlockId === blockId;
      let activeActionBlockId = state.activeActionBlockId;
      let selectedActionIndices = state.selectedActionIndices;

      if (removingActive) {
        activeActionBlockId = undefined;
        selectedActionIndices = [];
        for (let i = blocks.length - 1; i >= 0; i--) {
          if (isActionableBlockType(blocks[i].type)) {
            activeActionBlockId = blocks[i].id;
            break;
          }
        }
      }

      const focusedContext = state.focusedContext && state.focusedContext.blockId === blockId
        ? undefined
        : state.focusedContext;

      return { blocks, freshBlockIds, activeActionBlockId, selectedActionIndices, focusedContext };
    }),
  clearBlocks: () => set({
    blocks: [],
    freshBlockIds: [],
    activeActionBlockId: undefined,
    selectedActionIndices: [],
    focusedContext: undefined
  }),
  acknowledgeBlock: (blockId) =>
    set((state) => ({
      freshBlockIds: state.freshBlockIds.filter((id) => id !== blockId)
    })),
  setActiveActionBlock: (blockId) =>
    set((state) => ({
      activeActionBlockId: blockId,
      selectedActionIndices: state.activeActionBlockId === blockId ? state.selectedActionIndices : []
    })),
  toggleActionSelection: (blockId, index) =>
    set((state) => {
      if (!blockId || state.activeActionBlockId !== blockId) return state;
      const next = new Set<number>(state.selectedActionIndices);
      if (next.has(index)) next.delete(index);
      else next.add(index);
      const selectedActionIndices = Array.from(next).sort((a, b) => a - b);
      return { selectedActionIndices };
    }),
  setActionSelection: (blockId, indices) =>
    set(() => ({
      activeActionBlockId: blockId,
      selectedActionIndices: Array.from(new Set(indices))
        .filter((n) => Number.isFinite(n) && n > 0)
        .sort((a, b) => a - b)
    })),
  clearActionSelection: () => set({ selectedActionIndices: [] }),
  setFocusedContext: (focusedContext) => set({ focusedContext }),

  // ─── MCP Actions ────────────────────────────────────────────
  addMcpConnection: (connection) =>
    set((state) => ({
      mcpConnections: [...state.mcpConnections, connection]
    })),
  updateMcpConnection: (sessionId, updates) =>
    set((state) => ({
      mcpConnections: state.mcpConnections.map((c) =>
        c.sessionId === sessionId ? { ...c, ...updates } : c
      )
    })),
  removeMcpConnection: (sessionId) =>
    set((state) => ({
      mcpConnections: state.mcpConnections.filter((c) => c.sessionId !== sessionId)
    })),

  // ─── Logging Actions ────────────────────────────────────────
  addLogEntry: (entry) =>
    set((state) => {
      const next = [...state.logEntries, entry];
      return { logEntries: next.length > 500 ? next.slice(-500) : next };
    }),
  clearLogEntries: () => set({ logEntries: [] }),
  setLogSidebarOpen: (logSidebarOpen) => set({ logSidebarOpen }),
  setLogSidebarWidth: (logSidebarWidth) => set({ logSidebarWidth }),
  setLogFilter: (logFilter) => set({ logFilter }),

  // ─── Session Reset ──────────────────────────────────────────
  resetSession: () =>
    set({
      connectionState: 'idle',
      uiMode: 'widget',
      startupPhase: 'bootstrap-widget',
      startupMetrics: {},
      avatarRenderState: 'placeholder',
      expression: 'idle',
      gazeTarget: 'none',
      avatarActionCue: undefined,
      remoteStream: undefined,
      micStream: undefined,
      isMuted: false,
      transcript: [],
      textInputValue: '',
      textChatActive: false,
      activityStatus: '',
      conversationLanguage: undefined,
      assistantPlaybackState: 'idle',
      publicWebSearchCapability: 'unknown',
      publicWebSearchCapabilityDetail: undefined,
      lastHealthCheckAt: undefined,
      lastHealthCheckSource: 'none',
      publicWebSearchCapabilityCheckedAt: undefined,
      publicWebSearchCapabilitySource: 'none',
      blocks: [],
      freshBlockIds: [],
      activeActionBlockId: undefined,
      selectedActionIndices: [],
      focusedContext: undefined
    })
}));

export default useGrimoireStore;
