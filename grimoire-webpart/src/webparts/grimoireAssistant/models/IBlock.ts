/**
 * IBlock — UI Block types for the ActionPanel.
 * Each block is a typed data structure that the ActionPanel renders
 * using the block registry.
 */

import type { SearchQueryVariantKind } from './ISearchTypes';
import type { IMcpTargetContext } from '../services/mcp/McpTargetContext';

export type BlockType =
  | 'search-results'
  | 'document-library'
  | 'file-preview'
  | 'site-info'
  | 'user-card'
  | 'list-items'
  | 'permissions-view'
  | 'activity-feed'
  | 'chart'
  | 'confirmation-dialog'
  | 'selection-list'
  | 'progress-tracker'
  | 'error'
  | 'info-card'
  | 'markdown'
  | 'form';

// ─── Render Hints ───────────────────────────────────────────────

export interface IRenderHints {
  /** 1-based indices of items to highlight with accent border */
  highlight?: number[];
  /** 1-based index → annotation text shown as badge */
  annotate?: Record<number, string>;
  /** Render collapsed with expand button */
  collapse?: boolean;
}

export interface IBlock {
  /** Unique block ID */
  id: string;
  /** Block type (determines which component renders) */
  type: BlockType;
  /** Block title shown in the ActionPanel */
  title: string;
  /** Timestamp when the block was created */
  timestamp: Date;
  /** Whether the block can be dismissed by the user */
  dismissible: boolean;
  /** Typed data payload — shape depends on `type` */
  data: IBlockData;
  /** Optional render hints from LLM (highlight, annotate, collapse) */
  renderHints?: IRenderHints;
  /** Optional schema ID used by interaction adapter for this block */
  interactionSchemaId?: string;
  /** Optional origin tool name that produced this block */
  originTool?: string;
}

// ─── Block Data Interfaces ──────────────────────────────────────

export type IBlockData =
  | ISearchResultsData
  | IDocumentLibraryData
  | IFilePreviewData
  | ISiteInfoData
  | IUserCardData
  | IListItemsData
  | IPermissionsViewData
  | IActivityFeedData
  | IChartData
  | IConfirmationDialogData
  | ISelectionListData
  | IProgressTrackerData
  | IErrorData
  | IInfoCardData
  | IMarkdownData
  | IFormData;

export interface ISearchResultsData {
  kind: 'search-results';
  query: string;
  results: ISearchResult[];
  totalCount: number;
  source: string;
  queryVariants?: ISearchQueryVariantInfo[];
}

export interface ISearchResult {
  title: string;
  summary: string;
  url: string;
  lastModified?: string;
  author?: string;
  fileType?: string;
  siteName?: string;
  thumbnailUrl?: string;
  /** Which search APIs contributed this result (e.g., ['copilot-search', 'sharepoint-search']) */
  sources?: string[];
  language?: string;
}

export interface ISearchQueryVariantInfo {
  kind: SearchQueryVariantKind;
  query: string;
  language?: string;
}

export interface IDocumentLibraryData {
  kind: 'document-library';
  siteName: string;
  libraryName: string;
  items: IDocumentItem[];
  breadcrumb: string[];
}

export interface IDocumentItem {
  name: string;
  type: 'file' | 'folder';
  url: string;
  documentLibraryId?: string;
  fileOrFolderId?: string;
  size?: number;
  lastModified?: string;
  author?: string;
  fileType?: string;
}

export interface IFilePreviewData {
  kind: 'file-preview';
  fileName: string;
  fileUrl: string;
  fileType: string;
  size?: number;
  lastModified?: string;
  author?: string;
  previewUrl?: string;
  metadata?: Record<string, string>;
}

export interface ISiteInfoData {
  kind: 'site-info';
  siteName: string;
  siteUrl: string;
  description?: string;
  owner?: string;
  created?: string;
  lastModified?: string;
  storageUsed?: string;
  libraries?: string[];
  lists?: string[];
}

export interface IUserCardData {
  kind: 'user-card';
  displayName: string;
  email: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  phone?: string;
  photoUrl?: string;
}

export interface IListItemsData {
  kind: 'list-items';
  listName: string;
  columns: string[];
  items: Record<string, string>[];
  totalCount: number;
}

export interface IPermissionsViewData {
  kind: 'permissions-view';
  targetName: string;
  targetUrl: string;
  permissions: IPermissionEntry[];
}

export interface IPermissionEntry {
  principal: string;
  role: string;
  inherited: boolean;
}

export interface IActivityFeedData {
  kind: 'activity-feed';
  activities: IActivityItem[];
}

export interface IActivityItem {
  action: string;
  actor: string;
  target: string;
  timestamp: string;
}

export interface IChartData {
  kind: 'chart';
  chartType: 'bar' | 'pie' | 'line';
  title: string;
  labels: string[];
  values: number[];
}

export interface IConfirmationDialogData {
  kind: 'confirmation-dialog';
  message: string;
  confirmLabel: string;
  cancelLabel: string;
  onConfirmAction: string;
}

export interface ISelectionListData {
  kind: 'selection-list';
  prompt: string;
  items: ISelectionItem[];
  multiSelect: boolean;
}

export interface ISelectionItem {
  id: string;
  label: string;
  description?: string;
  selected?: boolean;
  itemType?: string;
  targetContext?: IMcpTargetContext;
}

export type ProgressTrackerStatus = 'running' | 'complete' | 'error';
export type ProgressTrackerStepStatus = 'pending' | 'running' | 'complete' | 'error';

export interface IProgressTrackerStep {
  id: string;
  label: string;
  status: ProgressTrackerStepStatus;
  detail?: string;
  sourceBlockId?: string;
  derivedBlockId?: string;
}

export interface IProgressTrackerData {
  kind: 'progress-tracker';
  label: string;
  progress: number;
  status: ProgressTrackerStatus;
  detail?: string;
  steps?: IProgressTrackerStep[];
  currentStep?: number;
}

export interface IErrorData {
  kind: 'error';
  message: string;
  detail?: string;
  retryAction?: string;
}

export interface IInfoCardData {
  kind: 'info-card';
  heading: string;
  body: string;
  icon?: string;
  url?: string;
  targetContext?: IMcpTargetContext;
}

export interface IMarkdownData {
  kind: 'markdown';
  content: string;
  /** Map of card number → Outlook ItemID (Graph message ID) extracted from Agent 365 links */
  itemIds?: Record<number, string>;
}

// ─── Form Block Types ───────────────────────────────────────────

export type FormFieldType =
  | 'text' | 'textarea' | 'email' | 'email-list'
  | 'datetime' | 'date' | 'toggle' | 'dropdown'
  | 'team-picker' | 'channel-picker' | 'people-picker'
  | 'number' | 'hidden';

export interface IFormFieldVisibilityRule {
  fieldKey: string;
  equals?: string;
  anyOf?: string[];
}

export interface IFormFieldDefinition {
  key: string;
  label: string;
  type: FormFieldType;
  required: boolean;
  placeholder?: string;
  defaultValue?: string;
  options?: Array<{ key: string; text: string }>;
  rows?: number;
  group?: string;
  width?: 'full' | 'half';
  visibleWhen?: IFormFieldVisibilityRule;
}

export type FormPresetId =
  | 'email-compose' | 'email-reply' | 'email-forward' | 'email-reply-all-thread'
  | 'event-create' | 'event-update'
  | 'teams-message' | 'teams-channel-message'
  | 'share-teams-chat' | 'share-teams-channel'
  | 'file-create' | 'word-document-create' | 'folder-create' | 'item-rename'
  | 'list-item-create' | 'list-item-update'
  | 'channel-create' | 'chat-create'
  | 'generic';

export type FormStatus = 'editing' | 'submitting' | 'success' | 'error';

export interface IFormSubmissionTarget {
  toolName: string;
  serverId: string;
  staticArgs: Record<string, unknown>;
  fieldToParamMap?: Record<string, string>;
  targetContext?: IMcpTargetContext;
}

export interface IFormData {
  kind: 'form';
  preset: FormPresetId;
  description?: string;
  fields: IFormFieldDefinition[];
  submissionTarget: IFormSubmissionTarget;
  status: FormStatus;
  errorMessage?: string;
  successMessage?: string;
}

// ─── Block Factory Helper ───────────────────────────────────────

let _blockIdCounter: number = 0;

export function createBlock(
  type: BlockType,
  title: string,
  data: IBlockData,
  dismissible: boolean = true,
  renderHints?: IRenderHints,
  metadata?: { interactionSchemaId?: string; originTool?: string }
): IBlock {
  _blockIdCounter++;
  const block: IBlock = {
    id: `block-${_blockIdCounter}-${Date.now()}`,
    type,
    title,
    timestamp: new Date(),
    dismissible,
    data
  };
  if (renderHints) {
    block.renderHints = renderHints;
  }
  if (metadata?.interactionSchemaId) {
    block.interactionSchemaId = metadata.interactionSchemaId;
  }
  if (metadata?.originTool) {
    block.originTool = metadata.originTool;
  }
  return block;
}
