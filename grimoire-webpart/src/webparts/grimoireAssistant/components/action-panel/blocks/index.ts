/**
 * Block Registry
 * Maps BlockType to React components for rendering in ActionPanel.
 * All 16 block types registered.
 */

import * as React from 'react';
import type { BlockType, IBlockData, IRenderHints } from '../../../models/IBlock';

import { SearchResultsBlock } from './SearchResultsBlock';
import { DocumentLibraryBlock } from './DocumentLibraryBlock';
import { FilePreviewBlock } from './FilePreviewBlock';
import { SiteInfoBlock } from './SiteInfoBlock';
import { InfoCardBlock } from './InfoCardBlock';
import { MarkdownBlock } from './MarkdownBlock';
import { ErrorBlock } from './ErrorBlock';
import { ProgressTrackerBlock } from './ProgressTrackerBlock';
import { UserCardBlock } from './UserCardBlock';
import { ListItemsBlock } from './ListItemsBlock';
import { PermissionsViewBlock } from './PermissionsViewBlock';
import { ActivityFeedBlock } from './ActivityFeedBlock';
import { ChartBlock } from './ChartBlock';
import { ConfirmationDialogBlock } from './ConfirmationDialogBlock';
import { SelectionListBlock } from './SelectionListBlock';
import { FormBlock } from './FormBlock';

/**
 * Registry mapping block type → React component.
 * Each component receives `{ data: T }` where T is the typed data interface.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const BLOCK_REGISTRY: Partial<Record<BlockType, React.FC<{ data: any; blockId?: string; renderHints?: IRenderHints }>>> = {
  'search-results': SearchResultsBlock,
  'document-library': DocumentLibraryBlock,
  'file-preview': FilePreviewBlock,
  'site-info': SiteInfoBlock,
  'info-card': InfoCardBlock,
  'markdown': MarkdownBlock,
  'error': ErrorBlock,
  'progress-tracker': ProgressTrackerBlock,
  'user-card': UserCardBlock,
  'list-items': ListItemsBlock,
  'permissions-view': PermissionsViewBlock,
  'activity-feed': ActivityFeedBlock,
  'chart': ChartBlock,
  'confirmation-dialog': ConfirmationDialogBlock,
  'selection-list': SelectionListBlock,
  'form': FormBlock
};

/**
 * Get the block renderer component for a given block type.
 * Returns undefined if no component is registered (fallback to placeholder).
 */
export function getBlockComponent(type: BlockType): React.FC<{ data: IBlockData; blockId?: string; renderHints?: IRenderHints }> | undefined {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return BLOCK_REGISTRY[type] as React.FC<{ data: any; blockId?: string; renderHints?: IRenderHints }> | undefined;
}
