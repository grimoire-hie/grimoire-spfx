/**
 * BlockTracker — Maintains metadata for all active UI blocks.
 * Tracks block lifecycle (loading → ready → acknowledged → interacted → dismissed)
 * and provides a compact summary of the current visual state.
 */

import {
  IBlock,
  BlockType,
  ISearchResultsData,
  IDocumentLibraryData,
  IFilePreviewData,
  ISiteInfoData,
  IUserCardData,
  IListItemsData,
  IActivityFeedData,
  IConfirmationDialogData,
  ISelectionListData,
  IErrorData,
  IInfoCardData,
  IProgressTrackerData,
  IPermissionsViewData,
  IChartData,
  IMarkdownData,
  IFormData
} from '../../models/IBlock';
import { ITrackedBlock, IBlockReference, IHieTurnLineage } from './HIETypes';

export class BlockTracker {
  private blocks: Map<string, ITrackedBlock> = new Map();

  public track(block: IBlock, turnLineage?: IHieTurnLineage): void {
    const isLoading = this.isBlockLoading(block);
    this.blocks.set(block.id, {
      id: block.id,
      type: block.type,
      title: block.title,
      originTool: block.originTool,
      turnId: turnLineage?.turnId,
      rootTurnId: turnLineage?.rootTurnId || turnLineage?.turnId,
      parentTurnId: turnLineage?.parentTurnId,
      summary: isLoading ? '' : this.summarizeBlock(block),
      itemCount: this.countItems(block),
      references: isLoading ? [] : this.extractReferences(block),
      state: isLoading ? 'loading' : 'ready',
      contextInjected: false,
      createdAt: Date.now(),
      updatedAt: Date.now()
    });
  }

  public update(blockId: string, block: IBlock): void {
    const tracked = this.blocks.get(blockId);
    if (!tracked) return;
    tracked.originTool = block.originTool;
    tracked.summary = this.summarizeBlock(block);
    tracked.itemCount = this.countItems(block);
    tracked.references = this.extractReferences(block);
    tracked.state = 'ready';
    tracked.updatedAt = Date.now();
  }

  public markAcknowledged(blockId: string): void {
    const tracked = this.blocks.get(blockId);
    if (tracked && tracked.state === 'ready') {
      tracked.state = 'acknowledged';
      tracked.updatedAt = Date.now();
    }
  }

  public markInteracted(blockId: string): void {
    const tracked = this.blocks.get(blockId);
    if (tracked) {
      tracked.state = 'interacted';
      tracked.updatedAt = Date.now();
    }
  }

  public markContextInjected(blockId: string): void {
    const tracked = this.blocks.get(blockId);
    if (tracked) {
      tracked.contextInjected = true;
    }
  }

  public remove(blockId: string): void {
    const tracked = this.blocks.get(blockId);
    if (tracked) {
      tracked.state = 'dismissed';
    }
    this.blocks.delete(blockId);
  }

  public get(blockId: string): ITrackedBlock | undefined {
    return this.blocks.get(blockId);
  }

  public getReady(): ITrackedBlock[] {
    const result: ITrackedBlock[] = [];
    this.blocks.forEach((tracked) => {
      if (tracked.state === 'ready' && !tracked.contextInjected) {
        result.push(tracked);
      }
    });
    return result;
  }

  public getActiveSummary(): string {
    if (this.blocks.size === 0) return '';
    const parts: string[] = [];
    this.blocks.forEach((tracked) => {
      if (tracked.state !== 'dismissed' && tracked.summary) {
        parts.push(tracked.summary);
      }
    });
    return parts.join(' | ');
  }

  public clear(): void {
    this.blocks.clear();
  }

  public getSize(): number {
    return this.blocks.size;
  }

  /**
   * Get numbered references for a specific block.
   */
  public getReferences(blockId: string): IBlockReference[] {
    const tracked = this.blocks.get(blockId);
    return tracked ? tracked.references : [];
  }

  /**
   * Get all numbered references across all active blocks, ordered by block creation time.
   * Each reference includes the blockId in its detail for disambiguation.
   */
  public getAllReferences(): { blockId: string; blockType: BlockType; references: IBlockReference[] }[] {
    const result: { blockId: string; blockType: BlockType; references: IBlockReference[] }[] = [];
    this.blocks.forEach((tracked) => {
      if (tracked.state !== 'dismissed' && tracked.references.length > 0) {
        result.push({
          blockId: tracked.id,
          blockType: tracked.type,
          references: tracked.references
        });
      }
    });
    // Sort newest first so the most recent block's references take priority
    result.sort((a, b) => {
      const aTracked = this.blocks.get(a.blockId);
      const bTracked = this.blocks.get(b.blockId);
      return (bTracked?.updatedAt || 0) - (aTracked?.updatedAt || 0);
    });
    return result;
  }

  /**
   * Return active tracked blocks sorted by most recently updated first.
   */
  public getActiveBlocks(): ITrackedBlock[] {
    const result: ITrackedBlock[] = [];
    this.blocks.forEach((tracked) => {
      if (tracked.state !== 'dismissed') {
        result.push({ ...tracked, references: tracked.references.slice() });
      }
    });
    result.sort((a, b) => b.updatedAt - a.updatedAt);
    return result;
  }

  // ─── Private Helpers ────────────────────────────────────────

  private isBlockLoading(block: IBlock): boolean {
    if (block.type === 'search-results') {
      const data = block.data as ISearchResultsData;
      return data.source === 'pending' || (data.results.length === 0 && data.totalCount === 0);
    }
    if (block.type === 'document-library') {
      return (block.data as IDocumentLibraryData).items.length === 0;
    }
    if (block.type === 'permissions-view') {
      return (block.data as IPermissionsViewData).permissions.length === 0;
    }
    if (block.type === 'activity-feed') {
      return (block.data as IActivityFeedData).activities.length === 0;
    }
    return false;
  }

  private summarizeBlock(block: IBlock): string {
    switch (block.type) {
      case 'search-results': {
        const d = block.data as ISearchResultsData;
        if (d.results.length === 0) return `Search '${d.query}': no results`;
        const top = d.results.slice(0, 3).map((r, i) => `${i + 1}) ${r.title}`).join(', ');
        return `Search '${d.query}': ${d.totalCount} results. Top: ${top}`;
      }
      case 'document-library': {
        const d = block.data as IDocumentLibraryData;
        const folders = d.items.filter((i) => i.type === 'folder').length;
        const files = d.items.filter((i) => i.type === 'file').length;
        return `Library '${d.libraryName}': ${folders} folders, ${files} files`;
      }
      case 'file-preview': {
        const d = block.data as IFilePreviewData;
        return `File: ${d.fileName} (${d.fileType}${d.author ? ', by ' + d.author : ''})`;
      }
      case 'site-info': {
        const d = block.data as ISiteInfoData;
        return `Site: ${d.siteName}${d.description ? ' — ' + d.description.slice(0, 60) : ''}`;
      }
      case 'user-card': {
        const d = block.data as IUserCardData;
        return `Person: ${d.displayName}${d.jobTitle ? ', ' + d.jobTitle : ''}${d.department ? ', ' + d.department : ''}`;
      }
      case 'list-items': {
        const d = block.data as IListItemsData;
        return `List '${d.listName}': ${d.totalCount} items, columns: ${d.columns.join(', ')}`;
      }
      case 'permissions-view': {
        const d = block.data as IPermissionsViewData;
        return `Permissions for '${d.targetName}': ${d.permissions.length} entries`;
      }
      case 'activity-feed': {
        const d = block.data as IActivityFeedData;
        return `Activity feed: ${d.activities.length} activities`;
      }
      case 'chart': {
        const d = block.data as IChartData;
        return `Chart: ${d.title} (${d.chartType})`;
      }
      case 'confirmation-dialog': {
        const d = block.data as IConfirmationDialogData;
        return `Confirmation: "${d.message}"`;
      }
      case 'selection-list': {
        const d = block.data as ISelectionListData;
        return `Selection: "${d.prompt}" (${d.items.length} options)`;
      }
      case 'progress-tracker': {
        const d = block.data as IProgressTrackerData;
        return `Progress: ${d.label} ${d.progress}% (${d.status})`;
      }
      case 'error': {
        const d = block.data as IErrorData;
        return `Error: ${d.message}`;
      }
      case 'info-card': {
        const d = block.data as IInfoCardData;
        return `Info: ${d.heading}`;
      }
      case 'markdown': {
        const md = block.data as IMarkdownData;
        const mdRefs = this.extractMarkdownReferences(md.content);
        if (mdRefs.length > 0) {
          const top = mdRefs.slice(0, 3).map((r) => `${r.index}) ${r.title}`).join(', ');
          return `${block.title}: ${mdRefs.length} items. Top: ${top}`;
        }
        return block.title || 'Content block displayed';
      }
      case 'form': {
        const d = block.data as IFormData;
        return `Form (${d.preset}): ${d.status}`;
      }
      default:
        return `${block.type} block displayed`;
    }
  }

  private countItems(block: IBlock): number {
    switch (block.type) {
      case 'search-results':
        return (block.data as ISearchResultsData).results.length;
      case 'document-library':
        return (block.data as IDocumentLibraryData).items.length;
      case 'list-items':
        return (block.data as IListItemsData).items.length;
      case 'activity-feed':
        return (block.data as IActivityFeedData).activities.length;
      case 'permissions-view':
        return (block.data as IPermissionsViewData).permissions.length;
      case 'selection-list':
        return (block.data as ISelectionListData).items.length;
      case 'markdown':
        return this.extractMarkdownReferences((block.data as IMarkdownData).content).length;
      default:
        return 0;
    }
  }

  /**
   * Extract numbered item references from a block.
   * These enable the LLM to resolve positional references like "the second one".
   */
  private extractReferences(block: IBlock): IBlockReference[] {
    switch (block.type) {
      case 'search-results': {
        const d = block.data as ISearchResultsData;
        return d.results.map((r, i) => {
          const parts: string[] = [];
          if (r.author) parts.push(`by ${r.author}`);
          if (r.fileType) parts.push(`.${r.fileType}`);
          if (r.sources && r.sources.length > 0) parts.push(`[${r.sources.join('+')}]`);
          return {
            index: i + 1,
            title: r.title,
            url: r.url,
            itemType: r.fileType || 'document',
            detail: parts.length > 0 ? parts.join(', ') : undefined
          };
        });
      }
      case 'document-library': {
        const d = block.data as IDocumentLibraryData;
        return d.items.map((item, i) => ({
          index: i + 1,
          title: item.name,
          url: item.url,
          itemType: item.type,
          detail: item.fileType || undefined
        }));
      }
      case 'list-items': {
        const d = block.data as IListItemsData;
        return d.items.map((item, i) => {
          // Use first column value as title, rest as detail
          const vals = d.columns.map((col) => item[col] || '').filter(Boolean);
          return {
            index: i + 1,
            title: vals[0] || `Item ${i + 1}`,
            itemType: 'list-item',
            detail: vals.slice(1, 3).join(', ') || undefined
          };
        });
      }
      case 'selection-list': {
        const d = block.data as ISelectionListData;
        return d.items.map((item, i) => ({
          index: i + 1,
          title: item.label,
          itemType: 'option',
          detail: item.description || undefined
        }));
      }
      case 'activity-feed': {
        const d = block.data as IActivityFeedData;
        return d.activities.map((a, i) => ({
          index: i + 1,
          title: `${a.actor} ${a.action} ${a.target}`,
          itemType: 'activity',
          detail: a.timestamp
        }));
      }
      case 'permissions-view': {
        const d = block.data as IPermissionsViewData;
        return d.permissions.map((p, i) => ({
          index: i + 1,
          title: p.principal,
          itemType: 'permission',
          detail: `${p.role}${p.inherited ? ' (inherited)' : ''}`
        }));
      }
      case 'markdown': {
        const md = block.data as IMarkdownData;
        return this.extractMarkdownReferences(md.content);
      }
      default:
        return [];
    }
  }

  /**
   * Extract numbered references from markdown content.
   * Handles Agent 365 format: `1. **Subject:** *title*\n   **From:** sender\n...`
   * Extracts **Key:** value pairs, uses Subject/Event/Message as title.
   * Caps at 20 items.
   */
  private extractMarkdownReferences(content: string): IBlockReference[] {
    if (!content) return [];

    // Split on numbered item boundaries: "1." or "1)" at line start
    const parts = content.split(/(?:^|\n)\s*\d+[.)]\s*/);
    const items = parts.slice(1);
    if (items.length === 0) return [];

    const refs: IBlockReference[] = [];
    const cap = Math.min(items.length, 20);

    for (let i = 0; i < cap; i++) {
      const item = items[i];

      // Extract all **Key:** value pairs (value may contain *italic*, [links](url), etc.)
      const fields: Record<string, string> = {};
      const fieldPattern = /\*\*([^*]+)\*\*\s*([^\n]+)/g;
      let fm: RegExpExecArray | null;
      // eslint-disable-next-line no-cond-assign
      while ((fm = fieldPattern.exec(item)) !== null) {
        const key = fm[1].replace(/:$/, '').trim().toLowerCase();
        // Strip italic markers, markdown links [text](url), and trailing whitespace
        const val = fm[2]
          .replace(/^\*+|\*+$/g, '')
          .replace(/\s*\[[^\]]*\]\([^)]*\)/g, '')
          .trim();
        if (val) fields[key] = val;
      }

      // Title: prefer Subject, Event, Message, From (in that order)
      const title = fields.subject || fields.event || fields.message
        || fields.from || Object.values(fields)[0] || '';
      if (!title) continue;

      // Detail: From + Date (skip if same as title)
      const detailParts: string[] = [];
      if (fields.from && fields.from !== title) detailParts.push(fields.from);
      if (fields.date && detailParts.length < 2) detailParts.push(fields.date);

      refs.push({
        index: i + 1,
        title,
        itemType: 'item',
        detail: detailParts.length > 0 ? detailParts.join(', ') : undefined
      });
    }

    return refs;
  }
}
