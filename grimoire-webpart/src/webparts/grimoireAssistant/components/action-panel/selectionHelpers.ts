/**
 * selectionHelpers
 * Builds selection candidates for header-level Focus/Summarize/Chat actions.
 */

import type {
  IBlock,
  BlockType,
  IDocumentLibraryData,
  IFilePreviewData,
  ISiteInfoData,
  IUserCardData,
  ISearchResultsData,
  IListItemsData,
  IPermissionsViewData,
  IActivityFeedData,
  IInfoCardData,
  IChartData,
  IProgressTrackerData,
  IMarkdownData,
  ISelectionListData
} from '../../models/IBlock';
import {
  extractSharePointSiteUrl,
  inferDocumentLibraryBaseUrl,
  resolveDocumentLibraryItemUrl
} from '../../utils/sharePointUrlUtils';

export type ActionableBlockType =
  | 'search-results'
  | 'document-library'
  | 'file-preview'
  | 'site-info'
  | 'user-card'
  | 'list-items'
  | 'selection-list'
  | 'permissions-view'
  | 'activity-feed'
  | 'chart'
  | 'info-card'
  | 'markdown'
  | 'progress-tracker';

export type SelectionEntityKind =
  | 'document'
  | 'file'
  | 'folder'
  | 'email'
  | 'calendar-event'
  | 'message'
  | 'list-item'
  | 'person'
  | 'site'
  | 'permission'
  | 'activity'
  | 'info'
  | 'chart'
  | 'progress'
  | 'item';

export interface ISelectionCandidate {
  index: number;
  title: string;
  kind: SelectionEntityKind;
  payload: Record<string, unknown>;
  url?: string;
  itemType?: string;
}

export function isActionableBlockType(type: BlockType): type is ActionableBlockType {
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

interface INumberedGroup {
  num: number;
  lines: string[];
}

function splitIntoNumberedGroups(content: string): INumberedGroup[] {
  const lines = content.split('\n');
  const groups: INumberedGroup[] = [];
  let currentGroup: INumberedGroup | undefined;

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    const numMatch = /^(\d+)[.)]\s*(.*)$/.exec(trimmed);
    if (numMatch) {
      if (currentGroup) groups.push(currentGroup);
      currentGroup = { num: parseInt(numMatch[1], 10), lines: [] };
      if (numMatch[2]) currentGroup.lines.push(numMatch[2]);
      continue;
    }
    if (!currentGroup) continue;
    if (/^[-*_]{3,}$/.test(trimmed)) continue;
    currentGroup.lines.push(lines[i]);
  }
  if (currentGroup) groups.push(currentGroup);
  return groups;
}

export function extractGroupFields(lines: string[]): Record<string, string> {
  const fields: Record<string, string> = {};
  let fallbackTitle = '';
  let firstUrl = '';

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (!trimmed) continue;
    if (!firstUrl) {
      const urlMatch = /\[[^\]]*\]\((https?:\/\/[^)]+)\)/.exec(trimmed);
      if (urlMatch) firstUrl = urlMatch[1];
    }
    const kvMatch = /^\*\*([^*]+)\*\*[:\s]*(.+)/.exec(trimmed);
    if (kvMatch) {
      const key = kvMatch[1].replace(/:$/, '').trim();
      const val = kvMatch[2]
        .replace(/\*+/g, '')
        .replace(/\[[^\]]*\]\([^)]*\)/g, '')
        .trim();
      if (val) fields[key] = val;
    } else if (!fallbackTitle) {
      fallbackTitle = trimmed.replace(/\*\*/g, '').trim();
    }
  }

  if (Object.keys(fields).length === 0 && fallbackTitle) {
    fields.Title = fallbackTitle;
  }
  if (firstUrl) fields.url = firstUrl;
  return fields;
}

function detectMarkdownKind(fields: Record<string, string>): SelectionEntityKind {
  const hasEmail = !!(fields.Subject || fields.From || fields.Preview || fields.Cc || fields.To);
  if (hasEmail) return 'email';

  const hasEvent = !!(fields.Event || fields.Start || fields.End || fields.Location || fields.Attendees);
  if (hasEvent) return 'calendar-event';

  const hasMessage = !!(fields.Message || fields.Channel || fields.Chat || fields.Replies);
  if (hasMessage) return 'message';

  if (typeof fields.url === 'string' && fields.url.trim()) return 'document';
  return 'item';
}

function getSearchSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as ISearchResultsData;
  return data.results.map((result, idx) => {
    const index = idx + 1;
    return {
      index,
      title: result.title || `Result ${index}`,
      kind: result.fileType === 'folder' ? 'folder' : 'document',
      url: result.url,
      itemType: result.fileType || 'document',
      payload: {
        index,
        title: result.title,
        url: result.url,
        fileType: result.fileType,
        author: result.author,
        summary: result.summary,
        siteName: result.siteName
      }
    };
  });
}

function getListSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IListItemsData;
  return data.items.map((item, idx) => {
    const index = idx + 1;
    const vals = data.columns.map((col) => item[col] || '').filter(Boolean);
    const title = vals[0] || `Item ${index}`;
    return {
      index,
      title,
      kind: 'list-item',
      itemType: 'list-item',
      payload: {
        index,
        title,
        rowData: item
      }
    };
  });
}

function getSelectionListCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as ISelectionListData;
  return data.items.map((item, idx) => {
    const index = idx + 1;
    return {
      index,
      title: item.label || `Item ${index}`,
      kind: 'item',
      itemType: 'selection-item',
      payload: {
        index,
        title: item.label,
        label: item.label,
        description: item.description,
        id: item.id,
        prompt: data.prompt
      }
    };
  });
}

function getMarkdownSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IMarkdownData;
  const groups = splitIntoNumberedGroups(data.content || '');
  const itemIds = data.itemIds || {};

  return groups.map((group) => {
    const fields = extractGroupFields(group.lines);
    const kind = detectMarkdownKind(fields);
    const title = fields.Subject || fields.Event || fields.Message || fields.From
      || fields.Title || Object.values(fields)[0] || `Item ${group.num}`;
    const itemId = itemIds[group.num];
    const url = typeof fields.url === 'string' ? fields.url : undefined;

    return {
      index: group.num,
      title,
      kind,
      url,
      itemType: kind,
      payload: {
        index: group.num,
        title,
        kind,
        ...fields,
        itemId
      }
    };
  });
}

function getDocumentLibrarySelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IDocumentLibraryData;
  const siblingUrls = data.items.map((item) => item.url);
  const documentLibraryUrl = inferDocumentLibraryBaseUrl(siblingUrls);
  const siteUrl = extractSharePointSiteUrl(documentLibraryUrl || siblingUrls[0]);
  const siteName = /^https?:\/\//i.test(data.siteName || '') ? undefined : data.siteName;
  return data.items.map((item, idx) => {
    const index = idx + 1;
    const kind: SelectionEntityKind = item.type === 'folder' ? 'folder' : 'file';
    const resolvedUrl = resolveDocumentLibraryItemUrl(item.url, item.name, siblingUrls) || item.url;
    return {
      index,
      title: item.name || `Item ${index}`,
      kind,
      url: resolvedUrl,
      itemType: item.type,
      payload: {
        index,
        name: item.name,
        title: item.name,
        url: resolvedUrl,
        fileOrFolderUrl: resolvedUrl,
        fileOrFolderId: item.fileOrFolderId,
        documentLibraryId: item.documentLibraryId,
        documentLibraryUrl,
        documentLibraryName: data.libraryName,
        siteUrl,
        siteName,
        type: item.type,
        fileType: item.fileType,
        author: item.author,
        lastModified: item.lastModified,
        size: item.size
      }
    };
  });
}

function getFilePreviewSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IFilePreviewData;
  return [{
    index: 1,
    title: data.fileName || 'File',
    kind: 'file',
    url: data.fileUrl,
    itemType: data.fileType || 'file',
    payload: {
      index: 1,
      title: data.fileName,
      fileName: data.fileName,
      fileType: data.fileType,
      url: data.fileUrl,
      author: data.author,
      lastModified: data.lastModified
    }
  }];
}

function getSiteInfoSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as ISiteInfoData;
  return [{
    index: 1,
    title: data.siteName || 'Site',
    kind: 'site',
    url: data.siteUrl,
    itemType: 'site',
    payload: {
      index: 1,
      title: data.siteName,
      name: data.siteName,
      url: data.siteUrl,
      description: data.description,
      owner: data.owner
    }
  }];
}

function getUserCardSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IUserCardData;
  return [{
    index: 1,
    title: data.displayName || data.email || 'Person',
    kind: 'person',
    itemType: 'person',
    payload: {
      index: 1,
      title: data.displayName,
      displayName: data.displayName,
      email: data.email,
      jobTitle: data.jobTitle,
      department: data.department
    }
  }];
}

function getPermissionsSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IPermissionsViewData;
  return data.permissions.map((permission, idx) => {
    const index = idx + 1;
    return {
      index,
      title: permission.principal || `Principal ${index}`,
      kind: 'permission',
      itemType: 'permission',
      payload: {
        index,
        principal: permission.principal,
        role: permission.role,
        inherited: permission.inherited,
        targetName: data.targetName,
        targetUrl: data.targetUrl
      }
    };
  });
}

function getActivitySelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IActivityFeedData;
  return data.activities.map((activity, idx) => {
    const index = idx + 1;
    const title = `${activity.actor} ${activity.action} ${activity.target}`.trim();
    return {
      index,
      title: title || `Activity ${index}`,
      kind: 'activity',
      itemType: 'activity',
      payload: {
        index,
        title,
        actor: activity.actor,
        action: activity.action,
        target: activity.target,
        timestamp: activity.timestamp
      }
    };
  });
}

function getInfoCardSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IInfoCardData;
  const url = data.url
    || data.targetContext?.listUrl
    || data.targetContext?.fileOrFolderUrl
    || data.targetContext?.documentLibraryUrl
    || data.targetContext?.siteUrl;
  return [{
    index: 1,
    title: data.heading || 'Information',
    kind: 'info',
    url,
    itemType: 'info',
    payload: {
      index: 1,
      title: data.heading,
      heading: data.heading,
      body: data.body,
      icon: data.icon,
      url,
      targetContext: data.targetContext
    }
  }];
}

function getChartSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IChartData;
  return [{
    index: 1,
    title: data.title || 'Chart',
    kind: 'chart',
    itemType: 'chart',
    payload: {
      index: 1,
      title: data.title,
      chartType: data.chartType,
      labels: data.labels,
      values: data.values
    }
  }];
}

function getProgressSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  const data = block.data as IProgressTrackerData;
  return [{
    index: 1,
    title: data.label || 'Progress',
    kind: 'progress',
    itemType: 'progress',
    payload: {
      index: 1,
      title: data.label,
      label: data.label,
      progress: data.progress,
      status: data.status,
      detail: data.detail
    }
  }];
}

export function getSelectionCandidates(block: IBlock): ISelectionCandidate[] {
  if (!isActionableBlockType(block.type)) return [];

  switch (block.type) {
    case 'search-results':
      return getSearchSelectionCandidates(block);
    case 'document-library':
      return getDocumentLibrarySelectionCandidates(block);
    case 'file-preview':
      return getFilePreviewSelectionCandidates(block);
    case 'site-info':
      return getSiteInfoSelectionCandidates(block);
    case 'user-card':
      return getUserCardSelectionCandidates(block);
    case 'list-items':
      return getListSelectionCandidates(block);
    case 'selection-list':
      return getSelectionListCandidates(block);
    case 'permissions-view':
      return getPermissionsSelectionCandidates(block);
    case 'activity-feed':
      return getActivitySelectionCandidates(block);
    case 'chart':
      return getChartSelectionCandidates(block);
    case 'info-card':
      return getInfoCardSelectionCandidates(block);
    case 'markdown':
      return getMarkdownSelectionCandidates(block);
    case 'progress-tracker':
      return getProgressSelectionCandidates(block);
    default:
      return [];
  }
}
