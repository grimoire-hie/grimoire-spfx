/**
 * GenericBlockBuilder
 * Constructs typed IBlockData from an ISchemaInference + raw parsed JSON.
 * Replaces the 7 bespoke buildXyz() functions in McpResultMapper.
 */

import { createBlock } from '../../models/IBlock';
import type {
  IBlock,
  IDocumentLibraryData,
  ISearchResultsData,
  ISearchResult,
  IFilePreviewData,
  ISiteInfoData,
  IListItemsData,
  IUserCardData,
  ISelectionListData,
  IPermissionsViewData,
  IActivityFeedData,
  IInfoCardData,
  IMarkdownData
} from '../../models/IBlock';
import { deriveMcpTargetContextFromUnknown } from './McpTargetContext';
import type { ISchemaInference } from './SchemaInferrer';

// ─── Helpers ────────────────────────────────────────────────────

function str(obj: Record<string, unknown>, field: string): string {
  if (!field) return '';
  const val = obj[field];
  if (val === null || val === undefined) return '';
  if (typeof val === 'string') return val;
  return String(val);
}

function num(obj: Record<string, unknown>, field: string): number | undefined {
  if (!field) return undefined;
  const val = obj[field];
  return typeof val === 'number' ? val : undefined;
}

function toArray(parsed: unknown): Record<string, unknown>[] {
  if (Array.isArray(parsed)) {
    return parsed.filter((item): item is Record<string, unknown> =>
      typeof item === 'object' && item !== null && !Array.isArray(item)
    );
  }
  if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
    const obj = parsed as Record<string, unknown>;
    const container = obj.value || obj.items || obj.results || obj.data;
    if (Array.isArray(container)) {
      return container.filter((item): item is Record<string, unknown> =>
        typeof item === 'object' && item !== null && !Array.isArray(item)
      );
    }
  }
  return [];
}

function toRecord(parsed: unknown): Record<string, unknown> {
  if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
    return parsed as Record<string, unknown>;
  }
  return {};
}

function nestedRecord(obj: Record<string, unknown>, field: string): Record<string, unknown> {
  const value = obj[field];
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
}

function tryParseJson(text: string): unknown | undefined {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function flattenListValue(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    const parsed = tryParseJson(trimmed);
    if (parsed !== undefined) {
      return flattenListValue(parsed);
    }
    return trimmed || undefined;
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  if (Array.isArray(value)) {
    const flattened = value
      .map((v) => flattenListValue(v))
      .filter((v): v is string => typeof v === 'string' && v.length > 0);
    if (flattened.length === 0) return undefined;
    return flattened.join(', ');
  }

  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>;

    const user = obj.user;
    if (user && typeof user === 'object' && !Array.isArray(user)) {
      const nestedDisplayName = flattenListValue((user as Record<string, unknown>).displayName);
      if (nestedDisplayName) return nestedDisplayName;
    }

    const dateTime = obj.dateTime;
    if (typeof dateTime === 'string' && dateTime.trim()) return dateTime;

    const displayName = obj.displayName;
    if (typeof displayName === 'string' && displayName.trim()) return displayName;

    const name = obj.name;
    if (typeof name === 'string' && name.trim()) return name;

    const address = obj.address;
    if (typeof address === 'string' && address.trim()) return address;

    const emailAddress = obj.emailAddress;
    if (emailAddress && typeof emailAddress === 'object' && !Array.isArray(emailAddress)) {
      const nestedAddress = (emailAddress as Record<string, unknown>).address;
      if (typeof nestedAddress === 'string' && nestedAddress.trim()) return nestedAddress;
    }

    const template = obj.template;
    if (typeof template === 'string' && template.trim()) return template;

    const webUrl = obj.webUrl;
    if (typeof webUrl === 'string' && webUrl.trim()) return webUrl;

    try {
      const compact = JSON.stringify(obj);
      if (!compact || compact === '{}' || compact === '[]') return undefined;
      return compact.length > 180 ? `${compact.slice(0, 177)}...` : compact;
    } catch {
      return undefined;
    }
  }

  return undefined;
}

function bool(obj: Record<string, unknown>, field: string): boolean | undefined {
  const value = obj[field];
  return typeof value === 'boolean' ? value : undefined;
}

function formatYesNo(value: boolean | undefined): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  return value ? 'Yes' : 'No';
}

const SHAREPOINT_COLUMN_TYPE_LABELS: ReadonlyArray<{
  key: string;
  label: string;
  resolve?: (value: Record<string, unknown>) => string;
}> = [
  {
    key: 'text',
    label: 'Text',
    resolve: (value) => bool(value, 'allowMultipleLines') ? 'Multiple lines of text' : 'Text'
  },
  { key: 'number', label: 'Number' },
  { key: 'choice', label: 'Choice' },
  { key: 'boolean', label: 'Yes/No' },
  { key: 'dateTime', label: 'Date and time' },
  { key: 'personOrGroup', label: 'Person or Group' },
  { key: 'hyperlinkOrPicture', label: 'Link or Picture' },
  { key: 'lookup', label: 'Lookup' }
];

function detectSharePointColumnType(obj: Record<string, unknown>): string | undefined {
  for (let i = 0; i < SHAREPOINT_COLUMN_TYPE_LABELS.length; i++) {
    const candidate = SHAREPOINT_COLUMN_TYPE_LABELS[i];
    const rawValue = obj[candidate.key];
    if (!rawValue || typeof rawValue !== 'object' || Array.isArray(rawValue)) {
      continue;
    }
    const record = rawValue as Record<string, unknown>;
    return candidate.resolve ? candidate.resolve(record) : candidate.label;
  }
  return undefined;
}

function isSharePointColumnDefinition(obj: Record<string, unknown>): boolean {
  return !!(str(obj, 'displayName') || str(obj, 'name'))
    && !!(str(obj, 'columnGroup') || detectSharePointColumnType(obj));
}

function buildSharePointColumnInfoCardSummary(obj: Record<string, unknown>): {
  heading: string;
  body: string;
  targetContext?: ReturnType<typeof deriveMcpTargetContextFromUnknown>;
} | undefined {
  if (!isSharePointColumnDefinition(obj)) {
    return undefined;
  }

  const heading = str(obj, 'displayName') || str(obj, 'name') || 'Column';
  const internalName = str(obj, 'name');
  const description = str(obj, 'description');
  const typeLabel = detectSharePointColumnType(obj);
  const lines: string[] = [];

  if (typeLabel) lines.push(`Type: ${typeLabel}`);
  if (internalName) lines.push(`Internal name: ${internalName}`);
  if (description) lines.push(`Description: ${description}`);

  const flagLines: Array<[string, boolean | undefined]> = [
    ['Required', bool(obj, 'required')],
    ['Unique values', bool(obj, 'enforceUniqueValues')],
    ['Hidden', bool(obj, 'hidden')],
    ['Indexed', bool(obj, 'indexed')],
    ['Read-only', bool(obj, 'readOnly')]
  ];
  flagLines.forEach(([label, value]) => {
    const formatted = formatYesNo(value);
    if (formatted) {
      lines.push(`${label}: ${formatted}`);
    }
  });

  const choiceConfig = nestedRecord(obj, 'choice');
  const choiceValues = flattenListValue(choiceConfig.choices);
  if (choiceValues) {
    lines.push(`Choices: ${choiceValues}`);
  }

  if (lines.length === 0) {
    return undefined;
  }

  return {
    heading,
    body: lines.join('\n'),
    targetContext: deriveMcpTargetContextFromUnknown(obj, 'recovered')
  };
}

const LIST_COLUMN_BLACKLIST = new Set<string>([
  '@odata.context',
  '@odata.etag',
  'etag',
  'parentReference'
]);

const LIST_COLUMN_PRIORITY = [
  'displayName',
  'title',
  'subject',
  'name',
  'description',
  'template',
  'hidden',
  'createdDateTime',
  'lastModifiedDateTime',
  'createdBy',
  'lastModifiedBy',
  'webUrl',
  'url',
  'id'
] as const;

function normalizeListRow(item: Record<string, unknown>): Record<string, string> {
  const row: Record<string, string> = {};

  Object.keys(item).forEach((key) => {
    if (LIST_COLUMN_BLACKLIST.has(key)) {
      return;
    }

    const value = item[key];
    if (key === 'list' && value && typeof value === 'object' && !Array.isArray(value)) {
      const listInfo = value as Record<string, unknown>;
      const template = flattenListValue(listInfo.template);
      const hidden = typeof listInfo.hidden === 'boolean' ? String(listInfo.hidden) : undefined;
      if (template) row.template = template;
      if (hidden) row.hidden = hidden;
      return;
    }

    const flattened = flattenListValue(value);
    if (flattened) {
      row[key] = flattened;
    }
  });

  if (row.displayName && row.name && row.displayName === row.name) {
    delete row.name;
  }

  return row;
}

function buildListColumns(rows: Record<string, string>[]): string[] {
  const allColumnsSet = new Set<string>();
  rows.forEach((row) => {
    Object.keys(row).forEach((key) => allColumnsSet.add(key));
  });

  const allColumns = Array.from(allColumnsSet);
  if (allColumns.length === 0) return [];

  const preferred = LIST_COLUMN_PRIORITY.filter((key) => allColumnsSet.has(key));
  const remainder = allColumns
    .filter((key) => !preferred.includes(key as typeof preferred[number]))
    .sort((a, b) => a.localeCompare(b));
  return [...preferred, ...remainder].slice(0, 7);
}

function humanizeToolLabel(toolName: string): string {
  const withSpaces = toolName
    .replace(/[_-]+/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim();
  const withoutLeadingVerb = withSpaces.replace(/^(list|get|find|search)\s+/i, '');
  const candidate = withoutLeadingVerb || withSpaces || toolName;
  return candidate.charAt(0).toUpperCase() + candidate.slice(1);
}

// ─── Block Builders ─────────────────────────────────────────────

function buildUserCard(parsed: unknown, fm: Record<string, string>): IBlock {
  const obj = toRecord(parsed);
  const displayName = str(obj, fm.displayName) || 'User';
  const phone = fm.phone
    ? (Array.isArray(obj[fm.phone]) ? String((obj[fm.phone] as string[])[0] || '') : str(obj, fm.phone))
    : undefined;

  const data: IUserCardData = {
    kind: 'user-card',
    displayName,
    email: str(obj, fm.email),
    jobTitle: str(obj, fm.jobTitle) || undefined,
    department: str(obj, fm.department) || undefined,
    officeLocation: str(obj, fm.officeLocation) || undefined,
    phone: phone || undefined
  };
  return createBlock('user-card', displayName, data);
}

function buildSiteInfo(parsed: unknown, fm: Record<string, string>): IBlock {
  const obj = toRecord(parsed);
  const siteName = str(obj, fm.siteName) || 'Site';
  const data: ISiteInfoData = {
    kind: 'site-info',
    siteName,
    siteUrl: str(obj, fm.siteUrl),
    description: str(obj, fm.description) || undefined,
    created: str(obj, fm.created) || undefined,
    lastModified: str(obj, fm.lastModified) || undefined
  };
  return createBlock('site-info', siteName, data);
}

function buildFilePreview(parsed: unknown, fm: Record<string, string>, rawText: string): IBlock {
  const obj = toRecord(parsed);
  const fileName = str(obj, fm.fileName) || 'File';
  // Build metadata from all fields
  const metadata: Record<string, string> = {};
  Object.keys(obj).forEach((key) => {
    if (obj[key] !== null && obj[key] !== undefined) {
      metadata[key] = String(obj[key]);
    }
  });
  // If parsed is empty, use raw text
  if (Object.keys(metadata).length === 0 && rawText) {
    metadata.content = rawText;
  }
  const data: IFilePreviewData = {
    kind: 'file-preview',
    fileName,
    fileUrl: str(obj, fm.fileUrl),
    fileType: str(obj, fm.fileType),
    size: num(obj, fm.size),
    lastModified: str(obj, fm.lastModified) || undefined,
    author: str(obj, fm.author) || undefined,
    metadata
  };
  return createBlock('file-preview', fileName, data);
}

function buildDocumentLibrary(parsed: unknown, fm: Record<string, string>, toolName: string): IBlock {
  const items = toArray(parsed);
  const data: IDocumentLibraryData = {
    kind: 'document-library',
    siteName: '',
    libraryName: toolName === 'listDocumentLibraries' ? 'Libraries' : 'Contents',
    items: items.map((item) => {
      const parentReference = nestedRecord(item, 'parentReference');
      return {
        name: str(item, fm.name) || str(item, 'name') || str(item, 'displayName'),
        type: (item.folder || item.type === 'folder') ? 'folder' as const : 'file' as const,
        url: str(item, fm.url) || str(item, 'webUrl'),
        documentLibraryId: str(item, 'documentLibraryId') || str(item, 'driveId') || str(parentReference, 'driveId') || undefined,
        fileOrFolderId: str(item, 'fileOrFolderId') || str(item, 'id') || str(item, 'itemId') || undefined,
        size: num(item, fm.size || 'size'),
        lastModified: str(item, fm.lastModified || 'lastModifiedDateTime') || undefined,
        author: str(item, fm.author || 'createdBy') || undefined,
        fileType: item.file
          ? String((item.file as Record<string, unknown>).mimeType || '')
          : str(item, fm.fileType || 'fileType') || undefined
      };
    }),
    breadcrumb: []
  };
  const title = toolName === 'listDocumentLibraries' ? 'Document Libraries' : 'Folder Contents';
  return createBlock('document-library', title, data);
}

function buildSearchResults(parsed: unknown, fm: Record<string, string>, toolName: string): IBlock {
  const items = toArray(parsed);
  const results: ISearchResult[] = items.map((item) => ({
    title: str(item, fm.title) || str(item, 'name') || str(item, 'title'),
    summary: str(item, fm.summary) || str(item, 'summary') || str(item, 'bodyPreview'),
    url: str(item, fm.url) || str(item, 'webUrl'),
    fileType: str(item, fm.fileType || 'fileType') || undefined,
    lastModified: str(item, fm.lastModified || 'lastModifiedDateTime') || undefined,
    author: str(item, fm.author || 'createdBy') || undefined
  }));
  const data: ISearchResultsData = {
    kind: 'search-results',
    query: toolName,
    results,
    totalCount: results.length,
    source: 'MCP'
  };
  return createBlock('search-results', `Results: ${toolName}`, data);
}

function buildPermissionsView(parsed: unknown, fm: Record<string, string>): IBlock {
  const items = toArray(parsed);
  const data: IPermissionsViewData = {
    kind: 'permissions-view',
    targetName: '',
    targetUrl: '',
    permissions: items.map((item) => {
      let principal = str(item, fm.principal);
      // Handle Graph-style grantedTo: { user: { displayName } }
      if (!principal && item.grantedTo) {
        const grantedTo = item.grantedTo as Record<string, unknown>;
        if (grantedTo.user) {
          principal = String((grantedTo.user as Record<string, unknown>).displayName || '');
        }
      }
      let role = str(item, fm.role);
      // Handle Graph-style roles: ["read"]
      if (!role && Array.isArray(item.roles)) {
        role = (item.roles as string[]).join(', ');
      }
      return {
        principal,
        role,
        inherited: Boolean(item.inherited || item.inheritedFrom)
      };
    })
  };
  return createBlock('permissions-view', 'Permissions', data);
}

function buildActivityFeed(parsed: unknown, fm: Record<string, string>): IBlock {
  const items = toArray(parsed);
  const data: IActivityFeedData = {
    kind: 'activity-feed',
    activities: items.map((item) => ({
      actor: str(item, fm.actor),
      action: str(item, fm.action),
      target: str(item, fm.target),
      timestamp: str(item, fm.timestamp) || str(item, 'createdDateTime')
    }))
  };
  return createBlock('activity-feed', 'Activity', data);
}

function buildSelectionList(parsed: unknown, fm: Record<string, string>, toolName: string): IBlock {
  const items = toArray(parsed);
  const data: ISelectionListData = {
    kind: 'selection-list',
    prompt: toolName,
    items: items.map((item) => ({
      id: str(item, fm.id || 'id'),
      label: str(item, fm.label) || str(item, 'displayName') || str(item, 'name'),
      description: str(item, fm.description) || undefined,
      selected: false
    })),
    multiSelect: false
  };
  return createBlock('selection-list', toolName, data);
}

function buildListItems(parsed: unknown, toolName: string): IBlock {
  const items = toArray(parsed);
  const rows: Record<string, string>[] = items.map((item) => normalizeListRow(item));
  const columns = buildListColumns(rows);
  const listLabel = humanizeToolLabel(toolName);
  const data: IListItemsData = {
    kind: 'list-items',
    listName: listLabel,
    columns,
    items: rows,
    totalCount: items.length
  };
  return createBlock('list-items', listLabel, data);
}

function buildEntityInfoCardSummary(obj: Record<string, unknown>): {
  heading: string;
  body: string;
  url?: string;
  targetContext?: ReturnType<typeof deriveMcpTargetContextFromUnknown>;
} | undefined {
  const heading = str(obj, 'displayName') || str(obj, 'name') || str(obj, 'title');
  if (!heading) {
    return undefined;
  }

  const url = str(obj, 'webUrl') || str(obj, 'url');
  const description = str(obj, 'description');
  const listInfo = nestedRecord(obj, 'list');
  const template = flattenListValue(listInfo.template) || flattenListValue(obj.template);
  const created = str(obj, 'createdDateTime');
  const updated = str(obj, 'lastModifiedDateTime');

  const lines: string[] = [];
  if (template) lines.push(`Template: ${template}`);
  if (description) lines.push(`Description: ${description}`);
  if (created) lines.push(`Created: ${created}`);
  if (updated && updated !== created) lines.push(`Updated: ${updated}`);
  if (url) lines.push(`Open: ${url}`);

  if (lines.length === 0) {
    return undefined;
  }

  return {
    heading,
    body: lines.join('\n'),
    url: url || undefined,
    targetContext: deriveMcpTargetContextFromUnknown(obj, 'recovered')
  };
}

function buildInfoCard(parsed: unknown, fm: Record<string, string>, toolName: string): IBlock {
  const obj = toRecord(parsed);
  const columnSummary = buildSharePointColumnInfoCardSummary(obj);
  if (columnSummary && (toolName.toLowerCase() === 'createlistcolumn' || toolName.toLowerCase() === 'editlistcolumn')) {
    const data: IInfoCardData = {
      kind: 'info-card',
      heading: columnSummary.heading,
      body: columnSummary.body,
      targetContext: columnSummary.targetContext
    };
    return createBlock('info-card', columnSummary.heading, data);
  }

  const entitySummary = buildEntityInfoCardSummary(obj);
  if (entitySummary) {
    const data: IInfoCardData = {
      kind: 'info-card',
      heading: entitySummary.heading,
      body: entitySummary.body,
      url: entitySummary.url,
      targetContext: entitySummary.targetContext
    };
    return createBlock('info-card', entitySummary.heading, data);
  }

  const statusValue = str(obj, fm.heading);
  const heading = statusValue === 'true' ? `${toolName}: Success`
    : statusValue === 'false' ? `${toolName}: Failed`
    : statusValue || toolName;
  const data: IInfoCardData = {
    kind: 'info-card',
    heading,
    body: str(obj, fm.body) || JSON.stringify(parsed, null, 2)
  };
  return createBlock('info-card', heading, data);
}

function buildMarkdown(rawText: string, toolName: string): IBlock {
  const data: IMarkdownData = {
    kind: 'markdown',
    content: rawText
  };
  return createBlock('markdown', `MCP: ${toolName}`, data);
}

// ─── Public API ─────────────────────────────────────────────────

/**
 * Build a typed IBlock from a schema inference and raw parsed data.
 *
 * @param inference - The inferred block type, confidence, and field mapping
 * @param parsed - The parsed JSON data (or undefined if not JSON)
 * @param rawText - The raw text content from the MCP result
 * @param toolName - The tool name (used for titles and context)
 */
export function build(
  inference: ISchemaInference,
  parsed: unknown,
  rawText: string,
  toolName: string
): IBlock {
  const { blockType, fieldMap } = inference;

  switch (blockType) {
    case 'user-card':
      return buildUserCard(parsed, fieldMap);
    case 'site-info':
      return buildSiteInfo(parsed, fieldMap);
    case 'file-preview':
      return buildFilePreview(parsed, fieldMap, rawText);
    case 'document-library':
      return buildDocumentLibrary(parsed, fieldMap, toolName);
    case 'search-results':
      return buildSearchResults(parsed, fieldMap, toolName);
    case 'permissions-view':
      return buildPermissionsView(parsed, fieldMap);
    case 'activity-feed':
      return buildActivityFeed(parsed, fieldMap);
    case 'selection-list':
      return buildSelectionList(parsed, fieldMap, toolName);
    case 'list-items':
      return buildListItems(parsed, toolName);
    case 'info-card':
      return buildInfoCard(parsed, fieldMap, toolName);
    default:
      return buildMarkdown(rawText, toolName);
  }
}
