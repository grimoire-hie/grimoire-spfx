import type { IHieArtifactRecord, IHieSourceContext, IHieTaskContext } from '../hie/HIETypes';
import { resolveCurrentArtifactContext, resolveLatestArtifactContext } from '../hie/HieArtifactLinkage';
import { pickFirstNonEmptyString } from './OdspLocationResolver';

export type McpTargetSource = 'explicit-user' | 'hie-selection' | 'current-page' | 'recovered' | 'none';

export interface IMcpTargetContext {
  source?: McpTargetSource;
  siteUrl?: string;
  siteId?: string;
  siteName?: string;
  documentLibraryId?: string;
  documentLibraryUrl?: string;
  documentLibraryName?: string;
  fileOrFolderId?: string;
  fileOrFolderUrl?: string;
  fileOrFolderName?: string;
  listId?: string;
  listItemId?: string;
  listUrl?: string;
  listName?: string;
  columnId?: string;
  columnName?: string;
  columnDisplayName?: string;
  personIdentifier?: string;
  personEmail?: string;
  personDisplayName?: string;
  mailItemId?: string;
  calendarItemId?: string;
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
}

export type McpTargetContext = IMcpTargetContext;

function tryParseJsonRecord(value: unknown): Record<string, unknown> | undefined {
  if (!value) {
    return undefined;
  }
  if (typeof value === 'object' && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  if (typeof value !== 'string') {
    return undefined;
  }
  try {
    const parsed = JSON.parse(value);
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      return parsed as Record<string, unknown>;
    }
  } catch {
    // Ignore invalid JSON payloads.
  }
  return undefined;
}

function decodeUrlSegment(segment: string | undefined): string | undefined {
  if (!segment) {
    return undefined;
  }
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function extractIdsFromSharePointODataContext(value: unknown): Partial<IMcpTargetContext> {
  if (typeof value !== 'string') {
    return {};
  }

  const siteMatch = value.match(/sites\('([^']+)'\)/i);
  const listMatch = value.match(/lists\('([^']+)'\)/i);
  const decode = (candidate: string | undefined): string | undefined => {
    if (!candidate) {
      return undefined;
    }
    try {
      return decodeURIComponent(candidate);
    } catch {
      return candidate;
    }
  };

  return {
    ...(decode(siteMatch?.[1]) ? { siteId: decode(siteMatch?.[1]) } : {}),
    ...(decode(listMatch?.[1]) ? { listId: decode(listMatch?.[1]) } : {})
  };
}

function deriveSharePointContextFromUrl(urlValue: string): Partial<IMcpTargetContext> {
  try {
    const parsed = new URL(urlValue);
    const pathSegments = parsed.pathname.split('/').filter(Boolean);
    if (pathSegments.length === 0) {
      return {};
    }

    const siteRootIndex = pathSegments.findIndex((segment) => /^(sites|teams)$/i.test(segment));
    const siteSegments = siteRootIndex >= 0 && pathSegments.length > siteRootIndex + 1
      ? pathSegments.slice(siteRootIndex, siteRootIndex + 2)
      : [];
    const sitePath = siteSegments.length > 0 ? `/${siteSegments.join('/')}` : undefined;
    const siteUrl = sitePath ? `${parsed.origin}${sitePath}` : undefined;
    const siteName = decodeUrlSegment(siteSegments[1]);
    const lastSegment = decodeUrlSegment(pathSegments[pathSegments.length - 1]);
    const isListUrl = pathSegments.some((segment) => segment.toLowerCase() === 'lists');
    const hasFileExtension = !!lastSegment && /\.[a-z0-9]{2,8}$/i.test(lastSegment);

    const afterSiteSegments = siteSegments.length > 0
      ? pathSegments.slice(siteRootIndex + 2)
      : pathSegments;
    const firstAfterSite = decodeUrlSegment(afterSiteSegments[0]);
    const firstAfterSiteUrl = sitePath && afterSiteSegments[0]
      ? `${parsed.origin}${sitePath}/${afterSiteSegments[0]}`
      : undefined;

    const derived: Partial<IMcpTargetContext> = {};
    if (siteUrl) {
      derived.siteUrl = siteUrl;
      derived.siteName = siteName;
    }
    if (isListUrl) {
      derived.listUrl = urlValue;
      derived.listName = pickFirstNonEmptyString(lastSegment, firstAfterSite);
    } else if (hasFileExtension) {
      derived.fileOrFolderUrl = urlValue;
      derived.fileOrFolderName = lastSegment;
      if (firstAfterSiteUrl) {
        derived.documentLibraryUrl = firstAfterSiteUrl;
        derived.documentLibraryName = firstAfterSite;
      }
    } else if (firstAfterSiteUrl) {
      derived.documentLibraryUrl = urlValue === firstAfterSiteUrl ? firstAfterSiteUrl : firstAfterSiteUrl;
      derived.documentLibraryName = firstAfterSite;
    }

    return derived;
  } catch {
    return {};
  }
}

function applyStringField(
  target: Partial<IMcpTargetContext>,
  key: keyof IMcpTargetContext,
  ...values: unknown[]
): void {
  const value = pickFirstNonEmptyString(...values);
  if (value) {
    (target as Record<string, unknown>)[key] = value;
  }
}

function deriveTargetContextFromRecord(record: Record<string, unknown>): Partial<IMcpTargetContext> {
  const derived: Partial<IMcpTargetContext> = {};
  const rowData = tryParseJsonRecord(record.rowData);
  const parentReference = tryParseJsonRecord(record.parentReference);
  const listRecord = tryParseJsonRecord(record.list);
  const nestedSite = tryParseJsonRecord(record.site);
  const userRecord = tryParseJsonRecord(record.user);
  const channelRecord = tryParseJsonRecord(record.channel);
  const teamRecord = tryParseJsonRecord(record.team);
  Object.assign(derived, extractIdsFromSharePointODataContext(record['@odata.context']));

  const url = pickFirstNonEmptyString(
    record.siteUrl,
    record.site_url,
    record.targetSiteUrl,
    record.fileOrFolderUrl,
    record.documentLibraryUrl,
    record.listUrl,
    record.webUrl,
    record.url,
    rowData?.webUrl,
    rowData?.url,
    nestedSite?.webUrl,
    nestedSite?.url
  );

  if (url) {
    Object.assign(derived, deriveSharePointContextFromUrl(url));
  }

  applyStringField(derived, 'siteId', record.siteId, nestedSite?.id, parentReference?.siteId);
  applyStringField(derived, 'siteUrl', record.siteUrl, record.site_url, record.targetSiteUrl, nestedSite?.webUrl, nestedSite?.url, derived.siteUrl);
  applyStringField(derived, 'siteName', record.siteName, nestedSite?.displayName, nestedSite?.name, derived.siteName);

  applyStringField(derived, 'documentLibraryId', record.documentLibraryId, record.driveId, parentReference?.driveId);
  applyStringField(derived, 'documentLibraryUrl', record.documentLibraryUrl, record.libraryUrl, record.library_url, record.targetLibraryUrl, derived.documentLibraryUrl);
  applyStringField(
    derived,
    'documentLibraryName',
    record.documentLibraryName,
    record.libraryName,
    record.library_name,
    record.targetLibrary,
    record.driveName,
    derived.documentLibraryName
  );

  const isLikelyList = !!pickFirstNonEmptyString(record.listId, listRecord?.template, record.displayName)
    && /\/Lists\//i.test(pickFirstNonEmptyString(record.listUrl, record.webUrl, record.url) || '')
    && !pickFirstNonEmptyString(record.listItemId, record.itemId);
  const hasListContext = !!pickFirstNonEmptyString(
    record.listId,
    record.listUrl,
    record.listName,
    listRecord?.id,
    derived.listUrl,
    derived.listName
  );
  const isLikelyFile = !!pickFirstNonEmptyString(record.fileOrFolderId, record.fileId)
    || !!pickFirstNonEmptyString(record.fileOrFolderUrl, derived.fileOrFolderUrl)
    || !!pickFirstNonEmptyString(record.fileType, record.itemType)
    || (!!pickFirstNonEmptyString(record.itemId) && !hasListContext);
  const isLikelyListItem = hasListContext
    && !!pickFirstNonEmptyString(record.listItemId, record.itemId)
    && !isLikelyList
    && !isLikelyFile;
  const isLikelyMailItem = !!pickFirstNonEmptyString(record.mailItemId, record.messageId)
    || (
      !!pickFirstNonEmptyString(record.itemId)
      && !!pickFirstNonEmptyString(record.Subject, record.subject, record.From, record.from, record.sender, record.senderEmail)
    );

  if (isLikelyList) {
    applyStringField(derived, 'listId', record.listId, record.id);
  } else if (isLikelyFile) {
    applyStringField(derived, 'fileOrFolderId', record.fileOrFolderId, record.fileId, record.itemId, record.id);
  }

  const isLikelyColumnDefinition = !!pickFirstNonEmptyString(record.columnId, record.id)
    && (
      !!pickFirstNonEmptyString(record.columnGroup, record.columnDisplayName)
      || record.text !== undefined
      || record.number !== undefined
      || record.choice !== undefined
      || record.boolean !== undefined
      || record.dateTime !== undefined
      || record.personOrGroup !== undefined
      || record.lookup !== undefined
      || record.hyperlinkOrPicture !== undefined
    );

  applyStringField(derived, 'listId', derived.listId, listRecord?.id, record.listId);
  applyStringField(derived, 'listItemId', record.listItemId, isLikelyListItem ? record.itemId : undefined);
  applyStringField(derived, 'listUrl', record.listUrl, derived.listUrl, record.webUrl);
  applyStringField(
    derived,
    'listName',
    record.listName,
    isLikelyColumnDefinition ? undefined : record.displayName,
    isLikelyColumnDefinition ? undefined : record.name,
    derived.listName
  );

  if (isLikelyColumnDefinition) {
    applyStringField(derived, 'columnId', record.columnId, record.id);
    applyStringField(derived, 'columnName', record.columnName, record.name);
    applyStringField(derived, 'columnDisplayName', record.columnDisplayName, record.displayName, record.title);
  }

  applyStringField(derived, 'fileOrFolderId', derived.fileOrFolderId, record.fileOrFolderId, record.fileId);
  applyStringField(derived, 'fileOrFolderUrl', record.fileOrFolderUrl, derived.fileOrFolderUrl);
  applyStringField(derived, 'fileOrFolderName', record.fileOrFolderName, record.fileName, record.title, record.name, derived.fileOrFolderName);

  const personEmail = pickFirstNonEmptyString(
    record.personEmail,
    record.email,
    record.mail,
    record.userPrincipalName,
    userRecord?.email,
    userRecord?.mail,
    userRecord?.userPrincipalName
  );
  const personDisplayName = pickFirstNonEmptyString(
    record.personDisplayName,
    record.displayName,
    record.name,
    userRecord?.displayName
  );
  if (personEmail || personDisplayName) {
    derived.personEmail = personEmail;
    derived.personDisplayName = personDisplayName;
    derived.personIdentifier = pickFirstNonEmptyString(record.personIdentifier, personEmail, personDisplayName);
  }

  applyStringField(derived, 'mailItemId', record.mailItemId, record.messageId, isLikelyMailItem ? record.itemId : undefined);
  applyStringField(derived, 'calendarItemId', record.calendarItemId, record.eventId);
  applyStringField(derived, 'teamId', record.teamId, teamRecord?.id);
  applyStringField(derived, 'teamName', record.teamName, teamRecord?.displayName, teamRecord?.name);
  applyStringField(derived, 'channelId', record.channelId, channelRecord?.id);
  applyStringField(derived, 'channelName', record.channelName, channelRecord?.displayName, channelRecord?.name);

  return rowData ? mergeMcpTargetContexts(deriveTargetContextFromRecord(rowData), derived) || {} : derived;
}

function hasContextFields(context: Partial<IMcpTargetContext> | undefined): context is IMcpTargetContext {
  if (!context) {
    return false;
  }
  return Object.keys(context).some((key) => key !== 'source' && !!context[key as keyof IMcpTargetContext]);
}

export function mergeMcpTargetContexts(
  ...contexts: Array<Partial<IMcpTargetContext> | undefined>
): IMcpTargetContext | undefined {
  const merged: Partial<IMcpTargetContext> = {};
  contexts.forEach((context) => {
    if (!context) {
      return;
    }
    Object.entries(context).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== '') {
        (merged as Record<string, unknown>)[key] = value;
      }
    });
  });
  return hasContextFields(merged) ? merged as IMcpTargetContext : undefined;
}

export function deriveMcpTargetContextFromUnknown(
  value: unknown,
  source?: McpTargetSource
): IMcpTargetContext | undefined {
  if (!value) {
    return undefined;
  }
  if (typeof value === 'string') {
    const directUrlContext = deriveSharePointContextFromUrl(value);
    if (hasContextFields(directUrlContext)) {
      return { ...directUrlContext, source };
    }
    return undefined;
  }
  if (Array.isArray(value)) {
    const merged = value.reduce<IMcpTargetContext | undefined>((acc, entry) => {
      return mergeMcpTargetContexts(acc, deriveMcpTargetContextFromUnknown(entry, source));
    }, undefined);
    return merged ? { ...merged, source: merged.source || source } : undefined;
  }

  const record = tryParseJsonRecord(value);
  if (!record) {
    return undefined;
  }

  const direct = deriveTargetContextFromRecord(record);
  const selectedItems = Array.isArray(record.selectedItems)
    ? record.selectedItems.map((entry) => deriveMcpTargetContextFromUnknown(entry, source))
    : [];
  const nested = mergeMcpTargetContexts(
    deriveMcpTargetContextFromUnknown(record.targetContext, source),
    deriveMcpTargetContextFromUnknown(record.selectedItem, source),
    deriveMcpTargetContextFromUnknown(record.rowData, source),
    deriveMcpTargetContextFromUnknown(record.file, source),
    deriveMcpTargetContextFromUnknown(record.site, source),
    deriveMcpTargetContextFromUnknown(record.person, source),
    ...selectedItems
  );

  const merged = mergeMcpTargetContexts(direct, nested);
  return merged ? { ...merged, source: merged.source || source } : undefined;
}

export function deriveMcpTargetContextFromHie(
  sourceContext?: IHieSourceContext,
  taskContext?: IHieTaskContext,
  artifacts?: Readonly<Record<string, IHieArtifactRecord>>
): IMcpTargetContext | undefined {
  const artifactContext = taskContext && artifacts
    ? resolveCurrentArtifactContext(taskContext, artifacts)
    : undefined;
  const latestArtifactContext = artifacts
    ? resolveLatestArtifactContext(artifacts)
    : undefined;

  const preferredContext = mergeMcpTargetContexts(
    deriveMcpTargetContextFromUnknown(sourceContext?.targetContext, 'hie-selection'),
    deriveMcpTargetContextFromUnknown(sourceContext?.selectedItems, 'hie-selection'),
    deriveMcpTargetContextFromUnknown(taskContext?.targetContext, 'hie-selection'),
    deriveMcpTargetContextFromUnknown(taskContext?.selectedItems, 'hie-selection'),
    deriveMcpTargetContextFromUnknown(artifactContext?.currentArtifact?.targetContext, 'hie-selection'),
    deriveMcpTargetContextFromUnknown(artifactContext?.primaryArtifact?.targetContext, 'hie-selection')
  );

  if (preferredContext) {
    return preferredContext;
  }

  return mergeMcpTargetContexts(
    deriveMcpTargetContextFromUnknown(latestArtifactContext?.currentArtifact?.targetContext, 'hie-selection'),
    deriveMcpTargetContextFromUnknown(latestArtifactContext?.primaryArtifact?.targetContext, 'hie-selection')
  );
}

export function describeMcpTargetContext(context?: IMcpTargetContext): string | undefined {
  if (!context) {
    return undefined;
  }
  return pickFirstNonEmptyString(
    context.fileOrFolderName,
    context.listName,
    context.documentLibraryName,
    context.personDisplayName,
    context.channelName && context.teamName ? `${context.teamName}/${context.channelName}` : undefined,
    context.teamName,
    context.siteName,
    context.fileOrFolderUrl,
    context.listUrl,
    context.documentLibraryUrl,
    context.siteUrl,
    context.personEmail
  );
}
