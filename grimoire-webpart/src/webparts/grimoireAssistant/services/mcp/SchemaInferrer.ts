/**
 * SchemaInferrer
 * Inspects parsed MCP tool result JSON and infers the best block type
 * + field mapping based on data shape, not tool name.
 *
 * Priority: catalog blockHint > shape detection > 'markdown' fallback.
 */

import type { BlockType } from '../../models/IBlock';

// ─── Public Types ───────────────────────────────────────────────

export interface ISchemaInference {
  blockType: BlockType;
  confidence: 'high' | 'medium' | 'low';
  /** Maps block data fields → source data field names */
  fieldMap: Record<string, string>;
}

// ─── Helpers ────────────────────────────────────────────────────

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

function hasAny(obj: Record<string, unknown>, ...keys: string[]): string | undefined {
  return keys.find((k) => k in obj && obj[k] !== undefined && obj[k] !== null);
}

/** Return the first element of an array (for inspecting shape of array items) */
function sampleItem(data: unknown[]): Record<string, unknown> | undefined {
  const item = data[0];
  return isRecord(item) ? item : undefined;
}

// ─── Action Result Detection ────────────────────────────────────

/**
 * Detect action/mutation results (create, delete, update, send, etc.)
 * These should show as info-card, not data blocks.
 */
function detectActionResult(parsed: unknown): ISchemaInference | undefined {
  if (!isRecord(parsed)) return undefined;
  const obj = parsed;

  // Pattern: { success: bool, message: "..." } or { status: "completed", ... }
  const hasSuccessFlag = 'success' in obj || 'status' in obj;
  const lacksDisplayData = !('webUrl' in obj) && !('displayName' in obj) && !('mail' in obj);
  const isSmallObject = Object.keys(obj).length <= 5;

  if (hasSuccessFlag && lacksDisplayData && isSmallObject) {
    return {
      blockType: 'info-card',
      confidence: 'medium',
      fieldMap: {
        heading: hasAny(obj, 'status', 'success') || 'status',
        body: hasAny(obj, 'message', 'description', 'detail', 'error') || 'message'
      }
    };
  }
  return undefined;
}

// ─── Single Object Detectors ────────────────────────────────────

function detectUserCard(obj: Record<string, unknown>): ISchemaInference | undefined {
  const nameField = hasAny(obj, 'displayName', 'name');
  const emailField = hasAny(obj, 'mail', 'email', 'userPrincipalName');
  const roleField = hasAny(obj, 'jobTitle', 'department');
  if (nameField && emailField && roleField) {
    return {
      blockType: 'user-card',
      confidence: 'high',
      fieldMap: {
        displayName: nameField,
        email: emailField,
        jobTitle: hasAny(obj, 'jobTitle') || '',
        department: hasAny(obj, 'department') || '',
        officeLocation: hasAny(obj, 'officeLocation') || '',
        phone: hasAny(obj, 'businessPhones', 'phone', 'mobilePhone') || ''
      }
    };
  }
  return undefined;
}

function detectSiteInfo(obj: Record<string, unknown>): ISchemaInference | undefined {
  const urlField = hasAny(obj, 'webUrl', 'url');
  const nameField = hasAny(obj, 'displayName', 'name');
  const metaField = hasAny(obj, 'description', 'createdDateTime', 'lastModifiedDateTime');
  // Must be a single object with site-like fields but NOT a file (no size/mimeType)
  if (urlField && nameField && metaField && !('size' in obj) && !('mimeType' in obj)) {
    return {
      blockType: 'site-info',
      confidence: 'high',
      fieldMap: {
        siteName: nameField,
        siteUrl: urlField,
        description: hasAny(obj, 'description') || '',
        created: hasAny(obj, 'createdDateTime') || '',
        lastModified: hasAny(obj, 'lastModifiedDateTime') || ''
      }
    };
  }
  return undefined;
}

function detectFilePreview(obj: Record<string, unknown>): ISchemaInference | undefined {
  const nameField = hasAny(obj, 'name', 'fileName');
  const urlField = hasAny(obj, 'webUrl', 'url');
  const fileSignal = hasAny(obj, 'size', 'mimeType', 'fileType', 'file');
  if (nameField && urlField && fileSignal) {
    return {
      blockType: 'file-preview',
      confidence: 'medium',
      fieldMap: {
        fileName: nameField,
        fileUrl: urlField,
        fileType: hasAny(obj, 'fileType', 'mimeType') || '',
        size: hasAny(obj, 'size') || '',
        lastModified: hasAny(obj, 'lastModifiedDateTime') || '',
        author: hasAny(obj, 'createdBy') || ''
      }
    };
  }
  return undefined;
}

// ─── Array Detectors ────────────────────────────────────────────

function detectDocumentLibrary(sample: Record<string, unknown>): ISchemaInference | undefined {
  // Items with name/webUrl and folder indicators or type=folder/file
  const nameField = hasAny(sample, 'name', 'displayName');
  const urlField = hasAny(sample, 'webUrl', 'url');
  const folderSignal = 'folder' in sample || 'type' in sample || 'file' in sample;
  if (nameField && urlField && folderSignal) {
    return {
      blockType: 'document-library',
      confidence: 'high',
      fieldMap: {
        name: nameField,
        url: urlField,
        size: hasAny(sample, 'size') || '',
        lastModified: hasAny(sample, 'lastModifiedDateTime') || '',
        author: hasAny(sample, 'createdBy') || '',
        fileType: hasAny(sample, 'fileType', 'mimeType') || ''
      }
    };
  }
  return undefined;
}

function detectSearchResults(sample: Record<string, unknown>): ISchemaInference | undefined {
  const urlField = hasAny(sample, 'webUrl', 'url');
  const titleField = hasAny(sample, 'name', 'title', 'displayName', 'subject');
  const summaryField = hasAny(sample, 'summary', 'bodyPreview', 'description');
  if (urlField && titleField && summaryField) {
    return {
      blockType: 'search-results',
      confidence: 'medium',
      fieldMap: {
        title: titleField,
        url: urlField,
        summary: summaryField,
        fileType: hasAny(sample, 'fileType') || '',
        lastModified: hasAny(sample, 'lastModifiedDateTime') || '',
        author: hasAny(sample, 'createdBy') || ''
      }
    };
  }
  return undefined;
}

function detectPermissionsView(sample: Record<string, unknown>): ISchemaInference | undefined {
  // Graph permissions: grantedTo/grantedToV2 + roles[]
  // Or simplified: principal + role
  const principalField = hasAny(sample, 'principal', 'grantedTo', 'grantedToV2');
  const roleField = hasAny(sample, 'role', 'roles');
  if (principalField && roleField) {
    return {
      blockType: 'permissions-view',
      confidence: 'medium',
      fieldMap: {
        principal: principalField,
        role: roleField,
        inherited: hasAny(sample, 'inherited', 'inheritedFrom') || ''
      }
    };
  }
  return undefined;
}

function detectActivityFeed(sample: Record<string, unknown>): ISchemaInference | undefined {
  const actorField = hasAny(sample, 'actor', 'user', 'initiatedBy');
  const actionField = hasAny(sample, 'action', 'activityType', 'activity');
  const targetField = hasAny(sample, 'target', 'resource', 'resourceReference');
  if (actorField && actionField) {
    return {
      blockType: 'activity-feed',
      confidence: 'medium',
      fieldMap: {
        actor: actorField,
        action: actionField,
        target: targetField || '',
        timestamp: hasAny(sample, 'timestamp', 'createdDateTime', 'activityDateTime') || ''
      }
    };
  }
  return undefined;
}

function detectSelectionList(sample: Record<string, unknown>): ISchemaInference | undefined {
  const idField = hasAny(sample, 'id');
  const labelField = hasAny(sample, 'displayName', 'topic', 'name', 'label');
  // Selection lists: have id+label but typically no webUrl
  if (idField && labelField && !hasAny(sample, 'webUrl', 'url')) {
    return {
      blockType: 'selection-list',
      confidence: 'low',
      fieldMap: {
        id: idField,
        label: labelField,
        description: hasAny(sample, 'description', 'chatType', 'membershipType') || ''
      }
    };
  }
  return undefined;
}

// ─── Main Inference ─────────────────────────────────────────────

/**
 * Infer the best block type and field mapping from parsed MCP tool result data.
 *
 * @param parsed - The parsed JSON data from the MCP tool result
 * @param catalogHint - Optional blockHint from the MCP catalog entry
 * @returns The inferred block type, confidence, and field mapping
 */
export function infer(parsed: unknown, catalogHint?: BlockType): ISchemaInference {
  // ── Catalog hint override ──
  // If the catalog explicitly specifies a block type, trust it and build
  // a best-effort field map from the data shape.
  if (catalogHint) {
    const fieldMap = buildFieldMapForHint(parsed, catalogHint);
    return { blockType: catalogHint, confidence: 'high', fieldMap };
  }

  // ── Action result detection (before data detection) ──
  const actionResult = detectActionResult(parsed);
  if (actionResult) return actionResult;

  // ── Single object detection ──
  if (isRecord(parsed)) {
    const userCard = detectUserCard(parsed);
    if (userCard) return userCard;

    const filePreview = detectFilePreview(parsed);
    if (filePreview) return filePreview;

    const siteInfo = detectSiteInfo(parsed);
    if (siteInfo) return siteInfo;

    // Single object that doesn't match specific types → markdown
    return markdownFallback();
  }

  // ── Array detection ──
  if (Array.isArray(parsed) && parsed.length > 0) {
    const sample = sampleItem(parsed);
    if (sample) {
      // Ordered by specificity (high → low)
      const docLib = detectDocumentLibrary(sample);
      if (docLib) return docLib;

      const permissions = detectPermissionsView(sample);
      if (permissions) return permissions;

      const activity = detectActivityFeed(sample);
      if (activity) return activity;

      const searchResults = detectSearchResults(sample);
      if (searchResults) return searchResults;

      const selectionList = detectSelectionList(sample);
      if (selectionList) return selectionList;

      // Array of objects with consistent keys → list-items (lowest specificity)
      return {
        blockType: 'list-items',
        confidence: 'low',
        fieldMap: {} // GenericBlockBuilder handles raw key extraction
      };
    }
  }

  // ── Fallback ──
  return markdownFallback();
}

function markdownFallback(): ISchemaInference {
  return {
    blockType: 'markdown',
    confidence: 'low',
    fieldMap: {}
  };
}

/**
 * Build a best-effort field map when a catalog hint forces a specific block type.
 * We still inspect the data to find the right source field names.
 */
function buildFieldMapForHint(parsed: unknown, hint: BlockType): Record<string, string> {
  if (Array.isArray(parsed) && parsed.length > 0) {
    const sample = sampleItem(parsed);
    if (!sample) return {};
    return buildFieldMapFromSample(sample, hint);
  }
  if (isRecord(parsed)) {
    return buildFieldMapFromSample(parsed, hint);
  }
  return {};
}

function buildFieldMapFromSample(sample: Record<string, unknown>, hint: BlockType): Record<string, string> {
  const map: Record<string, string> = {};

  switch (hint) {
    case 'user-card':
      map.displayName = hasAny(sample, 'displayName', 'name') || 'displayName';
      map.email = hasAny(sample, 'mail', 'email', 'userPrincipalName') || 'mail';
      map.jobTitle = hasAny(sample, 'jobTitle') || '';
      map.department = hasAny(sample, 'department') || '';
      break;
    case 'document-library':
      map.name = hasAny(sample, 'name', 'displayName') || 'name';
      map.url = hasAny(sample, 'webUrl', 'url') || 'webUrl';
      break;
    case 'search-results':
      map.title = hasAny(sample, 'name', 'title', 'displayName', 'subject') || 'name';
      map.url = hasAny(sample, 'webUrl', 'url') || 'webUrl';
      map.summary = hasAny(sample, 'summary', 'bodyPreview', 'description') || '';
      break;
    case 'file-preview':
      map.fileName = hasAny(sample, 'name', 'fileName') || 'name';
      map.fileUrl = hasAny(sample, 'webUrl', 'url') || 'webUrl';
      break;
    case 'site-info':
      map.siteName = hasAny(sample, 'displayName', 'name') || 'displayName';
      map.siteUrl = hasAny(sample, 'webUrl', 'url') || 'webUrl';
      break;
    case 'selection-list':
      map.id = hasAny(sample, 'id') || 'id';
      map.label = hasAny(sample, 'displayName', 'topic', 'name', 'label') || 'displayName';
      map.description = hasAny(sample, 'description', 'chatType') || '';
      break;
    case 'permissions-view':
      map.principal = hasAny(sample, 'principal', 'grantedTo') || 'principal';
      map.role = hasAny(sample, 'role', 'roles') || 'role';
      break;
    case 'activity-feed':
      map.actor = hasAny(sample, 'actor', 'user') || 'actor';
      map.action = hasAny(sample, 'action', 'activityType') || 'action';
      map.target = hasAny(sample, 'target', 'resource') || '';
      break;
    case 'list-items':
      // No specific mapping needed, GenericBlockBuilder handles raw keys
      break;
    default:
      break;
  }

  return map;
}
