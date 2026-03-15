/**
 * SitesService
 * SharePoint sites, drives, and items operations via Microsoft Graph.
 */

import { AadHttpClient } from '@microsoft/sp-http';
import { GraphService } from './GraphService';
import type { IGraphResponse } from './GraphService';

// ─── Graph API response shapes ──────────────────────────────────

interface IGraphSite {
  id: string;
  displayName: string;
  webUrl: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

interface IGraphDrive {
  id: string;
  name: string;
  webUrl: string;
  driveType: string;
  quota?: { used: number; total: number };
}

interface IGraphDriveItem {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { displayName?: string } };
  file?: { mimeType: string };
  folder?: { childCount: number };
  parentReference?: { driveId?: string };
}

interface IGraphList {
  id: string;
  displayName: string;
  webUrl: string;
  list?: { template: string };
}

interface IGraphPermission {
  id: string;
  roles: string[];
  grantedToV2?: {
    user?: { displayName?: string; email?: string };
    group?: { displayName?: string };
    siteUser?: { displayName?: string; loginName?: string };
  };
  grantedTo?: {
    user?: { displayName?: string; email?: string };
  };
  inheritedFrom?: { id: string };
}

interface IGraphListItem {
  id: string;
  fields: Record<string, string>;
}

interface IGraphCollection<T> {
  value: T[];
  '@odata.count'?: number;
}

// ─── Service ───────────────────────────────────────────────────

export interface ISiteInfo {
  siteName: string;
  siteUrl: string;
  description?: string;
  created?: string;
  lastModified?: string;
  libraries: string[];
  lists: string[];
}

export interface IDriveItemInfo {
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

export interface IFileDetails {
  fileName: string;
  fileUrl: string;
  fileType: string;
  size?: number;
  lastModified?: string;
  author?: string;
  metadata: Record<string, string>;
}

export interface IFileIds {
  itemId: string;
  driveId: string;
  fileName: string;
}

export interface IPermissionResult {
  principal: string;
  role: string;
  inherited: boolean;
}

export class SitesService {
  private graphService: GraphService;

  constructor(client: AadHttpClient) {
    this.graphService = new GraphService(client);
  }

  /**
   * Encode each segment but keep path separators for Graph root:/{path} syntax.
   */
  private encodeGraphPath(path: string): string {
    const normalized = path.replace(/^\/+|\/+$/g, '');
    if (!normalized) return '';
    return normalized
      .split('/')
      .map((segment) => encodeURIComponent(segment))
      .join('/');
  }

  /**
   * Search for SharePoint sites.
   */
  public async searchSites(query: string): Promise<IGraphResponse<ISiteInfo[]>> {
    const resp = await this.graphService.get<IGraphCollection<IGraphSite>>(
      `/sites?search=${encodeURIComponent(query)}&$top=10&$select=id,displayName,webUrl,description,createdDateTime,lastModifiedDateTime`
    );

    if (!resp.success || !resp.data) {
      return { success: false, error: resp.error, durationMs: resp.durationMs };
    }

    const sites: ISiteInfo[] = resp.data.value.map((s) => ({
      siteName: s.displayName,
      siteUrl: s.webUrl,
      description: s.description,
      created: s.createdDateTime,
      lastModified: s.lastModifiedDateTime,
      libraries: [],
      lists: []
    }));

    return { success: true, data: sites, durationMs: resp.durationMs };
  }

  /**
   * Get site info by URL (extracts hostname and path from full URL).
   */
  public async getSiteInfo(siteUrl: string): Promise<IGraphResponse<ISiteInfo>> {
    // Parse site URL → Graph path: /sites/{hostname}:{sitePath}
    let graphPath: string;
    try {
      const url = new URL(siteUrl);
      const hostname = url.hostname;
      const sitePath = url.pathname.replace(/\/$/, '');
      graphPath = `/sites/${hostname}:${sitePath}`;
    } catch {
      return { success: false, error: `Invalid site URL: ${siteUrl}` };
    }

    // Get site + drives + lists in sequence (Graph doesn't support $batch via AadHttpClient easily)
    const siteResp = await this.graphService.get<IGraphSite>(
      `${graphPath}?$select=id,displayName,webUrl,description,createdDateTime,lastModifiedDateTime`
    );

    if (!siteResp.success || !siteResp.data) {
      return { success: false, error: siteResp.error, durationMs: siteResp.durationMs };
    }

    const site = siteResp.data;
    const siteId = site.id;

    // Fetch drives (document libraries)
    const drivesResp = await this.graphService.get<IGraphCollection<IGraphDrive>>(
      `/sites/${siteId}/drives?$select=id,name,webUrl,driveType`
    );
    const libraries = drivesResp.success && drivesResp.data
      ? drivesResp.data.value.map((d) => d.name)
      : [];

    // Fetch lists
    const listsResp = await this.graphService.get<IGraphCollection<IGraphList>>(
      `/sites/${siteId}/lists?$select=id,displayName,webUrl&$top=20`
    );
    const lists = listsResp.success && listsResp.data
      ? listsResp.data.value.map((l) => l.displayName)
      : [];

    return {
      success: true,
      data: {
        siteName: site.displayName,
        siteUrl: site.webUrl,
        description: site.description,
        created: site.createdDateTime,
        lastModified: site.lastModifiedDateTime,
        libraries,
        lists
      },
      durationMs: siteResp.durationMs
    };
  }

  /**
   * Browse a document library (default drive) for a site.
   */
  public async browseDrive(
    siteUrl: string,
    folderPath?: string
  ): Promise<IGraphResponse<IDriveItemInfo[]>> {
    // Get site ID first
    let graphSitePath: string;
    try {
      const url = new URL(siteUrl);
      graphSitePath = `/sites/${url.hostname}:${url.pathname.replace(/\/$/, '')}`;
    } catch {
      return { success: false, error: `Invalid site URL: ${siteUrl}` };
    }

    const siteResp = await this.graphService.get<IGraphSite>(`${graphSitePath}?$select=id`);
    if (!siteResp.success || !siteResp.data) {
      return { success: false, error: siteResp.error };
    }

    const siteId = siteResp.data.id;
    const encodedFolderPath = folderPath ? this.encodeGraphPath(folderPath) : '';
    const itemsPath = folderPath
      ? `/sites/${siteId}/drive/root:/${encodedFolderPath}:/children`
      : `/sites/${siteId}/drive/root/children`;

    const resp = await this.graphService.get<IGraphCollection<IGraphDriveItem>>(
      `${itemsPath}?$select=id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy,file,folder,parentReference&$top=50`
    );

    if (!resp.success || !resp.data) {
      return { success: false, error: resp.error, durationMs: resp.durationMs };
    }

    const items: IDriveItemInfo[] = resp.data.value.map((item) => ({
      name: item.name,
      type: item.folder ? 'folder' as const : 'file' as const,
      url: item.webUrl,
      documentLibraryId: item.parentReference?.driveId,
      fileOrFolderId: item.id,
      size: item.size,
      lastModified: item.lastModifiedDateTime,
      author: item.lastModifiedBy?.user?.displayName,
      fileType: item.file ? item.name.split('.').pop() : undefined
    }));

    return { success: true, data: items, durationMs: resp.durationMs };
  }

  /**
   * Get list items from a named list on a site.
   */
  public async getListItems(
    siteUrl: string,
    listName: string,
    top: number = 25
  ): Promise<IGraphResponse<{ columns: string[]; items: Record<string, string>[]; totalCount: number }>> {
    let graphSitePath: string;
    try {
      const url = new URL(siteUrl);
      graphSitePath = `/sites/${url.hostname}:${url.pathname.replace(/\/$/, '')}`;
    } catch {
      return { success: false, error: `Invalid site URL: ${siteUrl}` };
    }

    const siteResp = await this.graphService.get<IGraphSite>(`${graphSitePath}?$select=id`);
    if (!siteResp.success || !siteResp.data) {
      return { success: false, error: siteResp.error };
    }

    const siteId = siteResp.data.id;
    const encodedListName = encodeURIComponent(listName);
    const listResp = await this.graphService.get<IGraphList>(
      `/sites/${siteId}/lists/${encodedListName}?$select=id,displayName,webUrl`
    );
    if (!listResp.success || !listResp.data) {
      return { success: false, error: listResp.error, durationMs: listResp.durationMs };
    }
    const listMetadata = listResp.data;

    const resp = await this.graphService.get<IGraphCollection<IGraphListItem>>(
      `/sites/${siteId}/lists/${encodedListName}/items?expand=fields&$top=${top}`
    );

    if (!resp.success || !resp.data) {
      return { success: false, error: resp.error, durationMs: resp.durationMs };
    }

    // Extract columns from the first item's fields
    const rawItems = resp.data.value;
    const columnSet = new Set<string>();
    rawItems.forEach((item) => {
      Object.keys(item.fields).forEach((key) => {
        // Skip internal Graph metadata fields
        if (!key.startsWith('@odata') && !key.startsWith('_')) {
          columnSet.add(key);
        }
      });
    });

    const columns: string[] = [];
    columnSet.forEach((c) => columns.push(c));

    const items = rawItems.map((item) => {
      const row: Record<string, string> = {};
      columns.forEach((col) => {
        row[col] = item.fields[col] !== undefined ? String(item.fields[col]) : '';
      });
      row.itemId = item.id;
      row.siteId = siteId;
      row.listId = listMetadata.id;
      row.listName = listMetadata.displayName || listName;
      row.listUrl = listMetadata.webUrl || '';
      return row;
    });

    return {
      success: true,
      data: { columns, items, totalCount: resp.data['@odata.count'] || items.length },
      durationMs: resp.durationMs
    };
  }

  /**
   * Get permissions for a site (by URL).
   */
  public async getPermissions(siteUrl: string): Promise<IGraphResponse<IPermissionResult[]>> {
    let graphSitePath: string;
    try {
      const url = new URL(siteUrl);
      graphSitePath = `/sites/${url.hostname}:${url.pathname.replace(/\/$/, '')}`;
    } catch {
      return { success: false, error: `Invalid site URL: ${siteUrl}` };
    }

    const siteResp = await this.graphService.get<IGraphSite>(`${graphSitePath}?$select=id`);
    if (!siteResp.success || !siteResp.data) {
      return { success: false, error: siteResp.error };
    }

    const siteId = siteResp.data.id;
    const resp = await this.graphService.get<IGraphCollection<IGraphPermission>>(
      `/sites/${siteId}/permissions`
    );

    if (!resp.success || !resp.data) {
      return { success: false, error: resp.error, durationMs: resp.durationMs };
    }

    const permissions: IPermissionResult[] = resp.data.value.map((p) => {
      const principal =
        p.grantedToV2?.user?.displayName ||
        p.grantedToV2?.group?.displayName ||
        p.grantedToV2?.siteUser?.displayName ||
        p.grantedTo?.user?.displayName ||
        'Unknown';
      const role = p.roles.join(', ');
      const inherited = !!p.inheritedFrom;
      return { principal, role, inherited };
    });

    return { success: true, data: permissions, durationMs: resp.durationMs };
  }

  /**
   * Parse a SharePoint file URL into its constituent parts.
   * Handles /sites/X, /personal/X (OneDrive), and root-site URLs.
   */
  private parseSharePointFileUrl(fileUrl: string): {
    hostname: string;
    sitePath: string;
    librarySegment: string;
    filePath: string;
  } | undefined {
    let url: URL;
    try {
      url = new URL(fileUrl);
    } catch {
      return undefined;
    }

    const hostname = url.hostname;
    // Decode percent-encoded path segments
    const fullPath = decodeURIComponent(url.pathname);
    // Remove trailing slash
    const cleanPath = fullPath.replace(/\/$/, '');
    const segments = cleanPath.split('/').filter(Boolean);

    if (segments.length < 2) return undefined;

    let sitePath = '';
    let remainingStart = 0;

    // /sites/X or /teams/X → site path is first 2 segments
    if ((segments[0] === 'sites' || segments[0] === 'teams') && segments.length >= 2) {
      sitePath = `/${segments[0]}/${segments[1]}`;
      remainingStart = 2;
    }
    // /personal/user_tenant_com → OneDrive personal site
    else if (segments[0] === 'personal' && segments.length >= 2) {
      sitePath = `/${segments[0]}/${segments[1]}`;
      remainingStart = 2;
    }
    // Root site: first segment is the library name directly
    else {
      sitePath = '';
      remainingStart = 0;
    }

    const remaining = segments.slice(remainingStart);

    // Handle _layouts/15/Doc.aspx?file=... URLs (Office Online viewer URLs)
    // These have the real filename in the `file` query parameter
    if (remaining.length >= 1 && remaining[0] === '_layouts') {
      const fileParam = url.searchParams.get('file');
      if (fileParam) {
        // Use default library segment 'Shared Documents' — getFileDetails will
        // fall back to searching all drives if the default drive doesn't match
        return { hostname, sitePath, librarySegment: 'Shared Documents', filePath: fileParam };
      }
      return undefined;
    }

    if (remaining.length < 2) return undefined; // Need at least library + file

    const librarySegment = remaining[0];
    const filePath = remaining.slice(1).join('/');

    return { hostname, sitePath, librarySegment, filePath };
  }

  /**
   * Build IFileDetails from a Graph drive item response.
   */
  private buildFileDetails(item: IGraphDriveItem, fileUrl: string): IFileDetails {
    const fileName = item.name;
    const fileType = fileName.includes('.') ? fileName.split('.').pop() || '' : '';
    const metadata: Record<string, string> = {};
    if (item.size !== undefined) metadata.Size = `${Math.round(item.size / 1024)} KB`;
    if (item.lastModifiedDateTime) metadata['Last Modified'] = item.lastModifiedDateTime;
    if (item.lastModifiedBy?.user?.displayName) metadata['Modified By'] = item.lastModifiedBy.user.displayName;
    if (item.file?.mimeType) metadata.Type = item.file.mimeType;

    return {
      fileName,
      fileUrl: item.webUrl || fileUrl,
      fileType,
      size: item.size,
      lastModified: item.lastModifiedDateTime,
      author: item.lastModifiedBy?.user?.displayName,
      metadata
    };
  }

  /**
   * Get file metadata by URL using path-based Graph access.
   * Parses the SharePoint URL to extract site path and file path,
   * then queries Graph via /sites/{hostname}:{sitePath}:/drive/root:/{filePath}.
   * Falls back to listing drives if the file is not in the default drive.
   */
  public async getFileDetails(fileUrl: string): Promise<IGraphResponse<IFileDetails>> {
    // Validate that fileUrl is actually a URL — the LLM sometimes passes a file name instead
    if (!fileUrl.startsWith('http://') && !fileUrl.startsWith('https://')) {
      return { success: false, error: `Invalid file URL (got "${fileUrl.substring(0, 60)}" — expected https://...)`, durationMs: 0 };
    }

    const parsed = this.parseSharePointFileUrl(fileUrl);
    if (!parsed) {
      // Can't parse URL — return basic details from the URL itself
      const urlFileName = decodeURIComponent(fileUrl.split('/').pop() || 'Unknown');
      const ext = urlFileName.includes('.') ? urlFileName.split('.').pop() || '' : '';
      return {
        success: true,
        data: {
          fileName: urlFileName,
          fileUrl,
          fileType: ext,
          metadata: {}
        },
        durationMs: 0
      };
    }

    const { hostname, sitePath, librarySegment, filePath } = parsed;
    const siteGraphRef = sitePath
      ? `/sites/${hostname}:${sitePath}:`
      : `/sites/${hostname}:/:`;
    const selectFields = '$select=id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy,file';

    // Step 1: Resolve site ID (avoids nested colon-paths which Graph rejects with 400)
    const siteIdResp = await this.graphService.get<IGraphSite>(
      `${siteGraphRef}?$select=id`
    );
    if (!siteIdResp.success || !siteIdResp.data) {
      // Graph can't find the site — return basic details from URL
      const urlFileName = decodeURIComponent(fileUrl.split('/').pop() || 'Unknown');
      const ext = urlFileName.includes('.') ? urlFileName.split('.').pop() || '' : '';
      return {
        success: true,
        data: { fileName: urlFileName, fileUrl, fileType: ext, metadata: {} },
        durationMs: siteIdResp.durationMs
      };
    }

    const siteId = siteIdResp.data.id;
    const encodedFilePath = this.encodeGraphPath(filePath);

    // Step 2: Try default drive first (most common — library is the site's default doc library)
    const defaultDrivePath = `/sites/${siteId}/drive/root:/${encodedFilePath}`;
    const resp = await this.graphService.get<IGraphDriveItem>(
      `${defaultDrivePath}?${selectFields}`
    );

    if (resp.success && resp.data) {
      return {
        success: true,
        data: this.buildFileDetails(resp.data, fileUrl),
        durationMs: (siteIdResp.durationMs || 0) + (resp.durationMs || 0)
      };
    }

    // Step 3: File is in a non-default drive — list drives and match by library name or webUrl
    const drivesResp = await this.graphService.get<IGraphCollection<IGraphDrive>>(
      `/sites/${siteId}/drives?$select=id,name,webUrl`
    );

    if (drivesResp.success && drivesResp.data) {
      const lowerLib = librarySegment.toLowerCase();
      // Match by internal name first, then by localized webUrl path (e.g. "Freigegebene Dokumente")
      const matchingDrive = drivesResp.data.value.find((d) =>
        d.name.toLowerCase() === lowerLib
      ) || drivesResp.data.value.find((d) => {
        try {
          const driveUrl = new URL(d.webUrl);
          const lastSegment = decodeURIComponent(driveUrl.pathname.split('/').filter(Boolean).pop() || '');
          return lastSegment.toLowerCase() === lowerLib;
        } catch { return false; }
      });
      if (matchingDrive) {
        const drivePath = `/drives/${matchingDrive.id}/root:/${encodedFilePath}`;
        const driveResp = await this.graphService.get<IGraphDriveItem>(
          `${drivePath}?${selectFields}`
        );
        if (driveResp.success && driveResp.data) {
          return {
            success: true,
            data: this.buildFileDetails(driveResp.data, fileUrl),
            durationMs: (siteIdResp.durationMs || 0) + (driveResp.durationMs || 0)
          };
        }
      }
    }

    // All Graph calls failed — graceful degradation: return basic info from URL
    const urlFileName = decodeURIComponent(fileUrl.split('/').pop() || 'Unknown');
    const ext = urlFileName.includes('.') ? urlFileName.split('.').pop() || '' : '';
    return {
      success: true,
      data: { fileName: urlFileName, fileUrl, fileType: ext, metadata: {} },
      durationMs: resp.durationMs
    };
  }

  /**
   * Resolve a SharePoint file URL to driveItem IDs needed by MCP readSmallTextFile.
   * Uses the same URL parsing + Graph resolution logic as getFileDetails().
   */
  public async resolveFileIds(fileUrl: string): Promise<IGraphResponse<IFileIds>> {
    if (!fileUrl.startsWith('http://') && !fileUrl.startsWith('https://')) {
      return { success: false, error: `Invalid file URL (got "${fileUrl.substring(0, 60)}" — expected https://...)`, durationMs: 0 };
    }

    const parsed = this.parseSharePointFileUrl(fileUrl);
    if (!parsed) {
      return { success: false, error: `Could not parse SharePoint file URL: ${fileUrl}` };
    }

    const { hostname, sitePath, librarySegment, filePath } = parsed;
    const siteGraphRef = sitePath
      ? `/sites/${hostname}:${sitePath}:`
      : `/sites/${hostname}:/:`;
    const selectFields = '$select=id,name,parentReference';

    // Step 1: Resolve site ID
    const siteIdResp = await this.graphService.get<IGraphSite>(
      `${siteGraphRef}?$select=id`
    );
    if (!siteIdResp.success || !siteIdResp.data) {
      return { success: false, error: siteIdResp.error || 'Could not resolve site', durationMs: siteIdResp.durationMs };
    }

    const siteId = siteIdResp.data.id;
    const encodedFilePath = this.encodeGraphPath(filePath);

    // Step 2: Try default drive first
    const defaultDrivePath = `/sites/${siteId}/drive/root:/${encodedFilePath}`;
    const resp = await this.graphService.get<IGraphDriveItem>(
      `${defaultDrivePath}?${selectFields}`
    );

    if (resp.success && resp.data && resp.data.parentReference?.driveId) {
      return {
        success: true,
        data: { itemId: resp.data.id, driveId: resp.data.parentReference.driveId, fileName: resp.data.name },
        durationMs: (siteIdResp.durationMs || 0) + (resp.durationMs || 0)
      };
    }

    // Step 3: Non-default drive — list drives and match by library name
    const drivesResp = await this.graphService.get<IGraphCollection<IGraphDrive>>(
      `/sites/${siteId}/drives?$select=id,name,webUrl`
    );

    if (drivesResp.success && drivesResp.data) {
      const lowerLib = librarySegment.toLowerCase();
      const matchingDrive = drivesResp.data.value.find((d) =>
        d.name.toLowerCase() === lowerLib
      ) || drivesResp.data.value.find((d) => {
        try {
          const driveUrl = new URL(d.webUrl);
          const lastSegment = decodeURIComponent(driveUrl.pathname.split('/').filter(Boolean).pop() || '');
          return lastSegment.toLowerCase() === lowerLib;
        } catch { return false; }
      });
      if (matchingDrive) {
        const drivePath = `/drives/${matchingDrive.id}/root:/${encodedFilePath}`;
        const driveResp = await this.graphService.get<IGraphDriveItem>(
          `${drivePath}?${selectFields}`
        );
        if (driveResp.success && driveResp.data) {
          return {
            success: true,
            data: { itemId: driveResp.data.id, driveId: matchingDrive.id, fileName: driveResp.data.name },
            durationMs: (siteIdResp.durationMs || 0) + (driveResp.durationMs || 0)
          };
        }
      }
    }

    return { success: false, error: `Could not resolve file IDs for: ${fileUrl}`, durationMs: resp.durationMs };
  }

  /**
   * Get recent activity for a site.
   */
  public async getActivities(
    siteUrl: string,
    top: number = 20
  ): Promise<IGraphResponse<Array<{ action: string; actor: string; target: string; timestamp: string }>>> {
    let graphSitePath: string;
    try {
      const url = new URL(siteUrl);
      graphSitePath = `/sites/${url.hostname}:${url.pathname.replace(/\/$/, '')}`;
    } catch {
      return { success: false, error: `Invalid site URL: ${siteUrl}` };
    }

    const siteResp = await this.graphService.get<IGraphSite>(`${graphSitePath}?$select=id`);
    if (!siteResp.success || !siteResp.data) {
      return { success: false, error: siteResp.error };
    }

    const siteId = siteResp.data.id;

    // Use the drive's recent endpoint to get activity
    interface IGraphActivity {
      id: string;
      name: string;
      webUrl: string;
      lastModifiedDateTime: string;
      lastModifiedBy?: { user?: { displayName?: string } };
    }
    const resp = await this.graphService.get<IGraphCollection<IGraphActivity>>(
      `/sites/${siteId}/drive/recent?$top=${top}&$select=id,name,webUrl,lastModifiedDateTime,lastModifiedBy`
    );

    if (!resp.success || !resp.data) {
      return { success: false, error: resp.error, durationMs: resp.durationMs };
    }

    const activities = resp.data.value.map((item) => ({
      action: 'Modified',
      actor: item.lastModifiedBy?.user?.displayName || 'Unknown',
      target: item.name,
      timestamp: item.lastModifiedDateTime
    }));

    return { success: true, data: activities, durationMs: resp.durationMs };
  }
}
