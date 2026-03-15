jest.mock('@microsoft/sp-http', () => ({
  AadHttpClient: {
    configurations: {
      v1: {}
    }
  }
}));
jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

import { GraphService } from './GraphService';
import { SitesService } from './SitesService';

describe('SitesService', () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('browses a site drive and maps items to IDriveItemInfo', async () => {
    const getSpy = jest.spyOn(GraphService.prototype, 'get')
      .mockResolvedValueOnce({
        success: true,
        data: { id: 'site-id-123' }
      })
      .mockResolvedValueOnce({
        success: true,
        data: {
          value: [
            {
              id: 'item-1',
              name: 'Report.docx',
              webUrl: 'https://contoso.sharepoint.com/sites/dev/Shared%20Documents/Report.docx',
              file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
              folder: undefined,
              parentReference: { driveId: 'drive-abc' },
              size: 4096,
              lastModifiedDateTime: '2026-03-10T08:00:00.000Z',
              lastModifiedBy: { user: { displayName: 'Test User' } }
            },
            {
              id: 'item-2',
              name: 'Images',
              webUrl: 'https://contoso.sharepoint.com/sites/dev/Shared%20Documents/Images',
              file: undefined,
              folder: { childCount: 5 },
              parentReference: { driveId: 'drive-abc' },
              size: undefined,
              lastModifiedDateTime: '2026-03-09T14:00:00.000Z',
              lastModifiedBy: { user: { displayName: 'System Account' } }
            }
          ]
        }
      });

    const service = new SitesService({} as never);
    const response = await service.browseDrive('https://contoso.sharepoint.com/sites/dev');

    expect(getSpy).toHaveBeenCalledTimes(2);
    expect(getSpy).toHaveBeenNthCalledWith(
      1,
      '/sites/contoso.sharepoint.com:/sites/dev?$select=id'
    );
    expect(getSpy).toHaveBeenNthCalledWith(
      2,
      '/sites/site-id-123/drive/root/children?$select=id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy,file,folder,parentReference&$top=50'
    );
    expect(response).toMatchObject({
      success: true,
      data: [
        {
          name: 'Report.docx',
          type: 'file',
          documentLibraryId: 'drive-abc',
          fileOrFolderId: 'item-1',
          fileType: 'docx'
        },
        {
          name: 'Images',
          type: 'folder',
          documentLibraryId: 'drive-abc',
          fileOrFolderId: 'item-2',
          fileType: undefined
        }
      ]
    });
  });

  it('browses a subfolder when folderPath is provided', async () => {
    const getSpy = jest.spyOn(GraphService.prototype, 'get')
      .mockResolvedValueOnce({
        success: true,
        data: { id: 'site-id-456' }
      })
      .mockResolvedValueOnce({
        success: true,
        data: { value: [] }
      });

    const service = new SitesService({} as never);
    const response = await service.browseDrive(
      'https://contoso.sharepoint.com/sites/dev',
      'Shared Documents/Reports'
    );

    expect(getSpy).toHaveBeenCalledTimes(2);
    expect(getSpy).toHaveBeenNthCalledWith(
      2,
      '/sites/site-id-456/drive/root:/Shared%20Documents/Reports:/children?$select=id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy,file,folder,parentReference&$top=50'
    );
    expect(response).toMatchObject({ success: true, data: [] });
  });

  it('returns an error for an invalid site URL', async () => {
    const service = new SitesService({} as never);
    const response = await service.browseDrive('not-a-url');

    expect(response).toMatchObject({
      success: false,
      error: 'Invalid site URL: not-a-url'
    });
  });

  it('propagates Graph API errors from site resolution', async () => {
    jest.spyOn(GraphService.prototype, 'get').mockResolvedValueOnce({
      success: false,
      error: 'Site not found'
    });

    const service = new SitesService({} as never);
    const response = await service.browseDrive('https://contoso.sharepoint.com/sites/missing');

    expect(response).toMatchObject({
      success: false,
      error: 'Site not found'
    });
  });
});
