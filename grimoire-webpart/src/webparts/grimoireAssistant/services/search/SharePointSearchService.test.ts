jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('../pnp/pnpConfig', () => ({
  getSP: jest.fn()
}));

import { getSP } from '../pnp/pnpConfig';
import { SharePointSearchService } from './SharePointSearchService';

describe('SharePointSearchService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('issues a classic SharePoint search query with language-aware culture settings', async () => {
    const searchMock = jest.fn().mockResolvedValue({
      PrimarySearchResults: []
    });
    (getSP as jest.Mock).mockReturnValue({ search: searchMock });

    const service = new SharePointSearchService();
    await service.search('rapport budget', { maxResults: 12, queryLanguage: 'fr' });

    const request = searchMock.mock.calls[0][0] as {
      Querytext: string;
      Culture: number;
      EnableQueryRules: boolean;
      TrimDuplicates: boolean;
    };

    expect(request.Querytext).toBe('rapport budget');
    expect(request.Culture).toBe(1036);
    expect(request.EnableQueryRules).toBe(false);
    expect(request.TrimDuplicates).toBe(true);
  });

  it('uses the planner-provided lexical query text when supplied', async () => {
    const searchMock = jest.fn().mockResolvedValue({
      PrimarySearchResults: []
    });
    (getSP as jest.Mock).mockReturnValue({ search: searchMock });

    const service = new SharePointSearchService();
    await service.search('docs about animals', {
      queryLanguage: 'en',
      requestQueryText: 'animals'
    });

    const request = searchMock.mock.calls[0][0] as {
      Querytext: string;
    };

    expect(request.Querytext).toBe('animals');
  });

  it('normalizes SharePoint results into the shared search result model', async () => {
    const searchMock = jest.fn().mockResolvedValue({
      PrimarySearchResults: [
        {
          Title: 'Budget FY26',
          Path: 'https://tenant.sharepoint.com/sites/finance/Shared%20Documents/Budget%20FY26.docx',
          HitHighlightedSummary: 'Preview <c0>budget</c0> summary',
          Author: 'Test User',
          LastModifiedTime: '2026-03-01T09:00:00.000Z',
          FileExtension: 'docx',
          SiteName: 'Finance',
          Rank: 42,
          Culture: '1033'
        }
      ]
    });
    (getSP as jest.Mock).mockReturnValue({ search: searchMock });

    const service = new SharePointSearchService();
    const response = await service.search('budget', { queryLanguage: 'en' });

    expect(response.success).toBe(true);
    expect(response.data).toEqual([
      expect.objectContaining({
        title: 'Budget FY26',
        summary: 'Preview budget summary',
        url: 'https://tenant.sharepoint.com/sites/finance/Shared%20Documents/Budget%20FY26.docx',
        author: 'Test User',
        fileType: 'docx',
        siteName: 'Finance',
        source: 'sharepoint-search',
        sourceNativeScore: 42,
        language: 'en'
      })
    ]);
  });

  it('filters strict single-token queries down to hits that actually contain the token', async () => {
    const searchMock = jest.fn().mockResolvedValue({
      PrimarySearchResults: [
        {
          Title: 'SPFx Guide',
          Path: 'https://tenant.sharepoint.com/sites/dev/Shared%20Documents/SPFx%20Guide.pdf',
          HitHighlightedSummary: 'Overview of SPFx development',
          FileExtension: 'pdf',
          Rank: 10,
          Culture: '1033'
        },
        {
          Title: 'Animal Stories and Facts',
          Path: 'https://tenant.sharepoint.com/sites/dev/Shared%20Documents/Animal%20Stories.docx',
          HitHighlightedSummary: 'Animals and habitats',
          FileExtension: 'docx',
          Rank: 9,
          Culture: '1033'
        }
      ]
    });
    (getSP as jest.Mock).mockReturnValue({ search: searchMock });

    const service = new SharePointSearchService();
    const response = await service.search('spfx', {
      queryLanguage: 'en',
      requestQueryText: 'spfx'
    });

    expect(response.success).toBe(true);
    expect(response.data).toHaveLength(1);
    expect(response.data?.[0].title).toBe('SPFx Guide');
  });

  it('does not treat the SharePoint hostname as a lexical keyword match', async () => {
    const searchMock = jest.fn().mockResolvedValue({
      PrimarySearchResults: [
        {
          Title: 'SPFx Guide',
          Path: 'https://tenant.sharepoint.com/sites/dev/Shared%20Documents/SPFx%20Guide.pdf',
          HitHighlightedSummary: 'Overview of SPFx development',
          FileExtension: 'pdf',
          Rank: 10,
          Culture: '1033'
        },
        {
          Title: 'Animal Stories and Facts',
          Path: 'https://tenant.sharepoint.com/sites/dev/Shared%20Documents/Animal%20Stories.docx',
          HitHighlightedSummary: 'Animals and habitats',
          FileExtension: 'docx',
          Rank: 9,
          Culture: '1033'
        }
      ]
    });
    (getSP as jest.Mock).mockReturnValue({ search: searchMock });

    const service = new SharePointSearchService();
    const response = await service.search('spfx', {
      queryLanguage: 'en',
      variantKind: 'keyword-fallback',
      requestQueryText: 'spfx, SharePoint Framework, SPFx, development framework, SharePoint customization'
    });

    expect(response.success).toBe(true);
    expect(response.data).toHaveLength(1);
    expect(response.data?.[0].title).toBe('SPFx Guide');
  });
});
