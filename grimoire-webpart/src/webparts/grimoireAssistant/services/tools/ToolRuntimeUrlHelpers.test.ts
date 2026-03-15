import { createBlock } from '../../models/IBlock';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { extractSiteUrl, resolveUrlFromBlocks } from './ToolRuntimeUrlHelpers';

describe('ToolRuntimeUrlHelpers', () => {
  afterEach(() => {
    useGrimoireStore.setState({ blocks: [], freshBlockIds: [] });
  });

  it('resolves URL from search results blocks by file name', () => {
    const searchBlock = createBlock('search-results', 'Search', {
      kind: 'search-results',
      query: 'q',
      results: [{
        title: 'Quarterly-Report.docx',
        summary: 'Q1 report',
        url: 'https://contoso.sharepoint.com/sites/fin/Shared%20Documents/Quarterly-Report.docx',
        sources: ['search']
      }],
      totalCount: 1,
      source: 'search'
    });

    useGrimoireStore.setState({ blocks: [searchBlock] });

    const resolved = resolveUrlFromBlocks('https://contoso.sharepoint.com/sites/fin/placeholder.docx', 'Quarterly-Report.docx');
    expect(resolved).toBe('https://contoso.sharepoint.com/sites/fin/Shared%20Documents/Quarterly-Report.docx');
  });

  it('resolves URL from document library blocks by item name', () => {
    const docBlock = createBlock('document-library', 'Library', {
      kind: 'document-library',
      siteName: 'Finance',
      libraryName: 'Documents',
      items: [{ name: 'Budget.xlsx', type: 'file', url: 'https://contoso.sharepoint.com/sites/fin/Shared%20Documents/Budget.xlsx' }],
      breadcrumb: []
    });

    useGrimoireStore.setState({ blocks: [docBlock] });

    const resolved = resolveUrlFromBlocks('https://contoso.sharepoint.com/sites/fin/other.xlsx', 'Budget.xlsx');
    expect(resolved).toBe('https://contoso.sharepoint.com/sites/fin/Shared%20Documents/Budget.xlsx');
  });

  it('extracts canonical site URL for site paths', () => {
    expect(extractSiteUrl('https://contoso.sharepoint.com/sites/finance/Shared%20Documents/File.docx'))
      .toBe('https://contoso.sharepoint.com/sites/finance');
    expect(extractSiteUrl('https://contoso.sharepoint.com/teams/eng/Docs/spec.md'))
      .toBe('https://contoso.sharepoint.com/teams/eng');
    expect(extractSiteUrl('https://contoso.sharepoint.com/personal/jane_doe/Docs/notes.txt'))
      .toBe('https://contoso.sharepoint.com/personal/jane_doe');
  });

  it('falls back to origin or undefined for invalid URL', () => {
    expect(extractSiteUrl('https://contoso.sharepoint.com/random/path/file.txt'))
      .toBe('https://contoso.sharepoint.com');
    expect(extractSiteUrl('not-a-url')).toBeUndefined();
  });
});
