import {
  extractSharePointSiteUrl,
  inferDocumentLibraryBaseUrl,
  isSharePointViewerUrl,
  resolveDocumentLibraryItemUrl
} from './sharePointUrlUtils';

describe('sharePointUrlUtils', () => {
  it('detects Office viewer URLs', () => {
    expect(isSharePointViewerUrl(
      'https://tenant.sharepoint.com/sites/cooking/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Recipe.docx&action=default'
    )).toBe(true);
    expect(isSharePointViewerUrl(
      'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/Recipe.docx'
    )).toBe(false);
  });

  it('infers the document-library base URL from sibling items', () => {
    const baseUrl = inferDocumentLibraryBaseUrl([
      'https://tenant.sharepoint.com/sites/cooking/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Recipe.docx&action=default',
      'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/chedice',
      'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/Washoku.pdf'
    ]);

    expect(baseUrl).toBe('https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente');
  });

  it('converts document-library viewer URLs into canonical file URLs when siblings expose the library path', () => {
    const resolved = resolveDocumentLibraryItemUrl(
      'https://tenant.sharepoint.com/sites/cooking/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Brotkultur_Deutschsprachige_Laender_DE.docx&action=default',
      'Brotkultur_Deutschsprachige_Laender_DE.docx',
      [
        'https://tenant.sharepoint.com/sites/cooking/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Brotkultur_Deutschsprachige_Laender_DE.docx&action=default',
        'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/chedice',
        'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/Washoku_Japanese_Cuisine_Philosophy_JA.pdf'
      ]
    );

    expect(resolved).toBe(
      'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/Brotkultur_Deutschsprachige_Laender_DE.docx'
    );
  });

  it('extracts the canonical SharePoint site URL from item URLs', () => {
    expect(extractSharePointSiteUrl(
      'https://tenant.sharepoint.com/sites/cooking/Freigegebene%20Dokumente/Recipe.docx'
    )).toBe('https://tenant.sharepoint.com/sites/cooking');
    expect(extractSharePointSiteUrl(
      'https://tenant.sharepoint.com/sites/cooking/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Recipe.docx&action=default'
    )).toBe('https://tenant.sharepoint.com/sites/cooking');
  });
});
