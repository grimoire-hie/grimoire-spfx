import { getToolAckText, isExplicitSelectionListRequest } from './ToolAcknowledgment';

describe('ToolAcknowledgment', () => {
  it('returns deterministic ack text for known tools', () => {
    expect(getToolAckText('search_sharepoint', { query: 'animals' }))
      .toBe(`Okay, I'm searching SharePoint for "animals".`);
    expect(getToolAckText('search_emails', {}))
      .toBe(`Okay, I'm pulling your emails.`);
    expect(getToolAckText('read_file_content', {}))
      .toBe(`Okay, I'm reading that file.`);
    expect(getToolAckText('read_file_content', {
      file_urls: ['https://contoso.sharepoint.com/a.docx', 'https://contoso.sharepoint.com/b.docx']
    })).toBe(`Okay, I'm reading those files.`);
  });

  it('localizes acknowledgments when the session language is non-English', () => {
    expect(getToolAckText('search_sharepoint', { query: 'SPFx' }, 'fr'))
      .toBe('D\'accord, je cherche dans SharePoint "SPFx".');
  });

  it('returns undefined for unknown tools', () => {
    expect(getToolAckText('unknown_tool', { query: 'x' })).toBeUndefined();
  });

  it('detects explicit selection-list requests', () => {
    expect(isExplicitSelectionListRequest('show options as radio buttons')).toBe(true);
    expect(isExplicitSelectionListRequest('give me a list to choose')).toBe(true);
    expect(isExplicitSelectionListRequest('search for animal docs')).toBe(false);
  });
});
