import type { IBlock, IInfoCardData, IListItemsData, IMarkdownData } from '../../models/IBlock';
import type { IMcpContent } from '../../models/IMcpTypes';
import { BlockTracker } from '../hie/BlockTracker';
import { mapMcpResultToBlock } from './McpResultMapper';

describe('McpResultMapper', () => {
  it('dedupes duplicate Outlook message ids in mail search replies', () => {
    const rawReply = [
      'Here are the most recent emails received today. Only **two** emails were found for today, so I’m listing both individually as requested.',
      '',
      '---',
      '',
      '### 1. **Subject:** AW: Recap: SPFx vs TeamsFx',
      '- **From:** Test User <user@contoso.com>',
      '- **Date:** 2026-03-06 15:01',
      '- **Preview:** First render with preview',
      '[1](https://outlook.office365.com/owa/?ItemID=AAMkTEST%2F123)',
      '',
      '---',
      '',
      '### 2. **Subject:** AW: Recap: SPFx vs TeamsFx',
      '- **From:** Test User <user@contoso.com>',
      '- **Date:** 2026-03-06 15:01',
      '[2](https://outlook.office365.com/owa/?ItemID=AAMkTEST%2F123)'
    ].join('\n');

    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({ reply: rawReply })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_MailTools', 'SearchMessages', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('markdown');

    const data = block.data as IMarkdownData;
    expect(data.content).toContain('1. **Subject:** AW: Recap: SPFx vs TeamsFx');
    expect(data.content).not.toContain('2. **Subject:** AW: Recap: SPFx vs TeamsFx');
    expect((data.content.match(/^\d+\.\s+\*\*Subject:\*\*/gm) || []).length).toBe(1);
    expect(data.itemIds).toEqual({ 1: 'AAMkTEST%2F123' });
  });

  it('normalizes bold-numbered mail replies so block tracking counts each email', () => {
    const rawReply = [
      'Here are your **3 most recent emails**, listed individually as requested:',
      '',
      '---',
      '',
      '**1.**',
      '**Subject:** How To Use This Library',
      '**From:** Test User',
      '**Date:** 2026-03-09 19:55',
      '**Preview:** First preview line',
      '[1](https://outlook.office365.com/owa/?ItemID=AAMkTEST%2F111)',
      '',
      '---',
      '',
      '**2.**',
      '**Subject:** RE: SharePoint questions',
      '**From:** Test User',
      '**Date:** 2026-03-09 18:02',
      '**Preview:** Second preview line',
      '[2](https://outlook.office365.com/owa/?ItemID=AAMkTEST%2F222)',
      '',
      '---',
      '',
      '**3.**',
      '**Subject:** Follow-up on demo',
      '**From:** John Smith',
      '**Date:** 2026-03-09 17:10',
      '**Preview:** Third preview line',
      '[3](https://outlook.office365.com/owa/?ItemID=AAMkTEST%2F333)'
    ].join('\n');

    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({ reply: rawReply })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_MailTools', 'SearchMessages', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('markdown');

    const data = block.data as IMarkdownData;
    expect(data.content).toContain('1. **Subject:** How To Use This Library');
    expect(data.content).toContain('2. **Subject:** RE: SharePoint questions');
    expect(data.content).toContain('3. **Subject:** Follow-up on demo');

    const tracker = new BlockTracker();
    tracker.track(block);
    expect(tracker.get(block.id)?.itemCount).toBe(3);
  });

  it('maps embedded JSON calendar payloads into readable markdown events', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: 'Calendar view retrieved successfully.\n{"value":[{"subject":"Prepare for community call","start":{"dateTime":"2026-03-06T14:00:00"},"end":{"dateTime":"2026-03-06T15:00:00"}}]}; CorrelationId: test'
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_CalendarTools', 'ListCalendarView', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('markdown');
    const data = block.data as IMarkdownData;
    expect(data.kind).toBe('markdown');
    expect(data.content).toContain('1. **Event:** Prepare for community call');
    expect(data.content).toContain('**Start:** 2026-03-06 14:00');
    expect(data.content).toContain('**End:** 2026-03-06 15:00');
  });

  it('maps top-level value arrays into markdown event cards', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: '{"value":[{"subject":"Daily sync","start":{"dateTime":"2026-03-07T09:30:00"},"end":{"dateTime":"2026-03-07T10:00:00"}}]}'
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_CalendarTools', 'ListCalendarView', content, pushBlock);

    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('markdown');
    const data = block.data as IMarkdownData;
    expect(data.content).toContain('1. **Event:** Daily sync');
    expect(data.content).toContain('**Start:** 2026-03-07 09:30');
  });

  it('maps calendar no-results text into an info card', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: 'No calendar events found for the given criteria.'
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_CalendarTools', 'ListCalendarView', content, pushBlock);

    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('info-card');
    const data = block.data as IInfoCardData;
    expect(data.body).toContain('No calendar events found');
  });

  it('recovers copilot text from rawResponse when reply is empty', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({
        conversationId: 'conv-123',
        reply: '',
        rawResponse: JSON.stringify({
          messages: [
            { role: 'user', text: 'Please check this URL' },
            { role: 'assistant', text: 'This GitHub profile exists and shows public repositories and profile details.' }
          ]
        })
      })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    const summary = mapMcpResultToBlock('mcp_M365Copilot', 'copilot_chat', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('markdown');
    const data = block.data as IMarkdownData;
    expect(data.content).toContain('This GitHub profile exists');
    expect(summary).toContain('This GitHub profile exists');
  });

  it('does not misreport empty copilot replies as no public results when payload has no text', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({
        conversationId: 'conv-123',
        timeZone: 'America/Los_Angeles',
        reply: '',
        rawResponse: '{"conversationId":"conv-123"}; CorrelationId: abc, TimeStamp: 2026-03-07_08:26:43'
      })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    const summary = mapMcpResultToBlock('mcp_M365Copilot', 'copilot_chat', content, pushBlock);

    expect(pushBlock).not.toHaveBeenCalled();
    expect(summary).toContain('Copilot returned an empty reply');
    expect(summary).not.toContain('The tool returned no results');
  });

  it('unwraps embedded response JSON for SharePoint list tools', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({
        message: 'Graph tool executed successfully.',
        response: JSON.stringify({
          value: [
            {
              id: '1',
              displayName: 'Documents',
              description: 'Shared docs',
              webUrl: 'https://tenant.sharepoint.com/sites/copilot-test/Shared%20Documents',
              '@odata.etag': '"1,22"',
              createdBy: { user: { displayName: 'Systemkonto' } },
              lastModifiedBy: { user: { displayName: 'SharePoint-App' } },
              parentReference: { siteId: 'tenant,site,web' },
              list: { template: 'documentLibrary', hidden: false }
            },
            {
              id: '2',
              displayName: 'Site Assets',
              description: 'Assets library',
              webUrl: 'https://tenant.sharepoint.com/sites/copilot-test/SiteAssets',
              createdBy: { user: { displayName: 'Systemkonto' } },
              list: { template: 'documentLibrary', hidden: false }
            }
          ]
        })
      })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_SharePointListsTools', 'listLists', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('list-items');
    expect(block.title).toBe('Lists');

    const data = block.data as IListItemsData;
    expect(data.totalCount).toBe(2);
    expect(data.items[0].displayName).toBe('Documents');
    expect(data.items[1].displayName).toBe('Site Assets');
    expect(data.items[0].createdBy).toBe('Systemkonto');
    expect(data.items[0].lastModifiedBy).toBe('SharePoint-App');
    expect(data.items[0].template).toBe('documentLibrary');
    expect(data.columns).toContain('displayName');
    expect(data.columns).toContain('template');
    expect(data.columns).toContain('createdBy');
    expect(data.columns).not.toContain('@odata.etag');
    expect(data.columns).not.toContain('parentReference');
  });

  it('renders successful createList results as a readable info card', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({
        message: 'Graph tool executed successfully.',
        response: JSON.stringify({
          id: '74584e25-9845-43d7-bbae-0c73cb8451cf',
          displayName: 'Project Tracking',
          webUrl: 'https://tenant.sharepoint.com/sites/copilot-test/Lists/Project%20Tracking',
          createdDateTime: '2026-03-14T08:15:56Z',
          parentReference: {
            siteId: 'tenant,site,web'
          },
          list: {
            template: 'genericList',
            hidden: false
          }
        })
      })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_SharePointListsTools', 'createList', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('info-card');
    expect(block.title).toBe('Project Tracking');

    const data = block.data as IInfoCardData;
    expect(data.heading).toBe('Project Tracking');
    expect(data.body).toContain('Template: genericList');
    expect(data.body).toContain('Created: 2026-03-14T08:15:56Z');
    expect(data.body).toContain('Open: https://tenant.sharepoint.com/sites/copilot-test/Lists/Project%20Tracking');
    expect(data.url).toBe('https://tenant.sharepoint.com/sites/copilot-test/Lists/Project%20Tracking');
    expect(data.targetContext).toEqual(expect.objectContaining({
      siteId: 'tenant,site,web',
      listId: '74584e25-9845-43d7-bbae-0c73cb8451cf',
      listName: 'Project Tracking',
      listUrl: 'https://tenant.sharepoint.com/sites/copilot-test/Lists/Project%20Tracking'
    }));
  });

  it('renders successful createListColumn results as a readable info card', () => {
    const content: IMcpContent[] = [{
      type: 'text',
      text: JSON.stringify({
        message: 'Graph tool executed successfully.',
        response: JSON.stringify({
          '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#sites(\'tenant.sharepoint.com%2Csite%2Cweb\')/lists(\'list-123\')/columns/$entity',
          id: 'column-123',
          columnGroup: 'Custom Columns',
          displayName: 'mynotes',
          name: 'mynotes',
          enforceUniqueValues: false,
          hidden: false,
          indexed: false,
          readOnly: false,
          required: false,
          text: {
            allowMultipleLines: false,
            maxLength: 255
          }
        })
      })
    }];

    const pushBlock = jest.fn<void, [IBlock]>();

    mapMcpResultToBlock('mcp_SharePointListsTools', 'createListColumn', content, pushBlock);

    expect(pushBlock).toHaveBeenCalledTimes(1);
    const block = pushBlock.mock.calls[0][0];
    expect(block.type).toBe('info-card');
    expect(block.title).toBe('mynotes');

    const data = block.data as IInfoCardData;
    expect(data.heading).toBe('mynotes');
    expect(data.body).toContain('Type: Text');
    expect(data.body).toContain('Internal name: mynotes');
    expect(data.body).toContain('Required: No');
    expect(data.body).not.toContain('"columnGroup"');
    expect(data.targetContext).toEqual(expect.objectContaining({
      siteId: 'tenant.sharepoint.com,site,web',
      listId: 'list-123',
      columnId: 'column-123',
      columnName: 'mynotes',
      columnDisplayName: 'mynotes'
    }));
  });
});
