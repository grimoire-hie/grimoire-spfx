import { HieInteractionFormatter } from './HieInteractionFormatter';
import { ASSISTANT_SUMMARY_TARGET_TEXT } from '../../config/assistantLengthLimits';

describe('HieInteractionFormatter', () => {
  it('includes selected item metadata for selection-list interactions', () => {
    const formatter = new HieInteractionFormatter({} as never);

    const message = formatter.format({
      blockId: 'block-1',
      blockType: 'selection-list',
      action: 'select',
      payload: {
        label: 'copilot-test',
        prompt: 'Found 3 sites for "copilot-test"',
        selectedIds: ['https://contoso.sharepoint.com/sites/copilot-test'],
        selectedItems: [
          {
            id: 'https://contoso.sharepoint.com/sites/copilot-test',
            label: 'copilot-test',
            description: 'https://contoso.sharepoint.com/sites/copilot-test'
          }
        ]
      },
      timestamp: Date.now(),
      eventName: 'block.interaction.select',
      schemaId: 'selection.select',
      source: 'block-ui',
      surface: 'action-panel',
      exposurePolicy: { mode: 'response-triggering', relevance: 'foreground' },
      correlationId: 'iact-test'
    });

    expect(message?.kind).toBe('interaction');
    expect(message?.body).toContain('copilot-test');
    expect(message?.body).toContain('selectedItems');
    expect(message?.body).toContain('https://contoso.sharepoint.com/sites/copilot-test');
  });

  it('includes nested row data for list row interactions', () => {
    const formatter = new HieInteractionFormatter({} as never);

    const message = formatter.format({
      blockId: 'block-2',
      blockType: 'list-items',
      action: 'click-list-row',
      payload: {
        index: 2,
        rowData: {
          displayName: 'Events',
          id: '00000000-0000-0000-0000-000000000003',
          webUrl: 'https://contoso.sharepoint.com/sites/copilot-test/Lists/Events',
          template: 'events'
        }
      },
      timestamp: Date.now(),
      eventName: 'block.interaction.click-list-row',
      schemaId: 'list-items.click-list-row',
      source: 'block-ui',
      surface: 'action-panel',
      exposurePolicy: { mode: 'response-triggering', relevance: 'foreground' },
      correlationId: 'iact-row'
    });

    expect(message?.kind).toBe('interaction');
    expect(message?.body).toContain('Events');
    expect(message?.body).toContain('https://contoso.sharepoint.com/sites/copilot-test/Lists/Events');
    expect(message?.body).toContain('rowData');
    expect(message?.body).toContain('"index":2');
  });

  it('uses the shared summarize target guidance in interaction instructions', () => {
    const formatter = new HieInteractionFormatter({} as never);

    const message = formatter.format({
      blockId: 'block-3',
      blockType: 'search-results',
      action: 'summarize',
      payload: {
        title: 'Migration Plan',
        url: 'https://contoso.sharepoint.com/sites/eng/Shared%20Documents/Migration%20Plan.docx'
      },
      timestamp: Date.now(),
      eventName: 'block.interaction.summarize',
      schemaId: 'hover.summarize',
      source: 'block-ui',
      surface: 'action-panel',
      exposurePolicy: { mode: 'response-triggering', relevance: 'foreground' },
      correlationId: 'iact-summarize'
    });

    expect(message?.kind).toBe('interaction');
    expect(message?.body).toContain(ASSISTANT_SUMMARY_TARGET_TEXT);
    expect(message?.body).not.toContain('700-1200 characters');
  });
});
