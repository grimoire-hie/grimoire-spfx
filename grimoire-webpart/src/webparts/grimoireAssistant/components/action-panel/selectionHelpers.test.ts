import { createBlock, type IInfoCardData } from '../../models/IBlock';
import { getSelectionCandidates } from './selectionHelpers';

describe('selectionHelpers info-card candidates', () => {
  it('preserves info-card URLs and target context for follow-up actions', () => {
    const block = createBlock('info-card', 'testnellov1', {
      kind: 'info-card',
      heading: 'testnellov1',
      body: 'Template: genericList',
      url: 'https://tenant.sharepoint.com/sites/ProjectNova/Lists/testnellov1',
      targetContext: {
        siteId: 'tenant,site,web',
        siteUrl: 'https://tenant.sharepoint.com/sites/ProjectNova',
        listId: 'list-testnellov1',
        listName: 'testnellov1',
        listUrl: 'https://tenant.sharepoint.com/sites/ProjectNova/Lists/testnellov1'
      }
    } as IInfoCardData);

    const candidates = getSelectionCandidates(block);

    expect(candidates).toHaveLength(1);
    expect(candidates[0].url).toBe('https://tenant.sharepoint.com/sites/ProjectNova/Lists/testnellov1');
    expect(candidates[0].payload).toEqual(expect.objectContaining({
      url: 'https://tenant.sharepoint.com/sites/ProjectNova/Lists/testnellov1',
      targetContext: expect.objectContaining({
        siteId: 'tenant,site,web',
        listId: 'list-testnellov1',
        listName: 'testnellov1'
      })
    }));
  });
});
