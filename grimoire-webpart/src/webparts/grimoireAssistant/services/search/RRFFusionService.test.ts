jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

import type { ICopilotSearchResult } from '../../models/ISearchTypes';
import { RRFFusionService } from './RRFFusionService';

function makeResult(
  title: string,
  source: ICopilotSearchResult['source'],
  overrides: Partial<ICopilotSearchResult> = {}
): ICopilotSearchResult {
  return {
    title,
    summary: `${title} summary`,
    url: `https://tenant.sharepoint.com/sites/test/${encodeURIComponent(title)}.docx`,
    source,
    ...overrides
  };
}

describe('RRFFusionService', () => {
  it('favors corroborated semantic results over sharepoint-only candidates', () => {
    const service = new RRFFusionService();

    const fused = service.fuseWithContext([
      [makeResult('Animal Stories and Facts', 'copilot-search')],
      [makeResult('Animal Stories and Facts', 'copilot-retrieval', { sourceNativeScore: 0.9 })],
      [
        makeResult('Renewable Energy Policy', 'sharepoint-search', {
          summary: 'Solar power and policy outlook',
          sourceNativeScore: 120
        }),
        makeResult('Animal Stories and Facts', 'sharepoint-search', {
          summary: 'Animal facts and stories',
          sourceNativeScore: 110
        })
      ]
    ], {
      queryText: 'docs about animals',
      queryLanguage: 'en'
    });

    expect(fused[0].title).toBe('Animal Stories and Facts');
    expect(fused[1].title).toBe('Renewable Energy Policy');
  });

  it('keeps semantic-rewrite contributions close to raw-query weight', () => {
    const service = new RRFFusionService();

    const fused = service.fuseWithContext([
      [
        makeResult('Animal Stories and Facts', 'copilot-search', {
          variantKind: 'semantic-rewrite',
          variantLanguage: 'en'
        })
      ],
      [
        makeResult('Animal Stories and Facts', 'copilot-retrieval', {
          variantKind: 'raw',
          variantLanguage: 'en',
          sourceNativeScore: 0.8
        })
      ]
    ], {
      queryText: 'docs about animals',
      queryLanguage: 'en'
    });

    expect(fused[0].title).toBe('Animal Stories and Facts');
    expect(fused[0].rrfScore).toBeGreaterThan(0);
  });
});
