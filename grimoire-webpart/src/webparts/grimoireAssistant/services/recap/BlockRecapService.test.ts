jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('../nano/NanoService', () => ({
  getNanoService: jest.fn()
}));

import { createBlock } from '../../models/IBlock';
import { getNanoService } from '../nano/NanoService';
import { BlockRecapService, canRecapBlock } from './BlockRecapService';

describe('BlockRecapService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('supports recap for selection lists and other result blocks', () => {
    const block = createBlock('selection-list', 'Pick a site', {
      kind: 'selection-list',
      prompt: 'Choose a site',
      multiSelect: false,
      items: [
        { id: '1', label: 'Finance', description: 'Finance collaboration site' },
        { id: '2', label: 'HR', description: 'Human resources site' }
      ]
    });

    const service = new BlockRecapService();
    const input = service.buildRecapInput(block);

    expect(canRecapBlock(block)).toBe(true);
    expect(input.prompt).toBe('Choose a site');
    expect(input.items).toEqual([
      expect.objectContaining({ title: 'Finance', summary: 'Finance collaboration site' }),
      expect.objectContaining({ title: 'HR', summary: 'Human resources site' })
    ]);
  });

  it('does not allow recap cards to recap themselves', () => {
    const block = createBlock('info-card', 'Recap: Search', {
      kind: 'info-card',
      heading: 'Recap: Search',
      body: 'Already summarized.'
    }, true, undefined, { originTool: 'block-recap:block-1' });

    expect(canRecapBlock(block)).toBe(false);
  });

  it('includes all visible search results in the recap input instead of silently trimming to six', () => {
    const results = Array.from({ length: 10 }, (_, index) => ({
      title: `Result ${index + 1}`,
      summary: `Summary for result ${index + 1}.`,
      url: `https://tenant.sharepoint.com/sites/dev/result-${index + 1}.docx`,
      sources: ['copilot-search']
    }));

    const block = createBlock('search-results', 'Search: roadmap', {
      kind: 'search-results',
      query: 'roadmap',
      totalCount: results.length,
      source: 'copilot-search',
      results
    });

    const service = new BlockRecapService();
    const input = service.buildRecapInput(block);

    expect(input.items).toHaveLength(10);
    expect(input.items[0]).toEqual(expect.objectContaining({ title: 'Result 1' }));
    expect(input.items[9]).toEqual(expect.objectContaining({ title: 'Result 10' }));
  });

  it('builds a deterministic fallback that can call out broader search tail items', () => {
    const block = createBlock('search-results', 'Search: animals', {
      kind: 'search-results',
      query: 'i am searching for info about animals',
      totalCount: 4,
      source: 'copilot-search+copilot-retrieval',
      results: [
        {
          title: 'Animal Stories and Facts',
          summary: 'Reference document about elephants, lions, wolves, and dolphins.',
          url: 'https://tenant.sharepoint.com/sites/dev/AnimalStories.docx',
          sources: ['copilot-search', 'copilot-retrieval']
        },
        {
          title: 'Histoires et Faits sur les Animaux',
          summary: 'French reference file about animals.',
          url: 'https://tenant.sharepoint.com/sites/dev/Animaux.docx',
          sources: ['copilot-search', 'copilot-retrieval']
        },
        {
          title: 'Farm_to_Table_Sustainable_Cooking_EN',
          summary: 'Long-form article about sustainable cooking and food systems.',
          url: 'https://tenant.sharepoint.com/sites/cooking/Farm.docx',
          sources: ['copilot-search']
        },
        {
          title: 'How To Use This Library',
          summary: 'Instructions page.',
          url: 'https://tenant.sharepoint.com/sites/appcatalog/howto.aspx',
          sources: ['copilot-retrieval']
        }
      ]
    });

    const service = new BlockRecapService();
    const text = service.buildFallbackRecap(service.buildRecapInput(block));

    expect(text).toContain('This recap covers all 4 visible results for "animals"');
    expect(text).toContain('The visible results are focused on "animals"');
    expect(text).toContain('"Animal Stories and Facts"');
    expect(text).toContain('"Histoires et Faits sur les Animaux"');
    expect(text).toContain('Result #4 is "How To Use This Library"');
    expect(text).toContain('Some lower-ranked items look broader than the strongest matches.');
  });

  it('uses the fast model when available and falls back when not', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue('The data mainly contain two site choices. Finance looks like the primary collaboration site, while HR is the secondary option.')
    });

    const block = createBlock('selection-list', 'Pick a site', {
      kind: 'selection-list',
      prompt: 'Choose a site',
      multiSelect: false,
      items: [
        { id: '1', label: 'Finance', description: 'Finance collaboration site' },
        { id: '2', label: 'HR', description: 'Human resources site' }
      ]
    });

    const service = new BlockRecapService();
    const text = await service.generate(block, {
      proxyUrl: 'https://example.com/api',
      proxyApiKey: 'test',
      backend: 'reasoning',
      deployment: 'grimoire-reasoning',
      apiVersion: '2024-10-21'
    });

    expect(text).toContain('Finance looks like the primary collaboration site');

    (getNanoService as jest.Mock).mockReturnValue(undefined);
    const fallback = await service.generate(block, undefined);
    expect(fallback).toContain('This recap covers all 2 visible selection list entries related to "Choose a site".');
    expect(fallback).toContain('Finance');
  });

  it('formats recap output into a lead sentence plus bullets for readability', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(
        'The visible results are tightly focused on SPFx. The strongest hits are SPFx, SPFx_de, and SPFx_ja. TeamsFx is related tooling rather than the same topic.'
      )
    });

    const block = createBlock('search-results', 'Search: SPFx', {
      kind: 'search-results',
      query: 'SPFx',
      totalCount: 4,
      source: 'copilot-search+copilot-retrieval+sharepoint-search',
      results: [
        {
          title: 'SPFx',
          summary: 'English overview of SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
          sources: ['copilot-search', 'copilot-retrieval']
        },
        {
          title: 'SPFx_de',
          summary: 'German overview of SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf',
          sources: ['copilot-search', 'copilot-retrieval']
        },
        {
          title: 'SPFx_ja',
          summary: 'Japanese overview of SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
          sources: ['copilot-retrieval']
        },
        {
          title: 'TeamsFx',
          summary: 'Teams toolkit overview.',
          url: 'https://tenant.sharepoint.com/sites/dev/TeamsFx.pdf',
          sources: ['sharepoint-search']
        }
      ]
    });

    const service = new BlockRecapService();
    const text = await service.generate(block, {
      proxyUrl: 'https://example.com/api',
      proxyApiKey: 'test',
      backend: 'reasoning',
      deployment: 'grimoire-reasoning',
      apiVersion: '2024-10-21'
    });

    expect(text).toContain('\n\n- ');
    expect(text.startsWith('The visible results are tightly focused on SPFx.')).toBe(true);
  });

  it('retries recap generation when the fast model returns empty content', async () => {
    const classify = jest.fn()
      .mockResolvedValueOnce('')
      .mockResolvedValueOnce('The visible results are tightly focused on SPFx. The strongest hits are the English, German, and Japanese SPFx documents.');

    (getNanoService as jest.Mock).mockReturnValue({ classify });

    const block = createBlock('search-results', 'Search: spfx', {
      kind: 'search-results',
      query: 'i am searching for info about spfx',
      totalCount: 3,
      source: 'copilot-search+copilot-retrieval',
      results: [
        {
          title: 'SPFx',
          summary: 'English reference document about SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
          sources: ['copilot-search', 'copilot-retrieval']
        },
        {
          title: 'SPFx_de',
          summary: 'German reference document about SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf',
          sources: ['copilot-search', 'copilot-retrieval']
        },
        {
          title: 'SPFx_ja',
          summary: 'Japanese reference document about SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
          sources: ['copilot-retrieval']
        }
      ]
    });

    const service = new BlockRecapService();
    const text = await service.generate(block, {
      proxyUrl: 'https://example.com/api',
      proxyApiKey: 'test',
      backend: 'reasoning',
      deployment: 'grimoire-reasoning',
      apiVersion: '2024-10-21'
    });

    expect(classify).toHaveBeenCalledTimes(2);
    expect(text).toContain('tightly focused on SPFx');
  });

  it('passes every visible result into the recap payload even when compaction is needed', async () => {
    const classify = jest.fn().mockResolvedValue(
      'This is a long recap for the visible set. It covers the strongest results and keeps the broader tail visible as well.'
    );
    (getNanoService as jest.Mock).mockReturnValue({ classify });

    const results = Array.from({ length: 100 }, (_, index) => ({
      title: `Visible Result ${index + 1} with a deliberately verbose title that keeps the payload large enough to require compaction`,
      summary: `This is the visible summary for result ${index + 1}. It contains extra detail so the recap payload has to compact lower-ranked entries without dropping them entirely.`,
      url: `https://tenant.sharepoint.com/sites/dev/visible-result-${index + 1}.docx`,
      sources: ['copilot-search', 'copilot-retrieval'],
      language: index % 2 === 0 ? 'en' : 'de',
      author: 'Adele Vance',
      siteName: 'Engineering'
    }));

    const block = createBlock('search-results', 'Search: migration', {
      kind: 'search-results',
      query: 'migration',
      totalCount: results.length,
      source: 'copilot-search+copilot-retrieval',
      results
    });

    const service = new BlockRecapService();
    await service.generate(block, {
      proxyUrl: 'https://example.com/api',
      proxyApiKey: 'test',
      backend: 'reasoning',
      deployment: 'grimoire-reasoning',
      apiVersion: '2024-10-21'
    });

    const payload = JSON.parse(classify.mock.calls[0][1] as string);
    expect(payload.visibleItemCount).toBe(100);
    expect(payload.items).toHaveLength(100);
    expect(payload.items[0].title).toContain('Visible Result 1');
    expect(payload.items[99].title).toContain('Visible Result 100');
    expect(payload.notes.join(' ')).toContain('represented more compactly');
  });

  it('builds a longer fallback recap that still references later visible items', () => {
    const results = Array.from({ length: 10 }, (_, index) => ({
      title: `Search Result ${index + 1}`,
      summary: `Summary snippet for search result ${index + 1} about governance, rollout, and migration work.`,
      url: `https://tenant.sharepoint.com/sites/dev/search-result-${index + 1}.docx`,
      sources: ['copilot-search'],
      language: 'en'
    }));

    const block = createBlock('search-results', 'Search: governance', {
      kind: 'search-results',
      query: 'governance',
      totalCount: results.length,
      source: 'copilot-search',
      results
    });

    const service = new BlockRecapService();
    const text = service.buildFallbackRecap(service.buildRecapInput(block));

    expect(text.length).toBeGreaterThan(800);
    expect(text).toContain('This recap covers all 10 visible results for "governance"');
    expect(text).toContain('#10 "Search Result 10"');
  });

  it('builds a search fallback that reads like a recap for tight multilingual result sets', () => {
    const block = createBlock('search-results', 'Search: spfx', {
      kind: 'search-results',
      query: 'i am searching for info about spfx',
      totalCount: 3,
      source: 'copilot-search+copilot-retrieval',
      results: [
        {
          title: 'SPFx',
          summary: 'English reference document about SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx.pdf',
          sources: ['copilot-search', 'copilot-retrieval'],
          language: 'en'
        },
        {
          title: 'SPFx_de',
          summary: 'German reference document about SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf',
          sources: ['copilot-search', 'copilot-retrieval'],
          language: 'de'
        },
        {
          title: 'SPFx_ja',
          summary: 'Japanese reference document about SharePoint Framework.',
          url: 'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
          sources: ['copilot-retrieval'],
          language: 'ja'
        }
      ]
    });

    const service = new BlockRecapService();
    const text = service.buildFallbackRecap(service.buildRecapInput(block));

    expect(text).toContain('This recap covers all 3 visible results for "spfx"');
    expect(text).toContain('The visible results are tightly focused on "spfx"');
    expect(text).toContain('The strongest hits are "SPFx", "SPFx_de", and "SPFx_ja"');
    expect(text).toContain('The visible results span EN, DE, JA');
  });
});
