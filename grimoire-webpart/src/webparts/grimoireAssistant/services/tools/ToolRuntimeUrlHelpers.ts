import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IDocumentLibraryData, ISearchResultsData } from '../../models/IBlock';
import { logService } from '../logging/LogService';

export function resolveUrlFromBlocks(url: string, name: string | undefined): string | undefined {
  const blocks = useGrimoireStore.getState().blocks;
  let urlFileName = '';
  try {
    urlFileName = decodeURIComponent(url.split('/').pop() || '').toLowerCase();
  } catch {
    // keep empty
  }

  const nameToMatch = (name || urlFileName || '').toLowerCase().replace(/\.[^.]+$/, '');
  if (!nameToMatch) return undefined;

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === 'search-results') {
      const data = block.data as ISearchResultsData;
      for (let j = 0; j < data.results.length; j++) {
        const r = data.results[j];
        const rTitle = r.title.toLowerCase().replace(/\.[^.]+$/, '');
        if (rTitle === nameToMatch || r.title.toLowerCase() === (name || '').toLowerCase()) {
          if (r.url && r.url !== url) {
            logService.info('graph', `Resolved URL from search results: "${name}" → ${r.url}`);
            return r.url;
          }
        }
      }
    } else if (block.type === 'document-library') {
      const data = block.data as IDocumentLibraryData;
      for (let j = 0; j < data.items.length; j++) {
        const item = data.items[j];
        const itemName = item.name.toLowerCase().replace(/\.[^.]+$/, '');
        if (itemName === nameToMatch || item.name.toLowerCase() === (name || '').toLowerCase()) {
          if (item.url && item.url !== url && item.url.indexOf('/_layouts/') === -1) {
            logService.info('graph', `Resolved URL from library block: "${name}" → ${item.url}`);
            return item.url;
          }
        }
      }
    }
  }

  return undefined;
}

export function extractSiteUrl(fileUrl: string): string | undefined {
  try {
    const url = new URL(fileUrl);
    const segments = url.pathname.split('/').filter(Boolean);
    if (segments.length >= 2 && (segments[0] === 'sites' || segments[0] === 'teams' || segments[0] === 'personal')) {
      return `${url.origin}/${segments[0]}/${segments[1]}`;
    }
    return url.origin;
  } catch {
    return undefined;
  }
}
