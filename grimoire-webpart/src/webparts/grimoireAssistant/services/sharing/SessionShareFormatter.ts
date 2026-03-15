import type {
  IBlock as IPanelBlock,
  IDocumentLibraryData,
  IErrorData,
  IFilePreviewData,
  IInfoCardData,
  IListItemsData,
  ISearchResult,
  ISearchResultsData,
  ISiteInfoData
} from '../../models/IBlock';
import type { ITranscriptEntry } from '../../store/useGrimoireStore';
import type { IBlockRecapItem } from '../recap/BlockRecapService';
import { BlockRecapService, canRecapBlock } from '../recap/BlockRecapService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { resolveCurrentArtifactContext } from '../hie/HieArtifactLinkage';
import { formatSearchQueryBreadthLine } from '../search/SearchQueryVariantPresentation';
import { SHARE_LENGTH_LIMITS } from '../../config/assistantLengthLimits';

export interface ISessionShareContent {
  subject: string;
  plainText: string;
  markdown: string;
  detailedPlainText: string;
  emailPlainText: string;
  attachmentUris: string[];
}

export interface ISessionShareOptions {
  blocks: IPanelBlock[];
  transcript: ITranscriptEntry[];
  activeBlockId?: string;
  selectedActionIndices?: number[];
  selectionBehavior?: 'default' | 'strict';
}

const MAX_TRANSCRIPT_ENTRIES = 6;
const MAX_TRANSCRIPT_CHARS = SHARE_LENGTH_LIMITS.transcriptMaxChars;
const MAX_BLOCK_SUMMARY_CHARS = SHARE_LENGTH_LIMITS.blockSummaryMaxChars;
const MAX_DETAILED_BLOCKS = 3;
const MAX_DETAILED_ITEMS = 6;

function clip(value: string, maxChars: number): string {
  const trimmed = value.trim();
  if (trimmed.length <= maxChars) return trimmed;
  return `${trimmed.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

function escapeMarkdown(value: string): string {
  return value.replace(/[\\`*_{}[\]()#+.!-]/g, '\\$&');
}

function humanizeBlockType(type: string): string {
  return type.replace(/-/g, ' ');
}

function dedupeStrings(values: string[]): string[] {
  const seen = new Set<string>();
  const deduped: string[] = [];

  values.forEach((value) => {
    const trimmed = value.trim();
    if (!trimmed || seen.has(trimmed)) return;
    seen.add(trimmed);
    deduped.push(trimmed);
  });

  return deduped;
}

function formatDetailLine(label: string, value: string | undefined): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  return `${label}: ${trimmed}`;
}

function selectItemsByIndex<TItem>(
  items: TItem[],
  selectedActionIndices?: number[],
  selectionBehavior: 'default' | 'strict' = 'default'
): TItem[] {
  if (!selectedActionIndices || selectedActionIndices.length === 0) {
    return selectionBehavior === 'strict' ? [] : items.slice(0, MAX_DETAILED_ITEMS);
  }

  const selectedItems: TItem[] = [];
  const seen = new Set<number>();
  selectedActionIndices.forEach((index) => {
    const zeroBased = index - 1;
    if (zeroBased < 0 || zeroBased >= items.length || seen.has(zeroBased)) return;
    seen.add(zeroBased);
    selectedItems.push(items[zeroBased]);
  });

  if (selectedItems.length > 0) {
    return selectedItems;
  }

  return selectionBehavior === 'strict' ? [] : items.slice(0, MAX_DETAILED_ITEMS);
}

function formatSearchResult(result: ISearchResult): string[] {
  return [
    result.title,
    ...[
      formatDetailLine('Type', result.fileType),
      formatDetailLine('Author', result.author),
      formatDetailLine('Modified', result.lastModified),
      formatDetailLine('Site', result.siteName),
      result.sources && result.sources.length > 0 ? `Sources: ${result.sources.join(', ')}` : undefined,
      formatDetailLine('URL', result.url)
    ].filter((line): line is string => !!line)
  ];
}

function formatDocumentLibraryItem(item: IDocumentLibraryData['items'][number]): string[] {
  return [
    item.name,
    ...[
      formatDetailLine('Kind', item.type),
      formatDetailLine('Type', item.fileType),
      formatDetailLine('Author', item.author),
      formatDetailLine('Modified', item.lastModified),
      formatDetailLine('URL', item.url)
    ].filter((line): line is string => !!line)
  ];
}

function formatFilePreview(data: IFilePreviewData): string[] {
  return [
    data.fileName,
    ...[
      formatDetailLine('Type', data.fileType),
      formatDetailLine('Author', data.author),
      formatDetailLine('Modified', data.lastModified),
      data.size ? `Size: ${String(data.size)}` : undefined,
      formatDetailLine('URL', data.fileUrl)
    ].filter((line): line is string => !!line)
  ];
}

function formatSiteInfo(data: ISiteInfoData): string[] {
  return [
    data.siteName,
    ...[
      formatDetailLine('Description', data.description),
      formatDetailLine('Owner', data.owner),
      formatDetailLine('Created', data.created),
      formatDetailLine('Modified', data.lastModified),
      formatDetailLine('URL', data.siteUrl)
    ].filter((line): line is string => !!line)
  ];
}

function formatListItem(item: Record<string, string>, index: number): string[] {
  const entries = Object.entries(item).filter(([, value]) => value && value.trim());
  if (entries.length === 0) {
    return [`Item ${index + 1}`];
  }

  const lines: string[] = [];
  entries.slice(0, 5).forEach(([key, value], entryIndex) => {
    lines.push(entryIndex === 0 ? `${key}: ${value}` : `${key}: ${value}`);
  });
  return lines;
}

function formatRecapItem(item: IBlockRecapItem): string[] {
  const lines = [item.title.trim()];
  const summary = item.summary?.trim();
  if (summary) {
    lines.push(summary);
  }
  (item.details || []).forEach((detail) => {
    const trimmed = detail.trim();
    if (trimmed) {
      lines.push(trimmed);
    }
  });
  return lines;
}

function isAttachableFileUrl(url: string): boolean {
  if (!url.trim()) return false;
  try {
    const parsed = new URL(url);
    if (!/(?:sharepoint|onedrive)\.com$/i.test(parsed.hostname)) {
      return false;
    }
    return /\.[a-z0-9]{2,8}$/i.test(parsed.pathname);
  } catch {
    return false;
  }
}

function collectBlockAttachmentUris(
  block: IPanelBlock,
  selectedActionIndices?: number[],
  selectionBehavior: 'default' | 'strict' = 'default'
): string[] {
  switch (block.type) {
    case 'search-results': {
      const data = block.data as ISearchResultsData;
      return selectItemsByIndex(data.results, selectedActionIndices, selectionBehavior)
        .map((result) => result.url)
        .filter(isAttachableFileUrl);
    }
    case 'document-library': {
      const data = block.data as IDocumentLibraryData;
      return selectItemsByIndex(data.items, selectedActionIndices, selectionBehavior)
        .filter((item) => item.type === 'file')
        .map((item) => item.url)
        .filter(isAttachableFileUrl);
    }
    case 'file-preview': {
      const data = block.data as IFilePreviewData;
      return isAttachableFileUrl(data.fileUrl) ? [data.fileUrl] : [];
    }
    default:
      return [];
  }
}

export function isShareableBlock(block: IPanelBlock): boolean {
  return block.type !== 'form' && block.type !== 'confirmation-dialog';
}

export function hasShareableSessionContent(blocks: IPanelBlock[], transcript: ITranscriptEntry[]): boolean {
  return blocks.some(isShareableBlock) || transcript.some((entry) => entry.role !== 'system' && entry.text.trim().length > 0);
}

function buildErrorSummary(block: IPanelBlock): string {
  const data = block.data as IErrorData;
  return clip(data.message || block.title, MAX_BLOCK_SUMMARY_CHARS);
}

function buildSearchResultsSummary(block: IPanelBlock): string {
  const data = block.data as ISearchResultsData;
  const visibleCount = data.results.length;
  const parts = [
    `${visibleCount} visible result${visibleCount === 1 ? '' : 's'} for "${data.query}".`
  ];
  const searchBreadthLine = formatSearchQueryBreadthLine(data.queryVariants);
  if (searchBreadthLine) {
    parts.push(searchBreadthLine);
  }

  const topTitles = data.results
    .slice(0, 3)
    .map((result) => result.title.trim())
    .filter(Boolean);
  if (topTitles.length > 0) {
    parts.push(
      topTitles.length === 1
        ? `Top visible hit: "${topTitles[0]}".`
        : `Top visible hits: ${topTitles.map((title) => `"${title}"`).join(', ')}.`
    );
  }

  return clip(parts.join(' '), MAX_BLOCK_SUMMARY_CHARS);
}

function describeBlock(block: IPanelBlock, recapService: BlockRecapService): string {
  if (block.type === 'error') {
    return buildErrorSummary(block);
  }

  if (block.type === 'search-results') {
    return buildSearchResultsSummary(block);
  }

  if (block.type === 'info-card') {
    const data = block.data as IInfoCardData;
    return clip(data.body || data.heading || block.title, MAX_BLOCK_SUMMARY_CHARS);
  }

  if (canRecapBlock(block)) {
    const input = recapService.buildRecapInput(block);
    return clip(recapService.buildFallbackRecap(input), MAX_BLOCK_SUMMARY_CHARS);
  }

  return clip(`${block.title} (${humanizeBlockType(block.type)})`, MAX_BLOCK_SUMMARY_CHARS);
}

interface IResolvedShareScope {
  primaryBlock?: IPanelBlock;
  orderedBlocks: IPanelBlock[];
  selectedActionIndices?: number[];
  sourceTurnId?: string;
  sourceRootTurnId?: string;
}

function dedupeBlocks(blocks: IPanelBlock[]): IPanelBlock[] {
  const seen = new Set<string>();
  return blocks.filter((block) => {
    if (seen.has(block.id)) return false;
    seen.add(block.id);
    return true;
  });
}

function resolveShareScopeFromHie(
  blocks: IPanelBlock[],
  selectedActionIndices?: number[]
): IResolvedShareScope | undefined {
  const taskContext = hybridInteractionEngine.getCurrentTaskContext();
  const artifacts = hybridInteractionEngine.getCurrentArtifacts();
  const tracker = hybridInteractionEngine.getBlockTracker();
  if (!taskContext) return undefined;

  const byId = new Map(blocks.map((block) => [block.id, block]));
  const artifactContext = resolveCurrentArtifactContext(taskContext, artifacts);
  const currentArtifact = artifactContext.currentArtifact;
  const primaryArtifact = artifactContext.primaryArtifact;
  const primaryFromArtifact = primaryArtifact?.blockId ? byId.get(primaryArtifact.blockId) : undefined;
  const primaryFromDerived = taskContext.derivedBlockId ? byId.get(taskContext.derivedBlockId) : undefined;
  const primaryFromSource = taskContext.sourceBlockId ? byId.get(taskContext.sourceBlockId) : undefined;
  const primaryBlock = primaryFromArtifact || primaryFromDerived || primaryFromSource;
  const linkedSourceBlock = (primaryArtifact?.sourceBlockId || currentArtifact?.sourceBlockId)
    ? byId.get(primaryArtifact?.sourceBlockId || currentArtifact?.sourceBlockId || '')
    : undefined;
  if (!primaryBlock || !isShareableBlock(primaryBlock)) {
    return undefined;
  }

  const sourceBlock = linkedSourceBlock && linkedSourceBlock.id !== primaryBlock.id
    ? linkedSourceBlock
    : (primaryFromSource && primaryFromSource.id !== primaryBlock.id ? primaryFromSource : undefined);
  const artifactBlocks = artifactContext.artifactChain
    .map((artifact) => artifact.blockId ? byId.get(artifact.blockId) : undefined)
    .filter((block): block is IPanelBlock => !!block && isShareableBlock(block));
  const selectedFromTask = taskContext.selectedItems
    ?.map((item) => item.index)
    .filter((index): index is number => typeof index === 'number' && index > 0);

  return {
    primaryBlock,
    orderedBlocks: dedupeBlocks([primaryBlock, ...artifactBlocks, ...(sourceBlock ? [sourceBlock] : [])]),
    selectedActionIndices: selectedFromTask && selectedFromTask.length > 0 ? selectedFromTask : selectedActionIndices,
    sourceTurnId: primaryArtifact?.sourceTurnId
      || currentArtifact?.sourceTurnId
      || taskContext.turnId
      || tracker.get(primaryBlock.id)?.turnId
      || (sourceBlock ? tracker.get(sourceBlock.id)?.turnId : undefined),
    sourceRootTurnId: primaryArtifact?.sourceRootTurnId
      || currentArtifact?.sourceRootTurnId
      || taskContext.rootTurnId
      || tracker.get(primaryBlock.id)?.rootTurnId
      || (sourceBlock ? tracker.get(sourceBlock.id)?.rootTurnId : undefined)
  };
}

function resolveVisibleShareScope(
  blocks: IPanelBlock[],
  activeBlockId?: string,
  selectedActionIndices?: number[]
): IResolvedShareScope | undefined {
  const primaryBlock = activeBlockId ? blocks.find((block) => block.id === activeBlockId) : undefined;
  const fallbackBlock = primaryBlock || blocks.slice(-1)[0];
  if (!primaryBlock) {
    if (!fallbackBlock) {
      return undefined;
    }
  }

  const resolvedPrimaryBlock = fallbackBlock!;
  const ordered = [resolvedPrimaryBlock];
  for (let i = blocks.length - 1; i >= 0; i--) {
    const candidate = blocks[i];
    if (candidate.id === resolvedPrimaryBlock.id) continue;
    ordered.push(candidate);
    if (ordered.length >= MAX_DETAILED_BLOCKS) {
      break;
    }
  }

  return {
    primaryBlock: resolvedPrimaryBlock,
    orderedBlocks: ordered,
    selectedActionIndices
  };
}

function resolveShareScope(
  blocks: IPanelBlock[],
  activeBlockId?: string,
  selectedActionIndices?: number[]
): IResolvedShareScope | undefined {
  return resolveShareScopeFromHie(blocks, selectedActionIndices)
    || resolveVisibleShareScope(blocks, activeBlockId, selectedActionIndices);
}

function filterShareableTranscriptEntries(transcript: ITranscriptEntry[]): ITranscriptEntry[] {
  return transcript.filter((entry) => entry.role !== 'system' && entry.text.trim().length > 0);
}

function resolveThreadTranscriptEntries(
  transcript: ITranscriptEntry[],
  turnId?: string,
  rootTurnId?: string
): ITranscriptEntry[] {
  const visibleTranscript = filterShareableTranscriptEntries(transcript);
  if (!turnId && !rootTurnId) {
    return [];
  }

  if (!rootTurnId) {
    return visibleTranscript
      .filter((entry) => entry.turnId === turnId)
      .slice(0, MAX_TRANSCRIPT_ENTRIES);
  }

  const threadEntries = visibleTranscript.filter((entry) => entry.rootTurnId === rootTurnId);
  if (threadEntries.length === 0) {
    return turnId
      ? visibleTranscript.filter((entry) => entry.turnId === turnId).slice(0, MAX_TRANSCRIPT_ENTRIES)
      : [];
  }
  if (!turnId) {
    return threadEntries;
  }

  const sourceTurnEntries = threadEntries.filter((entry) => entry.turnId === turnId);
  if (sourceTurnEntries.length === 0) {
    return threadEntries;
  }

  const sourceTurnEndAt = sourceTurnEntries[sourceTurnEntries.length - 1].timestamp.getTime();
  return threadEntries.filter((entry) => entry.timestamp.getTime() <= sourceTurnEndAt);
}

function buildBlockEmailLines(
  block: IPanelBlock,
  recapService: BlockRecapService,
  selectedActionIndices?: number[],
  selectionBehavior: 'default' | 'strict' = 'default'
): string[] {
  switch (block.type) {
    case 'search-results': {
      const data = block.data as ISearchResultsData;
      const results = selectItemsByIndex(data.results, selectedActionIndices, selectionBehavior);
      const lines = [block.title, `Query: ${data.query}`];
      const searchBreadthLine = formatSearchQueryBreadthLine(data.queryVariants);
      if (searchBreadthLine) {
        lines.push(searchBreadthLine);
      }
      results.forEach((result, index) => {
        lines.push(`${index + 1}. ${formatSearchResult(result).join('\n   ')}`);
      });
      return lines;
    }
    case 'document-library': {
      const data = block.data as IDocumentLibraryData;
      const items = selectItemsByIndex(data.items, selectedActionIndices, selectionBehavior);
      const lines = [block.title];
      if (data.breadcrumb.length > 0) {
        lines.push(`Path: ${data.breadcrumb.join(' / ')}`);
      }
      items.forEach((item, index) => {
        lines.push(`${index + 1}. ${formatDocumentLibraryItem(item).join('\n   ')}`);
      });
      return lines;
    }
    case 'file-preview': {
      const data = block.data as IFilePreviewData;
      return [block.title, `1. ${formatFilePreview(data).join('\n   ')}`];
    }
    case 'site-info': {
      const data = block.data as ISiteInfoData;
      return [block.title, `1. ${formatSiteInfo(data).join('\n   ')}`];
    }
    case 'list-items': {
      const data = block.data as IListItemsData;
      const items = selectItemsByIndex(data.items, selectedActionIndices, selectionBehavior);
      const lines = [block.title, `List: ${data.listName}`];
      items.forEach((item, index) => {
        lines.push(`${index + 1}. ${formatListItem(item, index).join('\n   ')}`);
      });
      return lines;
    }
    case 'info-card': {
      const data = block.data as IInfoCardData;
      return [data.heading || block.title, data.body || block.title];
    }
    default: {
      if (canRecapBlock(block)) {
        const input = recapService.buildRecapInput(block);
        const items = selectItemsByIndex(input.items, selectedActionIndices, selectionBehavior);
        const lines = [block.title];
        if (input.prompt?.trim()) {
          lines.push(`Context: ${input.prompt.trim()}`);
        }
        (input.notes || []).slice(0, 4).forEach((note) => {
          const trimmed = note.trim();
          if (trimmed) {
            lines.push(trimmed);
          }
        });
        items.forEach((item, index) => {
          lines.push(`${index + 1}. ${formatRecapItem(item).join('\n   ')}`);
        });
        return lines;
      }

      return [block.title, `${block.title} (${humanizeBlockType(block.type)})`];
    }
  }
}

function buildDetailedPlainText(
  orderedBlocks: IPanelBlock[],
  recapService: BlockRecapService,
  transcriptEntries: ITranscriptEntry[],
  activeBlockId?: string,
  selectedActionIndices?: number[],
  selectionBehavior: 'default' | 'strict' = 'default'
): { text: string; attachmentUris: string[] } {
  if (orderedBlocks.length === 0) {
    const transcriptLines = transcriptEntries
      .map((entry) => `${entry.role === 'user' ? 'User' : 'Assistant'}: ${clip(entry.text, MAX_TRANSCRIPT_CHARS)}`);
    return {
      text: transcriptLines.join('\n'),
      attachmentUris: []
    };
  }

  const sections: string[] = ['Shared from Grimoire.'];
  const attachmentUris: string[] = [];

  orderedBlocks.forEach((block, index) => {
    const selectedItems = block.id === activeBlockId ? selectedActionIndices : undefined;
    if (block.id === activeBlockId) {
      attachmentUris.push(...collectBlockAttachmentUris(block, selectedItems, selectionBehavior));
    }

    const lines = buildBlockEmailLines(block, recapService, selectedItems, selectionBehavior);
    if (lines.length === 0) return;

    const sectionLabel = index === 0 ? 'Current visible content' : `Additional visible context ${index}`;
    sections.push([sectionLabel, ...lines].join('\n'));
  });

  return {
    text: sections.join('\n\n').trim(),
    attachmentUris: dedupeStrings(attachmentUris)
  };
}

export class SessionShareFormatter {
  private readonly recapService = new BlockRecapService();

  public format(options: ISessionShareOptions): ISessionShareContent {
    const shareableBlocks = options.blocks.filter(isShareableBlock);
    const shareScope = resolveShareScope(shareableBlocks, options.activeBlockId, options.selectedActionIndices);
    const exactTranscriptEntries = resolveThreadTranscriptEntries(
      options.transcript,
      shareScope?.sourceTurnId,
      shareScope?.sourceRootTurnId
    );
    const transcriptEntries = exactTranscriptEntries.length > 0
      ? exactTranscriptEntries
      : filterShareableTranscriptEntries(options.transcript).slice(-MAX_TRANSCRIPT_ENTRIES);
    const subject = shareScope?.primaryBlock
      ? `Grimoire share — ${shareScope.primaryBlock.title}`
      : 'Grimoire share — Session summary';

    const transcriptPlain = transcriptEntries.map((entry) => (
      `${entry.role === 'user' ? 'User' : 'Assistant'}: ${clip(entry.text, MAX_TRANSCRIPT_CHARS)}`
    ));
    const transcriptMarkdown = transcriptEntries.map((entry) => (
      `- **${entry.role === 'user' ? 'User' : 'Assistant'}:** ${escapeMarkdown(clip(entry.text, MAX_TRANSCRIPT_CHARS))}`
    ));

    const blocksForSummary = shareScope?.orderedBlocks && shareScope.orderedBlocks.length > 0
      ? shareScope.orderedBlocks
      : shareableBlocks;
    const blockSummaries = blocksForSummary.map((block) => ({
      title: block.title,
      type: humanizeBlockType(block.type),
      summary: describeBlock(block, this.recapService)
    }));

    const plainTextSections: string[] = [subject];
    if (transcriptPlain.length > 0) {
      plainTextSections.push(['Conversation', ...transcriptPlain].join('\n'));
    }
    if (blockSummaries.length > 0) {
      plainTextSections.push([
        'Visible results',
        ...blockSummaries.map((block, index) => `${index + 1}. ${block.title} (${block.type})\n${block.summary}`)
      ].join('\n\n'));
    }

    const markdownSections: string[] = [`# ${escapeMarkdown(subject)}`];
    if (transcriptMarkdown.length > 0) {
      markdownSections.push(['## Conversation', ...transcriptMarkdown].join('\n'));
    }
    if (blockSummaries.length > 0) {
      markdownSections.push([
        '## Visible Results',
        ...blockSummaries.map((block) => `### ${escapeMarkdown(block.title)}\n\n${escapeMarkdown(block.summary)}`)
      ].join('\n\n'));
    }

    const detailedContent = buildDetailedPlainText(
      shareScope?.orderedBlocks || [],
      this.recapService,
      transcriptEntries,
      shareScope?.primaryBlock?.id,
      shareScope?.selectedActionIndices || options.selectedActionIndices,
      options.selectionBehavior
    );

    return {
      subject,
      plainText: plainTextSections.join('\n\n').trim(),
      markdown: markdownSections.join('\n\n').trim(),
      detailedPlainText: detailedContent.text,
      emailPlainText: detailedContent.text,
      attachmentUris: detailedContent.attachmentUris
    };
  }
}
