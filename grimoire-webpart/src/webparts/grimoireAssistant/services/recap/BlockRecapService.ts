import type {
  BlockType,
  IActivityFeedData,
  IBlock,
  IChartData,
  IDocumentLibraryData,
  IFilePreviewData,
  IInfoCardData,
  IListItemsData,
  IMarkdownData,
  IPermissionsViewData,
  IProgressTrackerData,
  ISearchResultsData,
  ISelectionListData,
  ISiteInfoData,
  IUserCardData
} from '../../models/IBlock';
import type { IProxyConfig } from '../../store/useGrimoireStore';
import {
  BLOCK_RECAP_RETRY_PROMPT,
  buildBlockRecapSystemPrompt
} from '../../config/promptCatalog';
import { RECAP_LENGTH_LIMITS } from '../../config/assistantLengthLimits';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { logService } from '../logging/LogService';
import { getNanoService } from '../nano/NanoService';
import { formatSearchQueryBreadthLine } from '../search/SearchQueryVariantPresentation';

export interface IBlockRecapItem {
  title: string;
  summary?: string;
  details?: string[];
}

export interface IBlockRecapInput {
  blockId: string;
  blockTitle: string;
  blockType: BlockType;
  prompt?: string;
  notes?: string[];
  items: IBlockRecapItem[];
}

interface ICompactRecapPayloadItem {
  title: string;
  summary?: string;
  details?: string[];
}

interface IRecapPayloadCompactionStep {
  richItemCount: number;
  richSummaryChars: number;
  richDetailCount: number;
  richDetailChars: number;
  compactSummaryChars: number;
  compactDetailCount: number;
  compactDetailChars: number;
}

const MAX_TOTAL_CHARS = RECAP_LENGTH_LIMITS.payloadMaxChars;
const MAX_SUMMARY_CHARS = RECAP_LENGTH_LIMITS.sourceSummaryMaxChars;
const MAX_RECAP_CHARS = RECAP_LENGTH_LIMITS.displayMaxChars;
const RECAP_TOOL_PREFIX = 'block-recap:';
const SEARCH_FILLER_TERMS = new Set([
  'about',
  'document',
  'documents',
  'find',
  'for',
  'info',
  'information',
  'looking',
  'search',
  'searching',
  'show'
]);

function truncate(value: string | undefined, maxChars: number): string {
  if (!value) return '';
  const trimmed = value.trim();
  if (!trimmed) return '';
  if (trimmed.length <= maxChars) return trimmed;
  return `${trimmed.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

function normalizeText(value: string): string {
  return value
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^0-9a-z\u00c0-\u024f\u0400-\u04ff\u0600-\u06ff\u0590-\u05ff\u3040-\u30ff\u4e00-\u9fff\uac00-\ud7af\u0e00-\u0e7f]+/gi, ' ')
    .trim();
}

function tokenizeSearchPrompt(value: string): string[] {
  const normalized = normalizeText(value);
  if (!normalized) return [];
  return Array.from(new Set(
    normalized
      .split(/\s+/)
      .map((term) => term.trim())
      .filter((term) => term.length >= 3 && !SEARCH_FILLER_TERMS.has(term))
  ));
}

function formatList(items: string[]): string {
  if (items.length === 0) return '';
  if (items.length === 1) return items[0];
  if (items.length === 2) return `${items[0]} and ${items[1]}`;
  return `${items.slice(0, -1).join(', ')}, and ${items[items.length - 1]}`;
}

function sentence(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) return '';
  return /[.!?\u3002\uFF01\uFF1F\u061F]$/.test(trimmed) ? trimmed : `${trimmed}.`;
}

function splitIntoSentences(text: string): string[] {
  const compact = text.replace(/\s+/g, ' ').trim();
  if (!compact) return [];
  const matches = compact.match(/[^.!?\u3002\uFF01\uFF1F\u061F]+(?:[.!?\u3002\uFF01\uFF1F\u061F]+|$)/g);
  if (!matches) return [compact];
  return matches
    .map((part) => part.trim())
    .filter(Boolean);
}

function formatRecapForDisplay(text: string): string {
  const trimmed = text.trim();
  if (!trimmed) return '';

  if (/\n\s*[-•]/.test(trimmed)) {
    return trimmed.replace(/\n{3,}/g, '\n\n');
  }

  const sentences = splitIntoSentences(trimmed);
  if (sentences.length <= 1) return trimmed;

  const lead = sentence(sentences[0]);
  const bullets = sentences
    .slice(1)
    .map((line) => `- ${sentence(line)}`);

  return [lead, bullets.join('\n')].filter(Boolean).join('\n\n');
}

function clipRecapText(text: string, maxChars: number = MAX_RECAP_CHARS): string {
  const trimmed = text.trim();
  if (!trimmed || trimmed.length <= maxChars) return trimmed;

  const sentences = splitIntoSentences(trimmed);
  if (sentences.length <= 1) {
    return truncate(trimmed, maxChars);
  }

  const kept: string[] = [];
  for (let i = 0; i < sentences.length; i++) {
    const next = kept.concat(sentence(sentences[i])).join(' ');
    if (next.length > maxChars) break;
    kept.push(sentence(sentences[i]));
  }

  if (kept.length > 0) {
    return kept.join(' ').trim();
  }

  return truncate(trimmed, maxChars);
}

function cleanSnippet(value: string | undefined, maxChars: number): string {
  return truncate(
    (value || '')
      .replace(/\s+/g, ' ')
      .replace(/\s*[-•]\s*/g, ' ')
      .trim(),
    maxChars
  );
}

function buildVisibleSetSentence(items: IBlockRecapItem[], startIndex: number): string {
  const labels = items
    .slice(startIndex)
    .map((item, index) => `#${startIndex + index + 1} "${truncate(item.title, 90)}"`);
  if (labels.length === 0) return '';
  return sentence(`The rest of the visible set includes ${formatList(labels)}`);
}

function buildItemCoverageSentence(
  item: IBlockRecapItem,
  index: number,
  options?: { summaryChars?: number; detailCount?: number; detailChars?: number }
): string {
  const summaryChars = options?.summaryChars ?? 140;
  const detailCount = options?.detailCount ?? 2;
  const detailChars = options?.detailChars ?? 90;
  const summary = cleanSnippet(item.summary, summaryChars);
  const details = (item.details || [])
    .filter((detail) => !/^rank\s*:/i.test(detail))
    .slice(0, detailCount)
    .map((detail) => cleanSnippet(detail, detailChars))
    .filter(Boolean);

  const fragments = [`Result #${index + 1} is "${item.title}"`];
  if (summary) {
    fragments.push(summary);
  }
  if (details.length > 0) {
    fragments.push(details.join(', '));
  }

  return sentence(fragments.join(' — '));
}

function getPayloadCompactionSteps(itemCount: number): IRecapPayloadCompactionStep[] {
  return [
    {
      richItemCount: itemCount,
      richSummaryChars: 220,
      richDetailCount: 2,
      richDetailChars: 90,
      compactSummaryChars: 220,
      compactDetailCount: 2,
      compactDetailChars: 90
    },
    {
      richItemCount: Math.min(8, itemCount),
      richSummaryChars: 200,
      richDetailCount: 2,
      richDetailChars: 90,
      compactSummaryChars: 110,
      compactDetailCount: 1,
      compactDetailChars: 75
    },
    {
      richItemCount: Math.min(5, itemCount),
      richSummaryChars: 180,
      richDetailCount: 2,
      richDetailChars: 80,
      compactSummaryChars: 70,
      compactDetailCount: 0,
      compactDetailChars: 0
    },
    {
      richItemCount: Math.min(3, itemCount),
      richSummaryChars: 140,
      richDetailCount: 1,
      richDetailChars: 70,
      compactSummaryChars: 0,
      compactDetailCount: 0,
      compactDetailChars: 0
    },
    {
      richItemCount: Math.min(2, itemCount),
      richSummaryChars: 100,
      richDetailCount: 0,
      richDetailChars: 0,
      compactSummaryChars: 0,
      compactDetailCount: 0,
      compactDetailChars: 0
    }
  ];
}

function compactPayloadItems(
  items: IBlockRecapItem[],
  step: IRecapPayloadCompactionStep
): ICompactRecapPayloadItem[] {
  return items.map((item, index) => {
    const isRich = index < step.richItemCount;
    const summaryLimit = isRich ? step.richSummaryChars : step.compactSummaryChars;
    const detailCount = isRich ? step.richDetailCount : step.compactDetailCount;
    const detailChars = isRich ? step.richDetailChars : step.compactDetailChars;
    const summary = summaryLimit > 0 ? cleanSnippet(item.summary, summaryLimit) : '';
    const details = detailCount > 0
      ? (item.details || [])
        .slice(0, detailCount)
        .map((detail) => cleanSnippet(detail, detailChars))
        .filter(Boolean)
      : [];
    const titleLimit = isRich ? 180 : (summaryLimit === 0 && detailCount === 0 ? 80 : 120);

    return {
      title: truncate(item.title, titleLimit),
      summary: summary || undefined,
      details: details.length > 0 ? details : undefined
    };
  });
}

function countOverlap(text: string, terms: string[]): number {
  if (terms.length === 0) return 0;
  const normalized = ` ${normalizeText(text)} `;
  let overlap = 0;
  terms.forEach((term) => {
    if (normalized.includes(` ${term} `)) overlap++;
  });
  return overlap;
}

function splitMarkdownItems(content: string): IBlockRecapItem[] {
  const lines = content
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  const numbered = lines.filter((line) => /^\d+[.)]\s+/.test(line));
  if (numbered.length > 0) {
    return numbered.map((line) => ({
      title: line.replace(/^\d+[.)]\s+/, '').replace(/\*\*/g, '').trim()
    }));
  }

  return lines.map((line, index) => ({
    title: index === 0 ? truncate(line.replace(/^#+\s*/, ''), 160) : `Line ${index + 1}`,
    summary: index === 0 ? undefined : truncate(line, 220)
  }));
}

function buildSearchInput(block: IBlock): IBlockRecapInput {
  const data = block.data as ISearchResultsData;
  const uniqueSources = new Set<string>();
  const uniqueLanguages = new Set<string>();

  data.results.forEach((result) => {
    (result.sources || []).forEach((source) => uniqueSources.add(source));
    if (result.language) uniqueLanguages.add(result.language.toUpperCase());
  });

  const notes: string[] = [];
  if (uniqueSources.size > 0) {
    notes.push(`Sources: ${Array.from(uniqueSources).join(', ')}`);
  }
  notes.push(`Visible results: ${data.results.length} of ${data.totalCount || data.results.length}`);
  if (uniqueLanguages.size > 1) {
    notes.push(`Languages: ${Array.from(uniqueLanguages).join(', ')}`);
  }
  const searchBreadthLine = formatSearchQueryBreadthLine(data.queryVariants);
  if (searchBreadthLine) {
    notes.push(searchBreadthLine);
  }

  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.query,
    notes,
    items: data.results.map((result, index) => ({
      title: result.title,
      summary: truncate(result.summary, MAX_SUMMARY_CHARS),
      details: [
        `Rank: ${index + 1}`,
        result.fileType ? `Type: ${result.fileType}` : '',
        result.author ? `Author: ${result.author}` : '',
        result.siteName ? `Site: ${result.siteName}` : '',
        result.language ? `Language: ${result.language.toUpperCase()}` : '',
        result.sources && result.sources.length > 0 ? `Sources: ${result.sources.join(', ')}` : ''
      ].filter(Boolean)
    }))
  };
}

function buildDocumentLibraryInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IDocumentLibraryData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.libraryName || data.siteName,
    notes: data.breadcrumb.length > 0 ? [`Path: ${data.breadcrumb.join(' / ')}`] : undefined,
    items: data.items.map((item) => ({
      title: item.name,
      details: [
        `Kind: ${item.type}`,
        item.fileType ? `Type: ${item.fileType}` : '',
        item.author ? `Author: ${item.author}` : '',
        item.lastModified ? `Modified: ${item.lastModified}` : ''
      ].filter(Boolean)
    }))
  };
}

function buildFilePreviewInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IFilePreviewData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.fileName,
    items: [{
      title: data.fileName,
      details: [
        data.fileType ? `Type: ${data.fileType}` : '',
        data.author ? `Author: ${data.author}` : '',
        data.lastModified ? `Modified: ${data.lastModified}` : '',
        data.size ? `Size: ${data.size}` : ''
      ].filter(Boolean)
    }]
  };
}

function buildSiteInfoInput(block: IBlock): IBlockRecapInput {
  const data = block.data as ISiteInfoData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.siteName,
    items: [{
      title: data.siteName,
      summary: truncate(data.description, MAX_SUMMARY_CHARS),
      details: [
        data.owner ? `Owner: ${data.owner}` : '',
        data.created ? `Created: ${data.created}` : '',
        data.lastModified ? `Modified: ${data.lastModified}` : '',
        data.storageUsed ? `Storage: ${data.storageUsed}` : '',
        data.libraries?.length ? `Libraries: ${data.libraries.length}` : '',
        data.lists?.length ? `Lists: ${data.lists.length}` : ''
      ].filter(Boolean)
    }]
  };
}

function buildUserCardInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IUserCardData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.displayName,
    items: [{
      title: data.displayName || data.email,
      details: [
        data.email ? `Email: ${data.email}` : '',
        data.jobTitle ? `Role: ${data.jobTitle}` : '',
        data.department ? `Department: ${data.department}` : '',
        data.officeLocation ? `Office: ${data.officeLocation}` : '',
        data.phone ? `Phone: ${data.phone}` : ''
      ].filter(Boolean)
    }]
  };
}

function buildListItemsInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IListItemsData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.listName,
    notes: [`Columns: ${data.columns.join(', ')}`],
    items: data.items.map((item, index) => {
      const fields = data.columns
        .map((column) => `${column}: ${item[column] || ''}`.trim())
        .filter((value) => !value.endsWith(':'));
      return {
        title: fields[0] || `Item ${index + 1}`,
        summary: fields.slice(1, 4).join(' | ')
      };
    })
  };
}

function buildSelectionListInput(block: IBlock): IBlockRecapInput {
  const data = block.data as ISelectionListData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.prompt,
    items: data.items.map((item) => ({
      title: item.label,
      summary: truncate(item.description, 220)
    }))
  };
}

function buildPermissionsInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IPermissionsViewData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.targetName,
    items: data.permissions.map((permission) => ({
      title: permission.principal,
      summary: permission.role,
      details: [permission.inherited ? 'Inherited' : 'Unique permission']
    }))
  };
}

function buildActivityInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IActivityFeedData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    items: data.activities.map((activity) => ({
      title: `${activity.actor} ${activity.action} ${activity.target}`.trim(),
      details: [activity.timestamp ? `When: ${activity.timestamp}` : ''].filter(Boolean)
    }))
  };
}

function buildChartInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IChartData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.title,
    notes: [`Chart type: ${data.chartType}`],
    items: data.labels.map((label, index) => ({
      title: label,
      details: [`Value: ${String(data.values[index] ?? '')}`]
    }))
  };
}

function buildInfoCardInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IInfoCardData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.heading,
    items: [{
      title: data.heading,
      summary: truncate(data.body, MAX_SUMMARY_CHARS),
      details: [data.icon ? `Icon: ${data.icon}` : ''].filter(Boolean)
    }]
  };
}

function buildMarkdownInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IMarkdownData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    items: splitMarkdownItems(data.content)
  };
}

function buildProgressInput(block: IBlock): IBlockRecapInput {
  const data = block.data as IProgressTrackerData;
  return {
    blockId: block.id,
    blockTitle: block.title,
    blockType: block.type,
    prompt: data.label,
    items: [{
      title: data.label,
      summary: truncate(data.detail, 220),
      details: [
        `Progress: ${data.progress}%`,
        `Status: ${data.status}`
      ]
    }]
  };
}

function buildRecapInput(block: IBlock): IBlockRecapInput | undefined {
  if (block.originTool?.startsWith(RECAP_TOOL_PREFIX)) return undefined;

  switch (block.type) {
    case 'search-results':
      return buildSearchInput(block);
    case 'document-library':
      return buildDocumentLibraryInput(block);
    case 'file-preview':
      return buildFilePreviewInput(block);
    case 'site-info':
      return buildSiteInfoInput(block);
    case 'user-card':
      return buildUserCardInput(block);
    case 'list-items':
      return buildListItemsInput(block);
    case 'selection-list':
      return buildSelectionListInput(block);
    case 'permissions-view':
      return buildPermissionsInput(block);
    case 'activity-feed':
      return buildActivityInput(block);
    case 'chart':
      return buildChartInput(block);
    case 'info-card':
      return buildInfoCardInput(block);
    case 'markdown':
      return buildMarkdownInput(block);
    case 'progress-tracker':
      return buildProgressInput(block);
    default:
      return undefined;
  }
}

function buildCompactPayload(input: IBlockRecapInput): string {
  const notes = (input.notes || []).map((note) => truncate(note, 180));
  const steps = getPayloadCompactionSteps(input.items.length);
  let serialized = '';

  for (let i = 0; i < steps.length; i++) {
    const step = steps[i];
    const compactedStart = step.richItemCount < input.items.length ? step.richItemCount + 1 : undefined;
    const payload = {
      blockTitle: truncate(input.blockTitle, 180),
      blockType: input.blockType,
      prompt: truncate(input.prompt, 180),
      visibleItemCount: input.items.length,
      notes: [
        ...notes,
        compactedStart
          ? `Coverage note: All ${input.items.length} visible items are included. Items ${compactedStart}-${input.items.length} are represented more compactly than the top-ranked items to stay within the recap budget.`
          : ''
      ].filter(Boolean),
      items: compactPayloadItems(input.items, step)
    };

    serialized = JSON.stringify(payload);
    if (serialized.length <= MAX_TOTAL_CHARS) {
      return serialized;
    }
  }

  return serialized;
}

function buildRetryPayload(input: IBlockRecapInput): string {
  const lines: string[] = [
    `Block type: ${input.blockType}`,
    `Block title: ${truncate(input.blockTitle, 180)}`,
    `Visible items: ${input.items.length}`
  ];

  if (input.prompt) {
    lines.push(`User intent: ${truncate(input.prompt, 220)}`);
  }

  (input.notes || []).slice(0, 6).forEach((note) => {
    lines.push(`Note: ${truncate(note, 220)}`);
  });

  if (input.items.length > 8) {
    lines.push(`Coverage note: All ${input.items.length} visible items are included. Items 9-${input.items.length} are represented more compactly than the strongest items to stay within the recap budget.`);
  }

  input.items.forEach((item, index) => {
    const isRich = index < 8;
    const parts = [truncate(item.title, 180)];
    if (item.summary) parts.push(truncate(item.summary, isRich ? 220 : 90));
    if (item.details && item.details.length > 0) {
      const detailCount = isRich ? 2 : 0;
      if (detailCount > 0) {
        parts.push(item.details.slice(0, detailCount).map((detail) => truncate(detail, 120)).join(' | '));
      }
    }
    lines.push(`${index + 1}. ${parts.filter(Boolean).join(' — ')}`);
  });

  return lines.join('\n').slice(0, MAX_TOTAL_CHARS);
}

function toDisplaySearchTopic(query: string): string {
  const terms = tokenizeSearchPrompt(query);
  if (terms.length === 0) return query.trim();
  if (terms.length === 1) {
    const raw = query.trim();
    if (raw && raw.replace(/[^a-z0-9]/gi, '').toLowerCase() === terms[0]) {
      return raw;
    }
    return terms[0];
  }
  return terms.join(' ');
}

function normalizeTopicKey(value: string): string {
  return normalizeText(
    value
      .replace(/[_-](en|de|fr|it|es|ja|zh|ko|nl|pl|pt|ru|sv|tr|uk)$/i, '')
      .replace(/\s+\((en|de|fr|it|es|ja|zh|ko|nl|pl|pt|ru|sv|tr|uk)\)$/i, '')
  );
}

function buildSearchFallback(input: IBlockRecapInput): string {
  const parts: string[] = [];
  const query = input.prompt || input.blockTitle;
  const topic = toDisplaySearchTopic(query);
  const queryTerms = tokenizeSearchPrompt(query);
  const topItems = input.items.slice(0, 5);
  const topTitles = topItems.map((item) => item.title).filter(Boolean);
  const strongestTitles = queryTerms.length > 0
    ? topItems
      .filter((item) => countOverlap(`${item.title} ${item.summary || ''}`, queryTerms) > 0)
      .map((item) => item.title)
      .filter(Boolean)
    : [];
  const headlineTitles = strongestTitles.length >= 2 ? strongestTitles : topTitles;
  const quotedHeadlineTitles = headlineTitles.map((title) => `"${title}"`);

  parts.push(sentence(`This recap covers all ${input.items.length} visible results${topic ? ` for "${topic}"` : ''}`));

  if (topTitles.length > 0 && quotedHeadlineTitles.length > 0) {
    const groupedTitles = new Map<string, number>();
    headlineTitles.forEach((title) => {
      const key = normalizeTopicKey(title);
      if (!key) return;
      groupedTitles.set(key, (groupedTitles.get(key) || 0) + 1);
    });
    const dominantTitleGroup = Array.from(groupedTitles.entries()).sort((a, b) => b[1] - a[1])[0];
    if (dominantTitleGroup && dominantTitleGroup[1] >= 2) {
      parts.push(sentence(`The visible results are tightly focused on "${topic || topTitles[0]}"`));
      parts.push(sentence(`The strongest hits are ${formatList(quotedHeadlineTitles)}, which look like language or format variants of the same topic`));
    } else {
      parts.push(sentence(`The visible results are focused on ${topic ? `"${topic}"` : 'the requested topic'}`));
      parts.push(sentence(`The strongest hits are ${formatList(quotedHeadlineTitles)}`));
    }
  }

  const languagesNote = (input.notes || []).find((note) => note.startsWith('Languages: '));
  if (languagesNote) {
    parts.push(sentence(`The visible results span ${languagesNote.replace(/^Languages:\s*/i, '')}`));
  }

  topItems.forEach((item, index) => {
    parts.push(buildItemCoverageSentence(item, index, { summaryChars: 130, detailCount: 2, detailChars: 80 }));
  });

  if (input.items.length > topItems.length) {
    parts.push(buildVisibleSetSentence(input.items, topItems.length));
  }

  if (queryTerms.length > 0 && input.items.length > 3) {
    const overlapScores = input.items.map((item) => countOverlap(`${item.title} ${item.summary || ''}`, queryTerms));
    const topHasSignal = overlapScores.slice(0, 3).some((score) => score > 0);
    const tailHasNoSignal = overlapScores.slice(3).some((score) => score === 0);
    if (topHasSignal && tailHasNoSignal) {
      parts.push('Some lower-ranked items look broader than the strongest matches.');
    }
  }

  parts.push('This recap is based only on the titles, snippets, and metadata currently visible in the panel, so it reflects likely themes and differences rather than hidden full-document details.');

  return clipRecapText(parts.join(' ').trim());
}

function buildGenericFallback(input: IBlockRecapInput): string {
  const parts: string[] = [];
  const topTitles = input.items.slice(0, 5).map((item) => item.title).filter(Boolean);
  const label = input.prompt || input.blockTitle;

  parts.push(sentence(`This recap covers all ${input.items.length} visible ${input.blockType.replace(/-/g, ' ')} entr${input.items.length === 1 ? 'y' : 'ies'} related to "${label}"`));

  if (topTitles.length > 0) {
    parts.push(sentence(`The strongest visible items are ${formatList(topTitles.map((title) => `"${title}"`))}`));
  }

  input.items.slice(0, 5).forEach((item, index) => {
    parts.push(buildItemCoverageSentence(item, index, { summaryChars: 120, detailCount: 2, detailChars: 80 }));
  });

  if (input.items.length > 5) {
    parts.push(buildVisibleSetSentence(input.items, 5));
  }

  if (input.notes && input.notes.length > 0) {
    input.notes.slice(0, 3).forEach((note) => {
      parts.push(note.endsWith('.') ? note : `${note}.`);
    });
  }

  parts.push('This fallback recap uses only the visible data in the current panel state, so it summarizes what is shown without inventing hidden details.');

  return clipRecapText(parts.join(' ').trim());
}

export function canRecapBlock(block: IBlock | undefined): boolean {
  if (!block || block.originTool?.startsWith(RECAP_TOOL_PREFIX)) return false;
  const input = buildRecapInput(block);
  return !!input && input.items.length > 0;
}

export function getRecapOriginTool(sourceBlockId: string): string {
  return `${RECAP_TOOL_PREFIX}${sourceBlockId}`;
}

export class BlockRecapService {
  public buildRecapInput(block: IBlock): IBlockRecapInput {
    const input = buildRecapInput(block);
    if (!input || input.items.length === 0) {
      throw new Error('This block does not have enough visible data to recap.');
    }
    return input;
  }

  public buildFallbackRecap(input: IBlockRecapInput): string {
    if (input.blockType === 'search-results') {
      return buildSearchFallback(input);
    }
    return buildGenericFallback(input);
  }

  public async generate(block: IBlock, proxyConfig?: IProxyConfig): Promise<string> {
    const input = this.buildRecapInput(block);
    const fallback = formatRecapForDisplay(this.buildFallbackRecap(input));
    const nano = getNanoService(proxyConfig);
    if (!nano) {
      logService.info('search', `Block recap using fallback for ${block.type}: no fast model available`);
      return fallback;
    }

    try {
      const tuning = getRuntimeTuningConfig().nano;
      const content = await nano.classify(
        buildBlockRecapSystemPrompt(input.blockType),
        buildCompactPayload(input),
        tuning.blockRecapTimeoutMs,
        tuning.blockRecapMaxTokens
      );

      let text = content?.trim();
      if (!text) {
        logService.warning('search', 'Block recap backend returned empty content; retrying with compact recap prompt');
        const retryContent = await nano.classify(
          BLOCK_RECAP_RETRY_PROMPT,
          buildRetryPayload(input),
          tuning.blockRecapTimeoutMs,
          Math.max(
            tuning.blockRecapMaxTokens + tuning.blockRecapRetryHeadroomTokens,
            tuning.blockRecapRetryMinTokens
          )
        );
        text = retryContent?.trim();
      }

      if (!text) {
        logService.warning('search', 'Block recap backend returned empty content twice; using fallback recap');
        return fallback;
      }

      logService.info('search', `Block recap generated via fast model for ${block.type}`);
      return formatRecapForDisplay(text);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Recap generation failed';
      logService.warning('search', `Block recap fallback: ${message}`);
      return fallback;
    }
  }
}
