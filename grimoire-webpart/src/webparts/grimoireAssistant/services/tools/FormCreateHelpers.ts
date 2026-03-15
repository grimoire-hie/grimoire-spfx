import type { FormPresetId, IBlock, IInfoCardData, IMarkdownData } from '../../models/IBlock';
import { BlockRecapService, canRecapBlock } from '../recap/BlockRecapService';

const INVALID_SHAREPOINT_FILE_NAME_CHARS = /["*:<>?/\\|]/g;

export function looksLikeWordDocumentCreateRequest(
  requestedPreset: FormPresetId,
  title: string,
  description: string | undefined,
  staticArgs: Record<string, unknown>,
  prefill: Record<string, string>,
  conversationHints: string[]
): boolean {
  if (requestedPreset !== 'file-create') {
    return false;
  }

  const cueText = [
    title,
    description,
    ...conversationHints,
    typeof staticArgs.fileName === 'string' ? staticArgs.fileName : undefined,
    typeof staticArgs.filename === 'string' ? staticArgs.filename : undefined,
    prefill.fileName,
    prefill.filename
  ]
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .join('\n');

  if (!cueText.trim()) {
    return false;
  }

  if (/\btext file\b/i.test(cueText) || /\.txt\b/i.test(cueText)) {
    return false;
  }

  return /\bword document\b/i.test(cueText)
    || /\bdocx\b/i.test(cueText)
    || /\.docx\b/i.test(cueText)
    || /\bdocument named\b/i.test(cueText)
    || /\bdocument called\b/i.test(cueText)
    || /\bcreate\s+(?:a\s+|an\s+|new\s+)?document\b/i.test(cueText);
}

function sanitizeContextualFileStem(value: string | undefined): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) {
    return undefined;
  }

  const withoutControlChars = Array.from(trimmed)
    .map((character) => character < ' ' ? ' ' : character)
    .join('');

  const normalized = withoutControlChars
    .replace(/^create\s+(?:a\s+|an\s+)?(?:new\s+)?(?:word\s+document|document|text\s+file|file)\s*:?\s*/i, '')
    .replace(/^(word\s+document|document|text\s+file|file)\s*:?\s*/i, '')
    .replace(/^summary\s*:\s*/i, '')
    .replace(/^recap\s*:\s*/i, 'Recap - ')
    .replace(/\bsearch\s*:\s*/gi, '')
    .replace(/[<>:"/\\|?*]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\s*-\s*/g, ' - ')
    .replace(/^[.\s-]+|[.\s-]+$/g, '')
    .trim();

  return normalized || undefined;
}

export function sanitizeCreatedFileName(
  fileName: string | undefined,
  extension: string
): string {
  const defaultStem = extension.toLowerCase() === 'docx' ? 'Document' : 'document';
  const trimmed = fileName?.trim();
  const fallbackName = `${defaultStem}.${extension}`;
  const candidate = trimmed || fallbackName;
  const extensionMatch = candidate.match(/\.([a-z0-9]{2,8})$/i);
  const resolvedExtension = extensionMatch?.[1]?.toLowerCase() || extension.toLowerCase();
  const rawStem = extensionMatch
    ? candidate.slice(0, Math.max(0, candidate.length - extensionMatch[0].length))
    : candidate;
  const withoutControlChars = Array.from(rawStem)
    .map((character) => character < ' ' ? ' ' : character)
    .join('');
  const sanitizedStem = withoutControlChars
    .replace(INVALID_SHAREPOINT_FILE_NAME_CHARS, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^[.\s]+|[.\s]+$/g, '')
    .trim();
  const stem = sanitizedStem || defaultStem;
  return `${stem}.${resolvedExtension}`;
}

function ensureFileExtension(fileName: string, extension: string): string {
  const trimmed = fileName.trim();
  if (!trimmed) {
    return `Document.${extension}`;
  }

  return /\.[a-z0-9]{2,8}$/i.test(trimmed)
    ? trimmed
    : `${trimmed}.${extension}`;
}

export function buildContextualFileName(
  preset: FormPresetId,
  title: string,
  activeBlockTitle?: string
): string {
  const extension = preset === 'word-document-create' ? 'docx' : 'txt';
  const candidate = sanitizeContextualFileStem(title)
    || sanitizeContextualFileStem(activeBlockTitle);
  if (!candidate) {
    return sanitizeCreatedFileName(undefined, extension);
  }

  return sanitizeCreatedFileName(ensureFileExtension(candidate, extension), extension);
}

export function extractFocusedBlockContent(
  blocks: IBlock[],
  activeBlockId: string | undefined,
  pickResolvedString: (...values: unknown[]) => string | undefined
): string | undefined {
  if (!activeBlockId) {
    return undefined;
  }

  const block = blocks.find((candidate) => candidate.id === activeBlockId);
  if (!block) {
    return undefined;
  }

  if (block.type === 'info-card') {
    const data = block.data as IInfoCardData;
    return pickResolvedString(data.body, data.heading, block.title);
  }

  if (block.type === 'markdown') {
    const data = block.data as IMarkdownData;
    return pickResolvedString(data.content, block.title);
  }

  if (canRecapBlock(block)) {
    const recapService = new BlockRecapService();
    return recapService.buildFallbackRecap(recapService.buildRecapInput(block));
  }

  return undefined;
}
