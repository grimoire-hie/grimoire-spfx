import type { BlockType } from '../../models/IBlock';
import { isExplicitSelectionListRequest } from './ToolAcknowledgment';

const ACTIONABLE_LIST_BLOCKS: ReadonlySet<BlockType> = new Set<BlockType>([
  'search-results',
  'markdown',
  'list-items',
  'document-library'
]);

const GENERIC_SELECTION_PROMPT_PATTERNS: ReadonlyArray<RegExp> = [
  /\bpick one\b/i,
  /\bwhich one\b/i,
  /\bopen\b/i,
  /\bpreview\b/i,
  /\bsummarize\b/i,
  /\bchoose (one|an item)\b/i,
  /\bselect (one|an item)\b/i
];

function isGenericSelectionPrompt(prompt: string): boolean {
  if (!prompt.trim()) return false;
  return GENERIC_SELECTION_PROMPT_PATTERNS.some((pattern) => pattern.test(prompt));
}

export function shouldSuppressSelectionList(
  prompt: string,
  latestBlockType: BlockType | undefined,
  latestUserUtterance: string
): boolean {
  if (!latestBlockType || !ACTIONABLE_LIST_BLOCKS.has(latestBlockType)) {
    return false;
  }
  if (!isGenericSelectionPrompt(prompt)) {
    return false;
  }
  if (isExplicitSelectionListRequest(latestUserUtterance)) {
    return false;
  }
  return true;
}
