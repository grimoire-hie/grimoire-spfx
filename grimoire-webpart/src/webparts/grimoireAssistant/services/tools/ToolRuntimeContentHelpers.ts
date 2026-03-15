import type { IRenderHints } from '../../models/IBlock';
import { TOOL_CONTEXT_LENGTH_LIMITS } from '../../config/assistantLengthLimits';

export function parseRenderHints(args: Record<string, unknown>): IRenderHints | undefined {
  const raw = args.render_hints as string | undefined;
  if (!raw) return undefined;
  try {
    return JSON.parse(raw) as IRenderHints;
  } catch {
    return undefined;
  }
}

export const TEXT_READABLE_EXTENSIONS = new Set([
  'txt', 'csv', 'json', 'xml', 'html', 'htm', 'css', 'js', 'ts',
  'md', 'yaml', 'yml', 'log', 'ini', 'cfg', 'conf', 'sh', 'bat',
  'ps1', 'py', 'java', 'cs', 'cpp', 'c', 'h', 'sql', 'r', 'jsx', 'tsx'
]);

export const WORD_EXTENSIONS = new Set(['docx']);
export const MAX_LLM_CHARS = TOOL_CONTEXT_LENGTH_LIMITS.llmInputMaxChars;

export function looksLikeBinary(text: string): boolean {
  if (text.startsWith('PK')) return true;
  let nonPrint = 0;
  const sample = text.substring(0, 200);
  for (let i = 0; i < sample.length; i++) {
    const code = sample.charCodeAt(i);
    if (code < 32 && code !== 9 && code !== 10 && code !== 13) nonPrint++;
  }
  return nonPrint > sample.length * 0.1;
}
