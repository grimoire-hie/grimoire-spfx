import {
  looksLikeBinary,
  MAX_LLM_CHARS,
  parseRenderHints,
  TEXT_READABLE_EXTENSIONS,
  WORD_EXTENSIONS
} from './ToolRuntimeContentHelpers';
import { TOOL_CONTEXT_LENGTH_LIMITS } from '../../config/assistantLengthLimits';

describe('ToolRuntimeContentHelpers', () => {
  it('parses valid render hints JSON', () => {
    const hints = parseRenderHints({ render_hints: JSON.stringify({ compact: true, priority: 'high' }) });
    expect(hints).toEqual({ compact: true, priority: 'high' });
  });

  it('returns undefined for invalid render hints JSON', () => {
    expect(parseRenderHints({ render_hints: '{bad-json' })).toBeUndefined();
    expect(parseRenderHints({})).toBeUndefined();
  });

  it('exposes expected content constants', () => {
    expect(MAX_LLM_CHARS).toBe(TOOL_CONTEXT_LENGTH_LIMITS.llmInputMaxChars);
    expect(TEXT_READABLE_EXTENSIONS.has('txt')).toBe(true);
    expect(WORD_EXTENSIONS.has('docx')).toBe(true);
  });

  it('detects binary-like payloads', () => {
    expect(looksLikeBinary('PK\u0003\u0004binary')).toBe(true);
    expect(looksLikeBinary('abc\u0001\u0002\u0003\u0004\u0005\u0006')).toBe(true);
    expect(looksLikeBinary('plain text content')).toBe(false);
  });
});
