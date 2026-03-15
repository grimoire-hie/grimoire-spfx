import {
  ASSISTANT_SUMMARY_TARGET_CHARS,
  ASSISTANT_SUMMARY_TARGET_TEXT,
  MARKDOWN_LENGTH_LIMITS,
  RECAP_LENGTH_LIMITS,
  SHARE_LENGTH_LIMITS,
  TOOL_CONTEXT_LENGTH_LIMITS,
  TRANSCRIPT_OVERLAY_LENGTH_LIMITS,
  formatCharacterRange
} from './assistantLengthLimits';

describe('assistantLengthLimits', () => {
  it('exposes the shared summary target text once', () => {
    expect(ASSISTANT_SUMMARY_TARGET_CHARS).toEqual({ min: 800, max: 1600 });
    expect(ASSISTANT_SUMMARY_TARGET_TEXT).toBe('800-1600 characters');
    expect(formatCharacterRange(ASSISTANT_SUMMARY_TARGET_CHARS)).toBe(ASSISTANT_SUMMARY_TARGET_TEXT);
  });

  it('groups the shared character-based limits by subsystem', () => {
    expect(RECAP_LENGTH_LIMITS).toEqual({
      payloadMaxChars: 12000,
      sourceSummaryMaxChars: 380,
      displayMaxChars: 1600
    });
    expect(TOOL_CONTEXT_LENGTH_LIMITS).toEqual({
      llmInputMaxChars: 30000,
      toolResultMaxChars: 2000,
      readToolContentMaxChars: 14000
    });
    expect(MARKDOWN_LENGTH_LIMITS.renderMaxChars).toBe(8000);
    expect(TRANSCRIPT_OVERLAY_LENGTH_LIMITS.textMaxChars).toBe(320);
    expect(SHARE_LENGTH_LIMITS).toEqual({
      transcriptMaxChars: 420,
      blockSummaryMaxChars: 900
    });
  });
});
