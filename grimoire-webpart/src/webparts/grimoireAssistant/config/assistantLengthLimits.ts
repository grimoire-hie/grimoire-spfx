export interface ICharacterRange {
  min: number;
  max: number;
}

export const ASSISTANT_SUMMARY_TARGET_CHARS: ICharacterRange = {
  min: 800,
  max: 1600
};

export const RECAP_LENGTH_LIMITS = {
  payloadMaxChars: 12_000,
  sourceSummaryMaxChars: 380,
  displayMaxChars: 1600
} as const;

export const TOOL_CONTEXT_LENGTH_LIMITS = {
  llmInputMaxChars: 30_000,
  toolResultMaxChars: 2_000,
  readToolContentMaxChars: 14_000
} as const;

export const MARKDOWN_LENGTH_LIMITS = {
  renderMaxChars: 8_000
} as const;

export const TRANSCRIPT_OVERLAY_LENGTH_LIMITS = {
  textMaxChars: 320
} as const;

export const SHARE_LENGTH_LIMITS = {
  transcriptMaxChars: 420,
  blockSummaryMaxChars: 900
} as const;

export function formatCharacterRange(range: ICharacterRange): string {
  return `${range.min}-${range.max} characters`;
}

export const ASSISTANT_SUMMARY_TARGET_TEXT = formatCharacterRange(ASSISTANT_SUMMARY_TARGET_CHARS);
