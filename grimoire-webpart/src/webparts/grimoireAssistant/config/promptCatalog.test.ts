import {
  BLOCK_RECAP_RETRY_PROMPT,
  buildBlockRecapSystemPrompt,
  getDefaultConversationGuidelineLines
} from './promptCatalog';
import { ASSISTANT_SUMMARY_TARGET_TEXT } from './assistantLengthLimits';

describe('promptCatalog recap and summary targets', () => {
  it('targets the shared character range for recap prompts', () => {
    expect(BLOCK_RECAP_RETRY_PROMPT).toContain(ASSISTANT_SUMMARY_TARGET_TEXT);
    expect(BLOCK_RECAP_RETRY_PROMPT).not.toContain('3-5 short sentences');

    const prompt = buildBlockRecapSystemPrompt('search-results');
    expect(prompt).toContain(ASSISTANT_SUMMARY_TARGET_TEXT);
    expect(prompt).toContain('Cover the full visible set');
    expect(prompt).not.toContain('2-4 bullet lines');
  });

  it('targets the shared character range in the shared summarize guidance', () => {
    const lines = getDefaultConversationGuidelineLines('Keep the user phrasing.');
    const combined = lines.join('\n');

    expect(combined).toContain(`target ${ASSISTANT_SUMMARY_TARGET_TEXT}`);
    expect(combined).not.toContain('target 700-1200 characters');
  });
});
