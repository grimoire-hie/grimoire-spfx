import type { IPromptRuntimeConfig } from './PromptRuntimeConfig';
import { getSystemPrompt } from './SystemPrompt';

describe('PromptRuntimeConfig overrides', () => {
  afterEach(() => {
    delete window.__GRIMOIRE_PROMPT_CONFIG__;
  });

  it('applies runtime overrides for conversation, visual context, and form composition sections', () => {
    const promptConfig: IPromptRuntimeConfig = {
      conversationGuidelinesOverride: ['Custom conversation rule'],
      visualContextProtocolOverride: ['Custom visual context rule'],
      formCompositionProtocolOverride: ['Custom form composition rule']
    };
    window.__GRIMOIRE_PROMPT_CONFIG__ = promptConfig;

    const prompt = getSystemPrompt('normal', 'classic');

    expect(prompt).toContain('Custom conversation rule');
    expect(prompt).toContain('Custom visual context rule');
    expect(prompt).toContain('Custom form composition rule');
    expect(prompt).not.toContain('1. **Be proactive with tools**');
    expect(prompt).not.toContain('You will receive automatic context updates during the conversation:');
    expect(prompt).not.toContain('When the user asks to WRITE, CREATE, SEND, or COMPOSE something:');
  });
});
