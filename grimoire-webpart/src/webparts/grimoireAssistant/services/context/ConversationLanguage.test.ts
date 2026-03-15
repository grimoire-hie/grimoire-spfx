import {
  buildConversationLanguageContextMessage,
  detectExplicitConversationLanguageSwitch,
  resolveConversationLanguage
} from './ConversationLanguage';

describe('ConversationLanguage', () => {
  it('detects explicit language switches', () => {
    expect(detectExplicitConversationLanguageSwitch('parliamo in italiano')).toBe('it');
    expect(detectExplicitConversationLanguageSwitch('please answer in french')).toBe('fr');
  });

  it('keeps the current non-English session language for ambiguous technical turns', () => {
    expect(resolveConversationLanguage('search SPFX', 'it', 'en')).toBe('it');
  });

  it('builds a silent context reminder for the realtime session', () => {
    expect(buildConversationLanguageContextMessage('de'))
      .toBe('[Conversation preference: Reply in German until the user explicitly switches language.]');
  });
});
