import {
  getToolCompletionAckFromOutput,
  getToolCompletionAckText,
  getVoiceToolCompletionAckText,
  shouldUseImmediateCompletionAck
} from './ToolCompletionAcknowledgment';

describe('ToolCompletionAcknowledgment', () => {
  it('returns SharePoint search completion text with item count', () => {
    expect(getToolCompletionAckText('search_sharepoint', 5))
      .toBe('I found 5 documents. They are in the panel.');
  });

  it('localizes completion text when the session language is non-English', () => {
    expect(getToolCompletionAckText('search_sharepoint', 3, 'it'))
      .toBe('Ho trovato 3 documenti. Sono nel pannello.');
  });

  it('returns zero-result completion text', () => {
    expect(getToolCompletionAckText('search_people', 0))
      .toBe('I did not find matching people.');
  });

  it('returns undefined for unsupported tools', () => {
    expect(getToolCompletionAckText('set_expression', 1)).toBeUndefined();
  });

  it('derives completion text from async tool output payload', () => {
    expect(getToolCompletionAckFromOutput(
      'search_sharepoint',
      JSON.stringify({ success: true, displayedResults: 5 })
    )).toBe('I found 5 documents. They are in the panel.');
  });

  it('allows immediate completion ack for simple search requests', () => {
    expect(shouldUseImmediateCompletionAck('search_sharepoint', 'i am searching for info about animals')).toBe(true);
  });

  it('blocks immediate completion ack for chained search requests', () => {
    expect(shouldUseImmediateCompletionAck('search_sharepoint', 'find documents about animals and summarize the first one')).toBe(false);
  });

  it('allows the generic capability overview ack', () => {
    expect(shouldUseImmediateCompletionAck('list_m365_servers', 'what do you offer?')).toBe(true);
  });

  it('blocks the generic capability overview ack for focused workload queries', () => {
    expect(shouldUseImmediateCompletionAck('list_m365_servers', 'what can you do for sharepoint?')).toBe(false);
  });

  it('returns a deterministic voice completion ack for successful compose forms', () => {
    expect(getVoiceToolCompletionAckText(
      'show_compose_form',
      { preset: 'email-compose' },
      JSON.stringify({ success: true, message: 'Form displayed.' })
    )).toBe('I opened the email draft in the panel.');
  });

  it('reuses the email draft completion ack for reply-all thread compose forms', () => {
    expect(getVoiceToolCompletionAckText(
      'show_compose_form',
      { preset: 'email-reply-all-thread' },
      JSON.stringify({ success: true, message: 'Form displayed.' })
    )).toBe('I opened the email draft in the panel.');
  });

  it('localizes compose completion ack for successful voice forms', () => {
    expect(getVoiceToolCompletionAckText(
      'show_compose_form',
      { preset: 'email-compose' },
      JSON.stringify({ success: true, message: 'Form displayed.' }),
      'fr'
    )).toBe('J’ai ouvert le brouillon d’e-mail dans le panneau.');
  });

  it('does not emit voice completion ack for failed compose forms', () => {
    expect(getVoiceToolCompletionAckText(
      'show_compose_form',
      { preset: 'email-compose' },
      JSON.stringify({ success: false, error: 'No form.' })
    )).toBeUndefined();
  });
});
