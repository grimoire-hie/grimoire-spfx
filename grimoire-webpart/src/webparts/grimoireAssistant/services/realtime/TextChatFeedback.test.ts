import { getUserVisibleTextChatError } from './TextChatFeedback';

describe('TextChatFeedback', () => {
  it('returns a timeout-specific message', () => {
    expect(getUserVisibleTextChatError('Request timed out'))
      .toBe('I hit a timeout while waiting for the answer. Please try again.');
  });

  it('returns a rate-limit-specific message', () => {
    expect(getUserVisibleTextChatError('Chat completion failed: rate limited after all retries'))
      .toBe('I hit a rate limit while generating the answer. Please try again in a few seconds.');
  });

  it('falls back to a generic error message', () => {
    expect(getUserVisibleTextChatError('socket hang up'))
      .toBe('I ran into an error while generating the answer. Please try again.');
  });
});
