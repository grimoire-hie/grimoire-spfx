export function getUserVisibleTextChatError(error: string): string {
  const normalized = error.trim().toLowerCase();

  if (!normalized) {
    return 'I ran into an error while generating the answer. Please try again.';
  }

  if (normalized.includes('timed out')) {
    return 'I hit a timeout while waiting for the answer. Please try again.';
  }

  if (normalized.includes('rate limited') || normalized.includes('429')) {
    return 'I hit a rate limit while generating the answer. Please try again in a few seconds.';
  }

  return 'I ran into an error while generating the answer. Please try again.';
}
