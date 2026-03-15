export interface INormalizedError {
  name: string;
  message: string;
}

export function normalizeError(error: unknown, fallbackMessage: string): INormalizedError {
  if (error instanceof Error) {
    return {
      name: error.name || 'Error',
      message: error.message || fallbackMessage
    };
  }

  if (typeof error === 'string') {
    return {
      name: 'Error',
      message: error || fallbackMessage
    };
  }

  if (error && typeof error === 'object') {
    const candidate = error as { name?: unknown; message?: unknown };
    return {
      name: typeof candidate.name === 'string' && candidate.name ? candidate.name : 'Error',
      message: typeof candidate.message === 'string' && candidate.message ? candidate.message : fallbackMessage
    };
  }

  return {
    name: 'Error',
    message: fallbackMessage
  };
}
