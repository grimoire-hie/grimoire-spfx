import type { WebPartContext } from '@microsoft/sp-webpart-base';

let _context: WebPartContext | undefined;

export const setContext = (context: WebPartContext): void => {
  _context = context;
};

export const getContext = (): WebPartContext => {
  if (!_context) {
    throw new Error('Context not initialized. Call getSP with context first.');
  }
  return _context;
};

export const isInitialized = (): boolean => _context !== undefined;
