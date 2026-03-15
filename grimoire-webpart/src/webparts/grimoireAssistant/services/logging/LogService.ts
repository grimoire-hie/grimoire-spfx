/**
 * LogService
 * Centralized log aggregator singleton.
 * All services call LogService to emit log entries → pushed to Zustand store.
 */

import { createLogEntry, LogLevel, LogCategory } from './LogTypes';

type LogHandler = (entry: ReturnType<typeof createLogEntry>) => void;

class LogServiceImpl {
  private handler: LogHandler | undefined;

  /**
   * Wire the log service to the Zustand store's addLogEntry action.
   */
  public setHandler(handler: LogHandler): void {
    this.handler = handler;
  }

  public info(category: LogCategory, message: string, detail?: string, durationMs?: number): void {
    this.emit('info', category, message, detail, durationMs);
  }

  public warning(category: LogCategory, message: string, detail?: string): void {
    this.emit('warning', category, message, detail);
  }

  public error(category: LogCategory, message: string, detail?: string, durationMs?: number): void {
    this.emit('error', category, message, detail, durationMs);
  }

  public debug(category: LogCategory, message: string, detail?: string): void {
    this.emit('debug', category, message, detail);
  }

  private emit(
    level: LogLevel,
    category: LogCategory,
    message: string,
    detail?: string,
    durationMs?: number
  ): void {
    const entry = createLogEntry(level, category, message, detail, durationMs);

    if (this.handler) {
      this.handler(entry);
    }

    // Also emit to browser console for debugging
    const prefix = `[Grimoire:${category}]`;
    switch (level) {
      case 'error':
        console.error(prefix, message, detail || '');
        break;
      case 'warning':
        console.warn(prefix, message, detail || '');
        break;
      case 'debug':
        console.debug(prefix, message, detail || '');
        break;
      default:
        console.log(prefix, message, detail || '');
    }
  }
}

export const logService = new LogServiceImpl();
