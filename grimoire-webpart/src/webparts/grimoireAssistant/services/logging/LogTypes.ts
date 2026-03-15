/**
 * LogTypes — Log entry interfaces for the logging sidebar
 */

export type LogLevel = 'info' | 'warning' | 'error' | 'debug';
export type LogCategory = 'mcp' | 'llm' | 'search' | 'graph' | 'system' | 'voice';

export interface ILogEntry {
  id: string;
  timestamp: Date;
  level: LogLevel;
  category: LogCategory;
  message: string;
  detail?: string;
  /** Duration in ms (for API calls) */
  durationMs?: number;
  /** HTTP status code */
  statusCode?: number;
  /** Whether the entry is expanded in the UI */
  expanded?: boolean;
}

let _logIdCounter: number = 0;

export function createLogEntry(
  level: LogLevel,
  category: LogCategory,
  message: string,
  detail?: string,
  durationMs?: number,
  statusCode?: number
): ILogEntry {
  _logIdCounter++;
  return {
    id: `log-${_logIdCounter}-${Date.now()}`,
    timestamp: new Date(),
    level,
    category,
    message,
    detail,
    durationMs,
    statusCode
  };
}
