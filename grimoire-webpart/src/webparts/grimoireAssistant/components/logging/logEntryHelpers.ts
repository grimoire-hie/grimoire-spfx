import type { ILogEntry } from '../../services/logging/LogTypes';
import type { IMcpExecutionTrace } from '../../services/mcp/McpExecutionAdapter';

export function hasLogEntryDetails(entry: ILogEntry): boolean {
  return typeof entry.detail === 'string' && entry.detail.trim().length > 0;
}

export function getLogEntryToggleLabel(expanded: boolean): string {
  return expanded ? 'Hide details' : 'Show details';
}

export interface IParsedMcpExecutionTrace {
  trace: IMcpExecutionTrace;
  toolLabel: string;
  serverLabel?: string;
  targetLabel?: string;
  targetSourceLabel?: string;
  resultLabel?: string;
  requiredLabel?: string;
  rawArgsLabel?: string;
  normalizedArgsLabel?: string;
  resolvedArgsLabel?: string;
  unwrapLabel?: string;
  recoveryLabel?: string;
}

function getTargetSourceLabel(source: IMcpExecutionTrace['targetSource']): string | undefined {
  switch (source) {
    case 'explicit-user':
      return 'Explicit user target';
    case 'hie-selection':
      return 'HIE selection';
    case 'current-page':
      return 'Current page fallback';
    case 'recovered':
      return 'Recovered target';
    default:
      return undefined;
  }
}

function formatCompactJson(value: unknown, maxChars = 220): string | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return undefined;
  }

  const serialized = JSON.stringify(value);
  if (!serialized) {
    return undefined;
  }

  return serialized.length <= maxChars
    ? serialized
    : `${serialized.slice(0, Math.max(0, maxChars - 3)).trimEnd()}...`;
}

function tryParseJsonRecord(detail?: string): Record<string, unknown> | undefined {
  if (typeof detail !== 'string' || !detail.trim()) {
    return undefined;
  }

  try {
    const parsed = JSON.parse(detail);
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : undefined;
  } catch {
    return undefined;
  }
}

export function parseMcpExecutionTrace(entry: ILogEntry): IParsedMcpExecutionTrace | undefined {
  if (
    entry.category !== 'mcp'
    || (entry.message !== 'MCP execution trace' && entry.message !== 'Share execution trace')
  ) {
    return undefined;
  }

  const record = tryParseJsonRecord(entry.detail);
  if (!record || typeof record.toolName !== 'string') {
    return undefined;
  }

  const trace = record as unknown as IMcpExecutionTrace;
  const recoverySteps = Array.isArray(trace.recoverySteps) ? trace.recoverySteps : [];
  const targetSourceLabel = getTargetSourceLabel(trace.targetSource);
  return {
    trace,
    toolLabel: [trace.serverName, trace.toolName].filter(Boolean).join(' / ') || trace.toolName,
    serverLabel: trace.serverName,
    targetLabel: trace.targetSummary || undefined,
    targetSourceLabel,
    resultLabel: [trace.finalBlockTitle, trace.inferredBlockType].filter(Boolean).join(' · ') || undefined
      || trace.finalSummary
      || undefined,
    requiredLabel: trace.requiredFields && trace.requiredFields.length > 0
      ? trace.requiredFields.join(', ')
      : undefined,
    rawArgsLabel: formatCompactJson(trace.rawArgs),
    normalizedArgsLabel: formatCompactJson(trace.normalizedArgs),
    resolvedArgsLabel: formatCompactJson(trace.resolvedArgs),
    unwrapLabel: trace.unwrapPath && trace.unwrapPath.length > 0
      ? trace.unwrapPath.join(' -> ')
      : undefined,
    recoveryLabel: recoverySteps.length > 0
      ? recoverySteps.join(' -> ')
      : undefined
  };
}

export function findLatestMcpExecutionTrace(entries: readonly ILogEntry[]): IParsedMcpExecutionTrace | undefined {
  for (let i = entries.length - 1; i >= 0; i--) {
    const trace = parseMcpExecutionTrace(entries[i]);
    if (trace) {
      return trace;
    }
  }

  return undefined;
}
