/**
 * ToolRuntimeHandlerRegistry
 * Canonical list of tools expected to be runtime-handled.
 * Kept dependency-free so tests can import it without SPFx runtime modules.
 */

import {
  CONTENT_RUNTIME_TOOL_NAMES,
  MCP_RUNTIME_TOOL_NAMES,
  SEARCH_RUNTIME_TOOL_NAMES,
  UI_PERSONAL_RUNTIME_TOOL_NAMES
} from './ToolRuntimeHandlerPartitions';

export const RUNTIME_HANDLED_TOOL_NAMES = [
  ...SEARCH_RUNTIME_TOOL_NAMES,
  ...CONTENT_RUNTIME_TOOL_NAMES,
  ...MCP_RUNTIME_TOOL_NAMES,
  ...UI_PERSONAL_RUNTIME_TOOL_NAMES
] as const;

export type RuntimeHandledToolName = typeof RUNTIME_HANDLED_TOOL_NAMES[number];

const RUNTIME_HANDLED_TOOL_NAME_SET: ReadonlySet<string> = new Set(
  RUNTIME_HANDLED_TOOL_NAMES as readonly string[]
);

export function isRuntimeHandledToolName(name: string): name is RuntimeHandledToolName {
  return RUNTIME_HANDLED_TOOL_NAME_SET.has(name);
}
