/**
 * ToolRuntimeHandlerPartitions
 * Dependency-free domain partitions for runtime-handled tool names.
 */

import {
  getToolCatalog,
  type ToolCatalogNameByPartition,
  type ToolRuntimePartition
} from '../../config/toolCatalog';

const TOOL_CATALOG = getToolCatalog();

function getPartitionToolNames<P extends ToolRuntimePartition>(partition: P): ToolCatalogNameByPartition<P>[] {
  return TOOL_CATALOG
    .filter((tool) => tool.runtimePartition === partition)
    .map((tool) => tool.name) as ToolCatalogNameByPartition<P>[];
}

export const SEARCH_RUNTIME_TOOL_NAMES = getPartitionToolNames('search') as readonly ToolCatalogNameByPartition<'search'>[];
export type SearchRuntimeToolName = ToolCatalogNameByPartition<'search'>;

export const CONTENT_RUNTIME_TOOL_NAMES = getPartitionToolNames('content') as readonly ToolCatalogNameByPartition<'content'>[];
export type ContentRuntimeToolName = ToolCatalogNameByPartition<'content'>;

export const MCP_RUNTIME_TOOL_NAMES = getPartitionToolNames('mcp') as readonly ToolCatalogNameByPartition<'mcp'>[];
export type McpRuntimeToolName = ToolCatalogNameByPartition<'mcp'>;

export const UI_PERSONAL_RUNTIME_TOOL_NAMES = getPartitionToolNames('ui-personal') as readonly ToolCatalogNameByPartition<'ui-personal'>[];
export type UiPersonalRuntimeToolName = ToolCatalogNameByPartition<'ui-personal'>;

export type PartitionedRuntimeToolName =
  | SearchRuntimeToolName
  | ContentRuntimeToolName
  | McpRuntimeToolName
  | UiPersonalRuntimeToolName;
