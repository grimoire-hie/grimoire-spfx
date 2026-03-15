import type { BlockType } from '../../models/IBlock';
import { getToolCatalog } from '../../config/toolCatalog';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';

const TOOL_CATALOG = getToolCatalog();

export const TOOL_ACTIVITY_LABELS: Partial<Record<RuntimeHandledToolName, string>> = TOOL_CATALOG.reduce(
  (acc, tool) => {
    if (tool.activityLabel) {
      acc[tool.name as RuntimeHandledToolName] = tool.activityLabel;
    }
    return acc;
  },
  {} as Partial<Record<RuntimeHandledToolName, string>>
);

export const DEFAULT_BLOCK_TYPES: Partial<Record<RuntimeHandledToolName, BlockType>> = TOOL_CATALOG.reduce(
  (acc, tool) => {
    if (tool.defaultBlockType) {
      acc[tool.name as RuntimeHandledToolName] = tool.defaultBlockType;
    }
    return acc;
  },
  {} as Partial<Record<RuntimeHandledToolName, BlockType>>
);

const ASYNC_RUNTIME_TOOL_NAMES = TOOL_CATALOG
  .filter((tool) => tool.asyncBehavior === 'async')
  .map((tool) => tool.name) as RuntimeHandledToolName[];

export const ASYNC_RUNTIME_TOOLS: ReadonlySet<RuntimeHandledToolName> = new Set(
  ASYNC_RUNTIME_TOOL_NAMES as readonly RuntimeHandledToolName[]
);
