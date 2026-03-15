import { getToolCatalog } from '../../config/toolCatalog';
import { isRuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import {
  ASYNC_RUNTIME_TOOLS,
  DEFAULT_BLOCK_TYPES,
  TOOL_ACTIVITY_LABELS
} from './ToolRuntimeMetadata';

describe('ToolRuntimeMetadata', () => {
  it('uses only runtime-handled names for activity labels and default block types', () => {
    Object.keys(TOOL_ACTIVITY_LABELS).forEach((name) => {
      expect(isRuntimeHandledToolName(name)).toBe(true);
    });
    Object.keys(DEFAULT_BLOCK_TYPES).forEach((name) => {
      expect(isRuntimeHandledToolName(name)).toBe(true);
    });
  });

  it('uses only runtime-handled names in async runtime set', () => {
    ASYNC_RUNTIME_TOOLS.forEach((name) => {
      expect(isRuntimeHandledToolName(name)).toBe(true);
    });
  });

  it('stays aligned with the shared tool catalog', () => {
    getToolCatalog().forEach((tool) => {
      const name = tool.name as keyof typeof TOOL_ACTIVITY_LABELS;
      const blockName = tool.name as keyof typeof DEFAULT_BLOCK_TYPES;
      const runtimeHandledName = tool.name as Parameters<typeof ASYNC_RUNTIME_TOOLS.has>[0];

      expect(TOOL_ACTIVITY_LABELS[name]).toBe(tool.activityLabel);
      expect(DEFAULT_BLOCK_TYPES[blockName]).toBe(tool.defaultBlockType);
      expect(ASYNC_RUNTIME_TOOLS.has(runtimeHandledName)).toBe(tool.asyncBehavior === 'async');
    });
  });
});
