import { getToolCatalog } from '../../config/toolCatalog';
import {
  CONTENT_RUNTIME_TOOL_NAMES,
  MCP_RUNTIME_TOOL_NAMES,
  SEARCH_RUNTIME_TOOL_NAMES,
  UI_PERSONAL_RUNTIME_TOOL_NAMES
} from './ToolRuntimeHandlerPartitions';
import { RUNTIME_HANDLED_TOOL_NAMES } from './ToolRuntimeHandlerRegistry';

describe('ToolRuntimeHandlerPartitions', () => {
  const partitions = [
    SEARCH_RUNTIME_TOOL_NAMES,
    CONTENT_RUNTIME_TOOL_NAMES,
    MCP_RUNTIME_TOOL_NAMES,
    UI_PERSONAL_RUNTIME_TOOL_NAMES
  ] as const;

  it('has no overlapping tool names across partitions', () => {
    const allPartitionNames = partitions.flat();
    const uniqueNames = new Set(allPartitionNames);
    expect(uniqueNames.size).toBe(allPartitionNames.length);
  });

  it('covers the exact same tools as runtime handler registry', () => {
    const allPartitionNames = partitions.flat().slice().sort();
    const registryNames = [...RUNTIME_HANDLED_TOOL_NAMES].sort();
    expect(allPartitionNames).toEqual(registryNames);
  });

  it('matches the shared tool catalog partition metadata', () => {
    const catalog = getToolCatalog();

    expect(SEARCH_RUNTIME_TOOL_NAMES).toEqual(
      catalog.filter((tool) => tool.runtimePartition === 'search').map((tool) => tool.name)
    );
    expect(CONTENT_RUNTIME_TOOL_NAMES).toEqual(
      catalog.filter((tool) => tool.runtimePartition === 'content').map((tool) => tool.name)
    );
    expect(MCP_RUNTIME_TOOL_NAMES).toEqual(
      catalog.filter((tool) => tool.runtimePartition === 'mcp').map((tool) => tool.name)
    );
    expect(UI_PERSONAL_RUNTIME_TOOL_NAMES).toEqual(
      catalog.filter((tool) => tool.runtimePartition === 'ui-personal').map((tool) => tool.name)
    );
  });
});
