import {
  isRuntimeHandledToolName,
  RUNTIME_HANDLED_TOOL_NAMES
} from './ToolRuntimeHandlerRegistry';

describe('ToolRuntimeHandlerRegistry', () => {
  it('recognizes every registered runtime-handled tool name', () => {
    RUNTIME_HANDLED_TOOL_NAMES.forEach((name) => {
      expect(isRuntimeHandledToolName(name)).toBe(true);
    });
  });

  it('rejects unknown tool names', () => {
    expect(isRuntimeHandledToolName('unknown_tool')).toBe(false);
    expect(isRuntimeHandledToolName('search-people')).toBe(false);
  });
});
