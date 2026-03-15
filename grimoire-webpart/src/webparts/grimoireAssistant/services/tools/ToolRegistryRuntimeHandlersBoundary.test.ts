const showMarkdownHandler = jest.fn();
const searchPeopleHandler = jest.fn();

jest.mock('./ToolRegistryRuntimeHandlersSearch', () => ({
  buildSearchRuntimeHandlers: jest.fn(() => ({
    search_people: (args: Record<string, unknown>, deps: unknown) => searchPeopleHandler(args, deps)
  }))
}));

jest.mock('./ToolRegistryRuntimeHandlersContent', () => ({
  buildContentRuntimeHandlers: jest.fn(() => ({}))
}));

jest.mock('./ToolRegistryRuntimeHandlersMcp', () => ({
  buildMcpRuntimeHandlers: jest.fn(() => ({}))
}));

jest.mock('./ToolRegistryRuntimeHandlersUiAndPersonal', () => ({
  buildUiAndPersonalRuntimeHandlers: jest.fn(() => ({
    show_markdown: (args: Record<string, unknown>, deps: unknown) => showMarkdownHandler(args, deps)
  }))
}));

import { dispatchRuntimeHandledTool } from './ToolRegistryRuntimeHandlers';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';

function createDeps(): IToolRuntimeHandlerDeps {
  return {
    store: {
      aadHttpClient: undefined,
      proxyConfig: undefined,
      getToken: undefined,
      mcpEnvironmentId: undefined,
      userContext: undefined,
      mcpConnections: [],
      pushBlock: jest.fn(),
      updateBlock: jest.fn(),
      removeBlock: jest.fn(),
      clearBlocks: jest.fn(),
      setExpression: jest.fn(),
      setActivityStatus: jest.fn()
    },
    awaitAsync: false,
    aadClient: undefined,
    sitesService: undefined,
    peopleService: undefined
  };
}

describe('ToolRegistryRuntimeHandlers boundary dispatch', () => {
  beforeEach(() => {
    showMarkdownHandler.mockReset();
    searchPeopleHandler.mockReset();
  });

  it('returns sync dispatch outcome from domain handler', () => {
    showMarkdownHandler.mockReturnValue({
      output: JSON.stringify({ success: true }),
      phase: 'complete'
    });
    const result = dispatchRuntimeHandledTool('show_markdown', { title: 'x' }, createDeps());

    expect(result).toEqual({
      output: JSON.stringify({ success: true }),
      phase: 'complete'
    });
  });

  it('returns async dispatch outcome from domain handler', async () => {
    searchPeopleHandler.mockResolvedValue({
      output: JSON.stringify({ success: false, error: 'bad request' }),
      phase: 'error'
    });
    const result = dispatchRuntimeHandledTool('search_people', { query: 'john' }, createDeps());

    await expect(result).resolves.toEqual({
      output: JSON.stringify({ success: false, error: 'bad request' }),
      phase: 'error'
    });
  });

  it('preserves explicit phase returned by handler', () => {
    showMarkdownHandler.mockReturnValue({
      output: JSON.stringify({ success: true }),
      phase: 'error'
    });
    const result = dispatchRuntimeHandledTool('show_markdown', { title: 'x' }, createDeps());

    expect(result).toEqual({
      output: JSON.stringify({ success: true }),
      phase: 'error'
    });
  });

  it('returns undefined for unknown tool names', () => {
    const result = dispatchRuntimeHandledTool('unknown_tool', {}, createDeps());
    expect(result).toBeUndefined();
  });
});
