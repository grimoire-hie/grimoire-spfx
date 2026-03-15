jest.mock('../graph/SitesService', () => ({
  SitesService: jest.fn()
}));

jest.mock('../graph/PeopleService', () => ({
  PeopleService: jest.fn()
}));

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    captureCurrentSourceContext: jest.fn(() => undefined)
  }
}));

jest.mock('./ToolRegistryRuntimeHandlers', () => ({
  dispatchRuntimeHandledTool: jest.fn()
}));

jest.mock('./ToolRuntimeDispatcher', () => ({
  dispatchToolRuntime: jest.fn()
}));

import { handleFunctionCall } from './handleFunctionCall';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeDispatchOutcome, ToolRuntimeHandlerResult } from './ToolRuntimeHandlerTypes';
import * as RuntimeHandlers from './ToolRegistryRuntimeHandlers';
import * as RuntimeDispatcher from './ToolRuntimeDispatcher';

function createStore(): IFunctionCallStore {
  return {
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
  };
}

describe('handleFunctionCall', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  it('delegates orchestration to ToolRuntimeDispatcher', () => {
    const store = createStore();
    (RuntimeDispatcher.dispatchToolRuntime as jest.Mock)
      .mockReturnValue(JSON.stringify({ success: true }));

    const output = handleFunctionCall('show_markdown', { title: 'x' }, store);

    expect(output).toBe(JSON.stringify({ success: true }));
    expect(RuntimeDispatcher.dispatchToolRuntime).toHaveBeenCalledTimes(1);
    const call = (RuntimeDispatcher.dispatchToolRuntime as jest.Mock).mock.calls[0][0] as {
      funcName: string;
      args: Record<string, unknown>;
      store: IFunctionCallStore;
      awaitAsync: boolean;
      invoke: () => ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> | undefined;
    };
    expect(call.funcName).toBe('show_markdown');
    expect(call.args).toEqual({ title: 'x' });
    expect(call.store).toBe(store);
    expect(call.awaitAsync).toBe(false);
    expect(typeof call.invoke).toBe('function');
  });

  it('wires runtime dependencies via invoke callback', () => {
    const store = createStore();
    const aadClient = { test: 'aad-client' } as unknown;
    store.aadHttpClient = aadClient as never;
    (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock)
      .mockReturnValue({ output: JSON.stringify({ success: true }), phase: 'complete' });
    (RuntimeDispatcher.dispatchToolRuntime as jest.Mock).mockImplementation((options: {
      invoke: () => ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> | undefined;
    }) => (options.invoke() as IToolRuntimeDispatchOutcome | undefined)?.output);

    const output = handleFunctionCall('search_sharepoint', { query: 'q' }, store, true);

    expect(output).toBe(JSON.stringify({ success: true }));
    expect(RuntimeHandlers.dispatchRuntimeHandledTool).toHaveBeenCalledTimes(1);
    const runtimeArgs = (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock).mock.calls[0];
    expect(runtimeArgs[0]).toBe('search_sharepoint');
    expect(runtimeArgs[1]).toEqual({ query: 'q' });
    expect(runtimeArgs[2].store).toBe(store);
    expect(runtimeArgs[2].awaitAsync).toBe(true);
    expect(runtimeArgs[2].aadClient).toBe(aadClient);
    expect(runtimeArgs[2].sitesService).toBeDefined();
    expect(runtimeArgs[2].peopleService).toBeDefined();
  });
});
