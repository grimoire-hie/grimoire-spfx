jest.mock('../graph/SitesService', () => ({
  SitesService: jest.fn(() => ({ kind: 'sites-service' }))
}));

jest.mock('../graph/PeopleService', () => ({
  PeopleService: jest.fn(() => ({ kind: 'people-service' }))
}));

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    captureCurrentSourceContext: jest.fn(() => undefined)
  }
}));

jest.mock('./ToolRegistryRuntimeHandlers', () => ({
  dispatchRuntimeHandledTool: jest.fn()
}));

import { handleFunctionCall } from './handleFunctionCall';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import * as RuntimeHandlers from './ToolRegistryRuntimeHandlers';
import * as RuntimeLifecycle from './ToolRegistryRuntime';

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

describe('handleFunctionCall integration', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  it('runs adapter + deps factory + dispatcher on sync success', () => {
    const store = createStore();
    (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock)
      .mockReturnValue({ output: JSON.stringify({ success: true }), phase: 'complete' });
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const output = handleFunctionCall('search_people', { query: 'john' }, store);

    expect(output).toBe(JSON.stringify({ success: true }));
    expect(RuntimeHandlers.dispatchRuntimeHandledTool).toHaveBeenCalledTimes(1);
    const runtimeCall = (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock).mock.calls[0];
    expect(runtimeCall[0]).toBe('search_people');
    expect(runtimeCall[1]).toEqual({ query: 'john' });
    expect(runtimeCall[2].store).toBe(store);
    expect(runtimeCall[2].awaitAsync).toBe(false);
    expect(runtimeCall[2].aadClient).toBeUndefined();
    expect(runtimeCall[2].sitesService).toBeUndefined();
    expect(runtimeCall[2].peopleService).toBeUndefined();
    expect(store.setActivityStatus).toHaveBeenNthCalledWith(1, 'Searching people...');
    expect(store.setActivityStatus).toHaveBeenLastCalledWith('');
    expect(completeSpy.mock.calls[0][4]).toBe('complete');
  });

  it('propagates async error result and still completes lifecycle', async () => {
    const store = createStore();
    const aadClient = { client: 'aad' } as unknown;
    store.aadHttpClient = aadClient as never;
    (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock)
      .mockReturnValue(Promise.resolve({
        output: JSON.stringify({ success: false, error: 'bad' }),
        phase: 'error'
      }));
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    await expect(handleFunctionCall('read_file_content', {}, store, true))
      .resolves.toBe(JSON.stringify({ success: false, error: 'bad' }));

    const runtimeCall = (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock).mock.calls[0];
    expect(runtimeCall[2].awaitAsync).toBe(true);
    expect(runtimeCall[2].aadClient).toBe(aadClient);
    expect(runtimeCall[2].sitesService).toBeDefined();
    expect(runtimeCall[2].peopleService).toBeDefined();
    expect(store.setActivityStatus).toHaveBeenNthCalledWith(1, 'Reading file...');
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });

  it('returns unknown-function error when runtime dispatch has no handler', () => {
    const store = createStore();
    (RuntimeHandlers.dispatchRuntimeHandledTool as jest.Mock).mockReturnValue(undefined);
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const output = handleFunctionCall('unknown_tool', {}, store);

    expect(output).toBe(JSON.stringify({ error: 'Unknown function: unknown_tool' }));
    expect(store.setActivityStatus).not.toHaveBeenCalled();
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });
});
