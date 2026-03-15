import { dispatchToolRuntime } from './ToolRuntimeDispatcher';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
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

describe('ToolRuntimeDispatcher', () => {
  beforeEach(() => {
    jest.restoreAllMocks();
  });

  it('completes with complete phase for sync dispatch outcome', () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const result = dispatchToolRuntime({
      funcName: 'show_markdown',
      args: { title: 'x' },
      store,
      awaitAsync: false,
      invoke: () => ({ output: JSON.stringify({ success: true }), phase: 'complete' })
    });

    expect(result).toBe(JSON.stringify({ success: true }));
    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('complete');
  });

  it('returns unknown tool error and completes with error phase', () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const result = dispatchToolRuntime({
      funcName: 'unknown_tool',
      args: {},
      store,
      awaitAsync: false,
      invoke: () => undefined
    });

    expect(result).toBe(JSON.stringify({ error: 'Unknown function: unknown_tool' }));
    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });

  it('completes with explicit error phase from async dispatch outcome', async () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const result = dispatchToolRuntime({
      funcName: 'read_file_content',
      args: {},
      store,
      awaitAsync: true,
      invoke: () => Promise.resolve({ output: JSON.stringify({ success: false, error: 'bad request' }), phase: 'error' })
    });

    await expect(result).resolves.toBe(JSON.stringify({ success: false, error: 'bad request' }));
    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });

  it('honors explicit phase from sync dispatch outcome', () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const result = dispatchToolRuntime({
      funcName: 'show_markdown',
      args: {},
      store,
      awaitAsync: false,
      invoke: () => ({ output: JSON.stringify({ success: true }), phase: 'error' })
    });

    expect(result).toBe(JSON.stringify({ success: true }));
    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });

  it('honors explicit phase from async dispatch outcome', async () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const result = dispatchToolRuntime({
      funcName: 'call_mcp_tool',
      args: {},
      store,
      awaitAsync: true,
      invoke: () => Promise.resolve({
        output: JSON.stringify({ success: false, error: 'simulated' }),
        phase: 'complete'
      })
    });

    await expect(result).resolves.toBe(JSON.stringify({ success: false, error: 'simulated' }));
    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('complete');
  });

  it('completes with error phase when async invoke rejects', async () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    const result = dispatchToolRuntime({
      funcName: 'call_mcp_tool',
      args: {},
      store,
      awaitAsync: true,
      invoke: () => Promise.reject(new Error('boom'))
    });

    await expect(result).rejects.toThrow('boom');
    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });

  it('completes with error phase when invoke throws synchronously', () => {
    const store = createStore();
    const completeSpy = jest.spyOn(RuntimeLifecycle, 'completeToolDispatch');

    expect(() => dispatchToolRuntime({
      funcName: 'show_markdown',
      args: {},
      store,
      awaitAsync: false,
      invoke: () => { throw new Error('sync-fail'); }
    })).toThrow('sync-fail');

    expect(completeSpy).toHaveBeenCalledTimes(1);
    expect(completeSpy.mock.calls[0][4]).toBe('error');
  });
});
