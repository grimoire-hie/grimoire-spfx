import { logService } from '../logging/LogService';
import { beginToolDispatch, completeToolDispatch } from './ToolRegistryRuntime';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { ToolRuntimeHandlerResult } from './ToolRuntimeHandlerTypes';
import { errorOutcome } from './ToolRuntimeOutcomeHelpers';

export type RuntimeToolInvoker = () => ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> | undefined;

export interface IRuntimeDispatchOptions {
  funcName: string;
  args: Record<string, unknown>;
  store: IFunctionCallStore;
  awaitAsync: boolean;
  invoke: RuntimeToolInvoker;
}

function isPromiseLike(value: unknown): value is Promise<ToolRuntimeHandlerResult> {
  return !!value && typeof (value as Promise<ToolRuntimeHandlerResult>).then === 'function';
}

export function dispatchToolRuntime(options: IRuntimeDispatchOptions): string | Promise<string> {
  const { funcName, args, store, awaitAsync, invoke } = options;
  const dispatchCtx = beginToolDispatch(funcName, args, store, awaitAsync);

  try {
    const runtimeResult = invoke();
    if (runtimeResult === undefined) {
      logService.warning('system', `Unknown function: ${funcName}`);
      const unknown = errorOutcome(JSON.stringify({ error: `Unknown function: ${funcName}` }));
      completeToolDispatch(dispatchCtx, funcName, args, store, 'error');
      return unknown.output;
    }

    if (isPromiseLike(runtimeResult)) {
      return runtimeResult.then((result) => {
        completeToolDispatch(dispatchCtx, funcName, args, store, result.phase);
        return result.output;
      }).catch((err: Error) => {
        completeToolDispatch(dispatchCtx, funcName, args, store, 'error');
        throw err;
      });
    }

    completeToolDispatch(dispatchCtx, funcName, args, store, runtimeResult.phase);
    return runtimeResult.output;
  } catch (err) {
    completeToolDispatch(dispatchCtx, funcName, args, store, 'error');
    throw err;
  }
}
