/**
 * handleFunctionCall — Shared tool dispatch for both voice (WebRTC) and text (HTTP) paths.
 * Runtime handlers own tool behavior; this file wires dependencies.
 */

import { dispatchRuntimeHandledTool } from './ToolRegistryRuntimeHandlers';
import { dispatchToolRuntime } from './ToolRuntimeDispatcher';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import { createToolRuntimeHandlerDeps } from './ToolRuntimeDepsFactory';
export { findExistingSession, connectToM365Server } from './ToolRuntimeSharedHelpers';

export function handleFunctionCall(
  funcName: string,
  args: Record<string, unknown>,
  store: IFunctionCallStore
): string;
export function handleFunctionCall(
  funcName: string,
  args: Record<string, unknown>,
  store: IFunctionCallStore,
  awaitAsync: true
): Promise<string>;
export function handleFunctionCall(
  funcName: string,
  args: Record<string, unknown>,
  store: IFunctionCallStore,
  awaitAsync?: boolean
): string | Promise<string> {
  const runtimeDeps = createToolRuntimeHandlerDeps(store, !!awaitAsync);

  return dispatchToolRuntime({
    funcName,
    args,
    store,
    awaitAsync: !!awaitAsync,
    invoke: () => dispatchRuntimeHandledTool(funcName, args, runtimeDeps)
  });
}
