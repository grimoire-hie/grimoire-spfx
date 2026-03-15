/**
 * ToolRegistryRuntimeHandlers
 * Composes all runtime handler domains and dispatches by tool name.
 */

import { isRuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type {
  IToolRuntimeHandlerDeps,
  ToolRuntimeHandler,
  ToolRuntimeHandlerResult
} from './ToolRuntimeHandlerTypes';
import { buildSearchRuntimeHandlers } from './ToolRegistryRuntimeHandlersSearch';
import { buildContentRuntimeHandlers } from './ToolRegistryRuntimeHandlersContent';
import { buildMcpRuntimeHandlers } from './ToolRegistryRuntimeHandlersMcp';
import { buildUiAndPersonalRuntimeHandlers } from './ToolRegistryRuntimeHandlersUiAndPersonal';
import { connectToM365Server, extractMcpReply, findExistingSession } from './ToolRuntimeSharedHelpers';

const runtimeHandlers: Record<RuntimeHandledToolName, ToolRuntimeHandler> = {
  ...buildSearchRuntimeHandlers({
    findExistingSession,
    connectToM365Server
  }),

  ...buildContentRuntimeHandlers({
    extractMcpReply,
    findExistingSession,
    connectToM365Server
  }),

  ...buildMcpRuntimeHandlers({
    findExistingSession,
    connectToM365Server
  }),

  ...buildUiAndPersonalRuntimeHandlers()
};

export function dispatchRuntimeHandledTool(
  funcName: string,
  args: Record<string, unknown>,
  deps: IToolRuntimeHandlerDeps
): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> | undefined {
  if (!isRuntimeHandledToolName(funcName)) return undefined;
  return runtimeHandlers[funcName](args, deps);
}
