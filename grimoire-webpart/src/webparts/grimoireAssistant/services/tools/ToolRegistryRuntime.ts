/**
 * ToolRegistryRuntime
 * Declarative runtime metadata and dispatch envelopes for tool execution.
 */

import { getTools } from '../realtime/ToolRegistry';
import { createCorrelationId, type IToolExecutionEnvelope } from '../hie/HAEContracts';
import { logService } from '../logging/LogService';
import { isRuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import {
  ASYNC_RUNTIME_TOOLS,
  DEFAULT_BLOCK_TYPES,
  TOOL_ACTIVITY_LABELS
} from './ToolRuntimeMetadata';
import type { BlockType } from '../../models/IBlock';

export interface IToolHandlerDefinition {
  name: string;
  inputSchema?: unknown;
  activityLabel?: string;
  asyncBehavior: 'sync' | 'async' | 'mixed';
  defaultBlockType?: BlockType;
}

export interface IToolDispatchContext {
  correlationId: string;
  startedAt: number;
  activityLabel?: string;
}

interface IToolDispatchStore {
  setActivityStatus: (status: string) => void;
}

const TOOL_HANDLERS: Record<string, IToolHandlerDefinition> = {};
const tools = getTools();
for (let i = 0; i < tools.length; i++) {
  const t = tools[i];
  const runtimeToolName = isRuntimeHandledToolName(t.name) ? t.name : undefined;
  TOOL_HANDLERS[t.name] = {
    name: t.name,
    inputSchema: t.parameters,
    activityLabel: runtimeToolName ? TOOL_ACTIVITY_LABELS[runtimeToolName] : undefined,
    asyncBehavior: runtimeToolName && ASYNC_RUNTIME_TOOLS.has(runtimeToolName) ? 'async' : 'sync',
    defaultBlockType: runtimeToolName ? DEFAULT_BLOCK_TYPES[runtimeToolName] : undefined
  };
}

export function getToolHandlerDefinition(name: string): IToolHandlerDefinition | undefined {
  return TOOL_HANDLERS[name];
}

export function beginToolDispatch(
  funcName: string,
  args: Record<string, unknown>,
  store: IToolDispatchStore,
  awaitAsync: boolean
): IToolDispatchContext {
  const handler = getToolHandlerDefinition(funcName);
  const correlationId = createCorrelationId('tool');
  const startedAt = performance.now();

  if (handler && handler.activityLabel) {
    store.setActivityStatus(handler.activityLabel);
  }

  const envelope: IToolExecutionEnvelope = {
    envelopeId: createCorrelationId('toolenv'),
    correlationId,
    createdAt: Date.now(),
    toolName: funcName,
    phase: 'dispatch',
    awaitAsync,
    args
  };
  logService.info('llm', `Tool dispatch: ${funcName}`, JSON.stringify(envelope));

  return {
    correlationId,
    startedAt,
    activityLabel: handler?.activityLabel
  };
}

export function completeToolDispatch(
  ctx: IToolDispatchContext,
  funcName: string,
  args: Record<string, unknown>,
  store: IToolDispatchStore,
  phase: 'complete' | 'error'
): void {
  const durationMs = Math.round(performance.now() - ctx.startedAt);
  const envelope: IToolExecutionEnvelope = {
    envelopeId: createCorrelationId('toolenv'),
    correlationId: ctx.correlationId,
    createdAt: Date.now(),
    toolName: funcName,
    phase,
    awaitAsync: false,
    args,
    durationMs
  };

  logService.info('llm', `Tool ${phase}: ${funcName}`, JSON.stringify(envelope), durationMs);

  if (ctx.activityLabel) {
    store.setActivityStatus('');
  }
}
