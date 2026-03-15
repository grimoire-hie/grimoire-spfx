import type { IToolRuntimeDispatchOutcome } from './ToolRuntimeHandlerTypes';

export function completeOutcome(output: string): IToolRuntimeDispatchOutcome {
  return { output, phase: 'complete' };
}

export function errorOutcome(output: string): IToolRuntimeDispatchOutcome {
  return { output, phase: 'error' };
}
