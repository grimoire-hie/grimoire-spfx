# Tool Runtime Architecture

This folder uses a layered runtime dispatch design to keep `handleFunctionCall` thin and testable.

## Layers

1. `handleFunctionCall.ts`  
   Entry adapter for voice/text paths. It wires runtime dependencies and delegates execution.

2. `ToolRuntimeDepsFactory.ts`  
   Builds handler dependencies (`aadClient`, `SitesService`, `PeopleService`) from store state.

3. `ToolRuntimeDispatcher.ts`  
   Owns dispatch lifecycle:
   - `beginToolDispatch` / `completeToolDispatch`
   - unknown tool fallback
   - sync/async handling

4. `ToolRegistryRuntimeHandlers.ts` + domain handler modules  
   Own tool behavior by domain:
   - search
   - content
   - MCP
   - UI/personal

5. `ToolRuntimeHandlerRegistry.ts` + `ToolRuntimeHandlerPartitions.ts`  
   Canonical runtime tool-name list and domain partitions, including typed guard.

6. `ToolRuntimeMetadata.ts`  
   Centralized runtime metadata (activity labels, default block types, async/completion policy).

## Outcome Contract

Domain handlers return explicit dispatch outcomes:

```ts
{ output: string; phase: 'complete' | 'error' }
```

## Invariants

- Runtime-handled names must match realtime tool registry coverage tests.
- Partition arrays must have no overlap.
- Registry boundary forwards explicit outcome objects for runtime-handled tools.
- Dispatcher must always close lifecycle with `completeToolDispatch` on all non-crash paths.
- Unknown tools must return a stable JSON error payload.
