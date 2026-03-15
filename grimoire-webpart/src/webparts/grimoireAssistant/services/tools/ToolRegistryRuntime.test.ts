import {
  beginToolDispatch,
  completeToolDispatch,
  getToolHandlerDefinition
} from './ToolRegistryRuntime';
import { getTools } from '../realtime/ToolRegistry';
import { RUNTIME_HANDLED_TOOL_NAMES } from './ToolRuntimeHandlerRegistry';

describe('ToolRegistryRuntime', () => {
  it('provides declarative handler metadata', () => {
    const def = getToolHandlerDefinition('search_sharepoint');
    expect(def).toBeDefined();
    expect(def?.name).toBe('search_sharepoint');
    expect(def?.defaultBlockType).toBe('search-results');
    expect(def?.asyncBehavior).toBe('async');
  });

  it('dispatch lifecycle sets and clears activity status', () => {
    const statuses: string[] = [];
    const store = {
      setActivityStatus: (s: string) => { statuses.push(s); }
    };

    const ctx = beginToolDispatch('search_people', { query: 'john' }, store, true);
    completeToolDispatch(ctx, 'search_people', { query: 'john' }, store, 'complete');

    expect(statuses.length).toBeGreaterThanOrEqual(2);
    expect(statuses[0]).toBe('Searching people...');
    expect(statuses[statuses.length - 1]).toBe('');
  });

  it('does not emit transcript status entries on dispatch', () => {
    const store = {
      setActivityStatus: jest.fn()
    };

    beginToolDispatch('read_file_content', {}, store, true);

    expect(store.setActivityStatus).toHaveBeenCalledWith('Reading file...');
  });

  it('covers all registered tools via runtime handlers', () => {
    const registryNames = getTools().map((t) => t.name).sort();
    const runtimeNames = [...RUNTIME_HANDLED_TOOL_NAMES].sort();
    expect(runtimeNames).toEqual(registryNames);
  });
});
