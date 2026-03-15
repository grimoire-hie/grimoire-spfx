jest.mock('../graph/SitesService', () => ({
  SitesService: jest.fn(() => ({ kind: 'sites-service' }))
}));

jest.mock('../graph/PeopleService', () => ({
  PeopleService: jest.fn(() => ({ kind: 'people-service' }))
}));

jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    captureCurrentSourceContext: jest.fn()
  }
}));

import * as SitesModule from '../graph/SitesService';
import * as PeopleModule from '../graph/PeopleService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { createToolRuntimeHandlerDeps } from './ToolRuntimeDepsFactory';
import type { IFunctionCallStore } from './ToolRuntimeContracts';

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

describe('ToolRuntimeDepsFactory', () => {
  beforeEach(() => {
    jest.resetAllMocks();
    (hybridInteractionEngine.captureCurrentSourceContext as jest.Mock).mockReturnValue(undefined);
  });

  it('does not create graph services without aad client', () => {
    const store = createStore();
    const deps = createToolRuntimeHandlerDeps(store, false);

    expect(deps.store).toBe(store);
    expect(deps.awaitAsync).toBe(false);
    expect(deps.aadClient).toBeUndefined();
    expect(deps.sitesService).toBeUndefined();
    expect(deps.peopleService).toBeUndefined();
    expect(deps.sourceContext).toBeUndefined();
    expect(SitesModule.SitesService).not.toHaveBeenCalled();
    expect(PeopleModule.PeopleService).not.toHaveBeenCalled();
  });

  it('creates graph services when aad client is present', () => {
    const store = createStore();
    const aadClient = { test: 'aad-client' } as unknown;
    store.aadHttpClient = aadClient as never;

    const deps = createToolRuntimeHandlerDeps(store, true);

    expect(deps.awaitAsync).toBe(true);
    expect(deps.aadClient).toBe(aadClient);
    expect(deps.sitesService).toBeDefined();
    expect(deps.peopleService).toBeDefined();
    expect(hybridInteractionEngine.captureCurrentSourceContext).toHaveBeenCalledTimes(1);
    expect(SitesModule.SitesService).toHaveBeenCalledWith(aadClient);
    expect(PeopleModule.PeopleService).toHaveBeenCalledWith(aadClient);
  });
});
