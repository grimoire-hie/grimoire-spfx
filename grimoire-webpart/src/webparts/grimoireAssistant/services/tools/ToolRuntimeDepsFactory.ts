import { SitesService } from '../graph/SitesService';
import { PeopleService } from '../graph/PeopleService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';

export function createToolRuntimeHandlerDeps(
  store: IFunctionCallStore,
  awaitAsync: boolean
): IToolRuntimeHandlerDeps {
  const aadClient = store.aadHttpClient;
  return {
    store,
    awaitAsync,
    aadClient,
    sitesService: aadClient ? new SitesService(aadClient) : undefined,
    peopleService: aadClient ? new PeopleService(aadClient) : undefined,
    sourceContext: hybridInteractionEngine.captureCurrentSourceContext()
  };
}
