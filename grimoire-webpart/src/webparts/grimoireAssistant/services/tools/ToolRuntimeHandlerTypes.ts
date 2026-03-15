import type { AadHttpClient } from '@microsoft/sp-http';
import type { IFunctionCallStore } from './ToolRuntimeContracts';
import type { PeopleService } from '../graph/PeopleService';
import type { SitesService } from '../graph/SitesService';
import type { IHieSourceContext } from '../hie/HIETypes';

export interface IToolRuntimeDispatchOutcome {
  output: string;
  phase: 'complete' | 'error';
}

export type ToolRuntimeHandlerResult = IToolRuntimeDispatchOutcome;

export interface IToolRuntimeHandlerDeps {
  store: IFunctionCallStore;
  awaitAsync: boolean;
  aadClient: AadHttpClient | undefined;
  sitesService: SitesService | undefined;
  peopleService: PeopleService | undefined;
  sourceContext?: IHieSourceContext;
}

export type ToolRuntimeHandler = (
  args: Record<string, unknown>,
  deps: IToolRuntimeHandlerDeps
) => ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult>;
