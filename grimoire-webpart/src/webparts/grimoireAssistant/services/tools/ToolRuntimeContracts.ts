import type { AadHttpClient } from '@microsoft/sp-http';
import type { IBlock } from '../../models/IBlock';
import type { IMcpConnection } from '../../models/IMcpTypes';
import type { IProxyConfig, PublicWebSearchCapabilityStatus } from '../../store/useGrimoireStore';
import type { Expression } from '../avatar/ExpressionEngine';
import type { IUserContext } from '../context/ContextService';

export interface IFunctionCallStore {
  aadHttpClient: AadHttpClient | undefined;
  userApiClient?: AadHttpClient | undefined;
  proxyConfig: IProxyConfig | undefined;
  getToken: ((resource: string) => Promise<string>) | undefined;
  mcpEnvironmentId: string | undefined;
  userContext: IUserContext | undefined;
  avatarEnabled?: boolean;
  publicWebSearchEnabled?: boolean;
  publicWebSearchCapability?: PublicWebSearchCapabilityStatus;
  publicWebSearchCapabilityDetail?: string;
  copilotWebGroundingEnabled?: boolean;
  setPublicWebSearchCapability?: (status: PublicWebSearchCapabilityStatus, detail?: string) => void;
  mcpConnections: IMcpConnection[];
  pushBlock: (block: IBlock) => void;
  updateBlock: (blockId: string, updates: Record<string, unknown>) => void;
  removeBlock: (blockId: string) => void;
  clearBlocks: () => void;
  setExpression: (expression: Expression) => void;
  setActivityStatus: (status: string) => void;
}
