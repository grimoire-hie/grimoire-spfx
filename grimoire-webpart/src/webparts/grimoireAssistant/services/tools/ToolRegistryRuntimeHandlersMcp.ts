import { useGrimoireStore } from '../../store/useGrimoireStore';
import type { IErrorData, IInfoCardData, IProgressTrackerData } from '../../models/IBlock';
import type { IMcpConnection, IMcpContent } from '../../models/IMcpTypes';
import { createBlock } from '../../models/IBlock';
import { logService } from '../logging/LogService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import {
  M365_MCP_CATALOG,
  findServerForTool,
  resolveServerUrl
} from '../../models/McpServerCatalog';
import { McpClientService } from '../mcp/McpClientService';
import { mapMcpResultToBlock } from '../mcp/McpResultMapper';
import type { IExecuteCatalogMcpToolResult } from '../mcp/McpExecutionAdapter';
import { executeCatalogMcpTool, extractStructuredMcpPayload } from '../mcp/McpExecutionAdapter';
import {
  derivePersonalOneDriveSiteUrl,
  pickFirstNonEmptyString,
  resolveDefaultDocumentLibraryId,
  resolvePersonalOneDriveRootLocation
} from '../mcp/OdspLocationResolver';
import { extractServerIdFromUrl } from '../mcp/mcpUtils';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type { ToolRuntimeHandler, ToolRuntimeHandlerResult } from './ToolRuntimeHandlerTypes';
import type { McpRuntimeToolName } from './ToolRuntimeHandlerPartitions';
import type { IToolRuntimeHandlerDeps } from './ToolRuntimeHandlerTypes';
import { trackCreatedBlock, trackToolCompletion } from './ToolRuntimeHieHelpers';
import { completeOutcome, errorOutcome } from './ToolRuntimeOutcomeHelpers';

interface IMcpRuntimeHelpers {
  findExistingSession: (connections: IMcpConnection[], serverUrl: string) => string | undefined;
  connectToM365Server: (
    mcpClient: McpClientService,
    serverUrl: string,
    serverName: string,
    getToken: ((resource: string) => Promise<string>) | undefined
  ) => Promise<string>;
}

export { applySchemaArgAliases } from '../mcp/McpExecutionAdapter';

function pickArgString(args: Record<string, unknown>, ...keys: string[]): string | undefined {
  for (let i = 0; i < keys.length; i++) {
    const value = args[keys[i]];
    if (typeof value !== 'string') {
      continue;
    }
    const trimmed = value.trim();
    if (trimmed) {
      return trimmed;
    }
  }
  return undefined;
}

function isPersonalOneDriveRequest(args: Record<string, unknown>): boolean {
  return args.personalOneDrive === true;
}

function extractPersonalOneDriveSearchCollection(payload: unknown): {
  items: Array<Record<string, unknown>>;
  rebuild: (items: Array<Record<string, unknown>>) => unknown;
} | undefined {
  if (Array.isArray(payload)) {
    return {
      items: payload.filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object' && !Array.isArray(entry)),
      rebuild: (items) => items
    };
  }

  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return undefined;
  }

  const record = payload as Record<string, unknown>;
  const keys = ['value', 'items', 'results'];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const candidate = record[key];
    if (!Array.isArray(candidate)) {
      continue;
    }
    return {
      items: candidate.filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object' && !Array.isArray(entry)),
      rebuild: (items) => ({
        ...record,
        [key]: items
      })
    };
  }

  return undefined;
}

function scopePersonalOneDriveSearchContent(
  content: IMcpContent[],
  personalSiteUrl: string
): { content?: IMcpContent[]; error?: string } {
  const payload = extractStructuredMcpPayload(content).payload;
  const collection = extractPersonalOneDriveSearchCollection(payload);
  if (!collection) {
    return {
      error: 'I could not safely scope the ODSP search results to your personal OneDrive because the MCP response did not include a verifiable result list.'
    };
  }

  if (collection.items.length === 0) {
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(collection.rebuild([]))
      }]
    };
  }

  const normalizedSiteUrl = personalSiteUrl.toLowerCase().replace(/\/+$/, '');
  const filteredItems: Array<Record<string, unknown>> = [];
  for (let i = 0; i < collection.items.length; i++) {
    const item = collection.items[i];
    const resultUrl = pickFirstNonEmptyString(item.webUrl, item.url);
    if (!resultUrl) {
      return {
        error: 'I could not safely scope the ODSP search results to your personal OneDrive because one or more results did not include a verifiable URL.'
      };
    }

    const normalizedResultUrl = resultUrl.toLowerCase();
    if (normalizedResultUrl === normalizedSiteUrl || normalizedResultUrl.startsWith(`${normalizedSiteUrl}/`)) {
      filteredItems.push(item);
    }
  }

  return {
    content: [{
      type: 'text',
      text: JSON.stringify(collection.rebuild(filteredItems))
    }]
  };
}

function completeWithInfoCard(
  title: string,
  body: string,
  deps: IToolRuntimeHandlerDeps,
  toolName: 'use_m365_capability'
): ToolRuntimeHandlerResult {
  const infoBlock = createBlock('info-card', title, {
    kind: 'info-card',
    heading: title,
    body,
    icon: 'Info'
  } as IInfoCardData);
  trackCreatedBlock({ pushBlock: deps.store.pushBlock }, infoBlock, deps);
  trackToolCompletion(toolName, '', true, 1, deps);
  return completeOutcome(JSON.stringify({
    success: true,
    unsupported: true,
    summary: body
  }));
}

function normalizeListArtifactLabel(template?: string): 'list' | 'document library' {
  return template === 'documentLibrary' ? 'document library' : 'list';
}

function tryHandleAlreadyExistsConflict(
  toolName: string,
  toolArgs: Record<string, unknown>,
  execution: IExecuteCatalogMcpToolResult,
  deps: Parameters<typeof trackCreatedBlock>[2],
  completeToolName: 'call_mcp_tool' | 'use_m365_capability',
  outputContext: { serverName?: string; pushBlock: (block: ReturnType<typeof createBlock>) => void }
): ToolRuntimeHandlerResult | undefined {
  const error = execution.error || '';
  const normalizedToolName = toolName.toLowerCase();
  const normalizedError = error.toLowerCase();
  const isAlreadyExistsConflict = normalizedToolName === 'createlist'
    && normalizedError.includes('status=409')
    && normalizedError.includes('name already exists');

  if (!isAlreadyExistsConflict) {
    return undefined;
  }

  const resolvedArgs = execution.resolvedArgs || {};
  const resolvedList = resolvedArgs.list && typeof resolvedArgs.list === 'object' && !Array.isArray(resolvedArgs.list)
    ? resolvedArgs.list as Record<string, unknown>
    : {};
  const displayName = pickArgString(
    resolvedArgs,
    'displayName',
    'display_name',
    'listName',
    'list_name',
    'name'
  ) || pickArgString(
    toolArgs,
    'displayName',
    'display_name',
    'listName',
    'list_name',
    'name'
  ) || 'Already exists';
  const template = pickArgString(resolvedList, 'template') || pickArgString(toolArgs, 'template');
  const artifactLabel = normalizeListArtifactLabel(template);
  const siteName = execution.targetContext?.siteName?.trim();
  const siteSuffix = siteName ? ` in ${siteName}` : '';
  const body = `The ${artifactLabel} "${displayName}" already exists${siteSuffix}.`;
  const infoBlock = createBlock('info-card', displayName, {
    kind: 'info-card',
    heading: displayName,
    body,
    icon: 'Info'
  } as IInfoCardData);

  trackCreatedBlock({ pushBlock: outputContext.pushBlock }, infoBlock, deps);
  trackToolCompletion(completeToolName, '', true, 1, deps);
  logService.info('mcp', `Tool ${toolName} already exists: ${displayName}`);

  return completeOutcome(JSON.stringify({
    success: true,
    alreadyExists: true,
    toolName,
    server: outputContext.serverName,
    summary: body
  }));
}

export function buildMcpRuntimeHandlers(
  helpers: IMcpRuntimeHelpers
): Pick<Record<RuntimeHandledToolName, ToolRuntimeHandler>, McpRuntimeToolName> {
  const { findExistingSession, connectToM365Server } = helpers;
  return {
    list_mcp_tools: (args, deps): ToolRuntimeHandlerResult => {
      const { store } = deps;
      logService.info('mcp', 'List MCP tools');
      const requestedSessionId = args.session_id as string | undefined;
      const allTools: Array<{ name: string; description: string; server: string }> = [];
      store.mcpConnections.forEach((c: IMcpConnection) => {
        if (c.state === 'connected' && (!requestedSessionId || c.sessionId === requestedSessionId)) {
          c.tools.forEach((t) => {
            allTools.push({ name: t.name, description: t.description, server: c.serverName });
          });
        }
      });

      if (allTools.length === 0) {
        return completeOutcome(JSON.stringify({
          success: true,
          tools: [],
          message: requestedSessionId
            ? `No tools found for session "${requestedSessionId}".`
            : 'No MCP servers connected. Use connect_mcp_server first.'
        }));
      }

      const toolList = allTools.map((t) => `- **${t.name}** (${t.server}): ${t.description}`).join('\n');
      const toolsBlock = createBlock('markdown', 'Available MCP Tools', {
        kind: 'markdown', content: `## Available MCP Tools\n\n${toolList}`
      });
      trackCreatedBlock(store, toolsBlock, deps);

      return completeOutcome(JSON.stringify({ success: true, tools: allTools, count: allTools.length }));
    },

    connect_mcp_server: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const { store, awaitAsync } = deps;
      const serverUrl = args.server_url as string;
      const serverName = (args.server_name as string) || serverUrl;
      logService.info('mcp', `Connect MCP: ${serverName} (${serverUrl})`);
      store.setExpression('thinking');

      const proxyConf = store.proxyConfig;
      if (!proxyConf) {
        return errorOutcome(JSON.stringify({ success: false, error: 'No proxy config available' }));
      }

      let alreadyConnected = false;
      store.mcpConnections.forEach((c: IMcpConnection) => {
        if (c.serverUrl === serverUrl && c.state === 'connected') {
          alreadyConnected = true;
        }
      });
      if (alreadyConnected) {
        return completeOutcome(JSON.stringify({ success: true, message: `Already connected to ${serverName}` }));
      }

      const acquireTokenAndConnect = async (): Promise<void> => {
        let bearerToken: string | undefined;
        const getTokenFn = store.getToken;
        if (getTokenFn) {
          try {
            const audience = new URL(serverUrl).origin;
            bearerToken = await getTokenFn(audience);
          } catch (err) {
            logService.warning('mcp', `Token acquisition failed for ${serverUrl}: ${(err as Error).message}. Connecting without token.`);
          }
        }

        const mcpClient = new McpClientService(proxyConf.proxyUrl, proxyConf.proxyApiKey);
        const result = await mcpClient.connect(serverUrl, serverName, bearerToken);
        const currentStore = useGrimoireStore.getState();
        const connection: IMcpConnection = {
          sessionId: result.sessionId,
          serverUrl,
          serverName: result.serverName || serverName,
          tools: result.tools,
          state: 'connected',
          connectedAt: new Date()
        };
        currentStore.addMcpConnection(connection);
        trackToolCompletion('connect_mcp_server', '', true, result.tools.length, deps);
        logService.info('mcp', `Connected to ${serverName}: ${result.tools.length} tools available`);
      };

      const asyncResult = acquireTokenAndConnect().then(() => {
        return completeOutcome(JSON.stringify({ success: true, message: `Connected to ${serverName}` }));
      }).catch((err: Error) => {
        trackToolCompletion('connect_mcp_server', '', false, 0, deps);
        logService.error('mcp', `Failed to connect to ${serverName}: ${err.message}`);
        return errorOutcome(JSON.stringify({ success: false, error: err.message }));
      });

      if (awaitAsync) return asyncResult;
      return completeOutcome(JSON.stringify({ success: true, message: `Connecting to ${serverName}... Tools will be available shortly.` }));
    },

    call_mcp_tool: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const { store, awaitAsync } = deps;
      const toolName = args.tool_name as string;
      const sessionId = args.session_id as string | undefined;
      let toolArgs: Record<string, unknown> = {};
      if (typeof args.arguments_json === 'string') {
        try {
          toolArgs = JSON.parse(args.arguments_json as string) as Record<string, unknown>;
        } catch {
          return errorOutcome(JSON.stringify({ success: false, error: 'Invalid arguments_json — must be valid JSON.' }));
        }
      } else {
        toolArgs = (args.arguments_json as Record<string, unknown>) || {};
      }

      logService.info('mcp', `Call MCP tool: ${toolName}`);
      store.setExpression('thinking');

      const proxyConf2 = store.proxyConfig;
      if (!proxyConf2) {
        return errorOutcome(JSON.stringify({ success: false, error: 'No proxy config available' }));
      }

      let targetSessionId = sessionId;
      let targetConnection = targetSessionId
        ? store.mcpConnections.find((c) => c.sessionId === targetSessionId && c.state === 'connected')
        : undefined;
      let realMcpToolName = toolName;
      const lowerName = toolName.toLowerCase();

      if (targetSessionId && !targetConnection) {
        const byUrl = store.mcpConnections.find((c) =>
          c.state === 'connected' && extractServerIdFromUrl(c.serverUrl) === targetSessionId
        );
        targetSessionId = byUrl ? byUrl.sessionId : undefined;
        targetConnection = byUrl;
      }

      if (!targetSessionId) {
        const matches: Array<{ connection: IMcpConnection; toolName: string }> = [];
        store.mcpConnections.forEach((c: IMcpConnection) => {
          if (c.state !== 'connected') return;
          const match = c.tools.find((t) => t.name.toLowerCase() === lowerName);
          if (match) {
            matches.push({ connection: c, toolName: match.name });
          }
        });

        if (matches.length === 1) {
          targetSessionId = matches[0].connection.sessionId;
          targetConnection = matches[0].connection;
          realMcpToolName = matches[0].toolName;
        } else if (matches.length > 1) {
          const candidates = matches.map((m) => `${m.connection.serverName} (${m.connection.sessionId})`).join(', ');
          return errorOutcome(JSON.stringify({
            success: false,
            error: `Tool "${toolName}" exists on multiple connected servers. Specify session_id explicitly. Candidates: ${candidates}`
          }));
        }
      }

      if (!targetSessionId) {
        store.setExpression('confused');
        return errorOutcome(JSON.stringify({ success: false, error: `No connected MCP server has tool "${toolName}". Connect a server first.` }));
      }

      if (!targetConnection) {
        return errorOutcome(JSON.stringify({
          success: false,
          error: `Tool "${toolName}" is not associated with an active connected session.`
        }));
      }

      const retryServerUrl = targetConnection.serverUrl || '';
      const retryServerName = targetConnection.serverName || extractServerIdFromUrl(retryServerUrl);
      const mcpClient2 = new McpClientService(proxyConf2.proxyUrl, proxyConf2.proxyApiKey);
      const executePromise = executeCatalogMcpTool({
        serverId: extractServerIdFromUrl(retryServerUrl) || targetSessionId,
        serverName: retryServerName,
        serverUrl: retryServerUrl,
        toolName: realMcpToolName,
        rawArgs: toolArgs,
        connections: [targetConnection],
        getConnections: () => {
          const latest = useGrimoireStore.getState().mcpConnections.find((c) => c.sessionId === targetConnection!.sessionId);
          return latest ? [latest] : [targetConnection!];
        },
        mcpClient: mcpClient2,
        sessionHelpers: {
          findExistingSession,
          connectToM365Server
        },
        getToken: store.getToken,
        sourceContext: deps.sourceContext,
        taskContext: hybridInteractionEngine.getCurrentTaskContext(),
        artifacts: hybridInteractionEngine.getCurrentArtifacts(),
        currentSiteUrl: store.userContext?.currentSiteUrl
      });

      const asyncResult = executePromise.then((execution) => {
        const currentStore = useGrimoireStore.getState();
        logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));

        if (execution.success && execution.mcpResult) {
          const sid = extractServerIdFromUrl(retryServerUrl) || execution.serverId || '';
          const summary = mapMcpResultToBlock(sid, toolName, execution.mcpResult.content, currentStore.pushBlock,
            (block) => hybridInteractionEngine.onBlockCreated(block, deps.sourceContext)
          );
          trackToolCompletion('call_mcp_tool', '', true, 1, deps);
          return completeOutcome(JSON.stringify({ success: true, toolName, summary }));
        }

        const alreadyExistsOutcome = tryHandleAlreadyExistsConflict(
          toolName,
          toolArgs,
          execution,
          deps,
          'call_mcp_tool',
          { pushBlock: currentStore.pushBlock }
        );
        if (alreadyExistsOutcome) {
          return alreadyExistsOutcome;
        }

        const errMsg = execution.error || 'Tool execution failed';
        const errBlock = createBlock('error', `MCP Error: ${toolName}`, {
          kind: 'error', message: errMsg
        } as IErrorData);
        trackCreatedBlock(currentStore, errBlock, deps);
        trackToolCompletion('call_mcp_tool', '', false, 0, deps);
        hybridInteractionEngine.sendToolError(toolName, errMsg);
        logService.error('mcp', `Tool ${toolName} failed: ${execution.error}`);
        return errorOutcome(JSON.stringify({ success: false, error: errMsg }));
      }).catch((err: Error) => {
        const currentStore2 = useGrimoireStore.getState();
        const errBlock = createBlock('error', `MCP Error: ${toolName}`, {
          kind: 'error', message: err.message
        } as IErrorData);
        trackCreatedBlock(currentStore2, errBlock, deps);
        trackToolCompletion('call_mcp_tool', '', false, 0, deps);
        hybridInteractionEngine.sendToolError(toolName, err.message);
        logService.error('mcp', `Tool ${toolName} error: ${err.message}`);
        return errorOutcome(JSON.stringify({ success: false, error: err.message }));
      });

      if (awaitAsync) return asyncResult;
      return completeOutcome(JSON.stringify({ success: true, message: `Executing ${toolName}... Results will appear shortly.` }));
    },

    use_m365_capability: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
      const { store, awaitAsync } = deps;
      const toolName = args.tool_name as string;
      const serverHint = args.server_hint as string | undefined;
      let toolArgs: Record<string, unknown> = {};
      if (args.arguments_json) {
        if (typeof args.arguments_json === 'string') {
          try {
            toolArgs = JSON.parse(args.arguments_json as string) as Record<string, unknown>;
          } catch {
            return errorOutcome(JSON.stringify({ success: false, error: 'Invalid arguments_json — must be valid JSON.' }));
          }
        } else {
          toolArgs = args.arguments_json as Record<string, unknown>;
        }
      }
      logService.info('mcp', `M365 capability: ${toolName}`);
      store.setExpression('thinking');

      let server = serverHint ? M365_MCP_CATALOG.find((s) => s.id === serverHint) : undefined;
      if (!server) {
        server = findServerForTool(toolName);
      }
      if (!server) {
        store.setExpression('confused');
        return errorOutcome(JSON.stringify({ success: false, error: `No M365 server found with tool "${toolName}". Use list_m365_servers to see available tools.` }));
      }

      const envId = store.mcpEnvironmentId;
      if (!envId) {
        store.setExpression('confused');
        return errorOutcome(JSON.stringify({ success: false, error: 'MCP Environment ID is not configured. The admin must set it in the web part property pane under "M365 MCP Servers".' }));
      }

      const proxyConf = store.proxyConfig;
      if (!proxyConf) {
        return errorOutcome(JSON.stringify({ success: false, error: 'No proxy config available.' }));
      }

      const serverUrl = resolveServerUrl(server.id, envId);
      const serverId = server.id;
      const serverName = server.name;
      const progressData: IProgressTrackerData = {
        kind: 'progress-tracker',
        label: `Connecting to server for ${toolName}...`,
        progress: 20,
        status: 'running'
      };
      const progressBlock = createBlock('progress-tracker', `Running ${toolName}...`, progressData);
      trackCreatedBlock(store, progressBlock, deps);

      const executeM365Tool = async (): Promise<ToolRuntimeHandlerResult> => {
        const currentStore = useGrimoireStore.getState();
        const mcpClient = new McpClientService(proxyConf.proxyUrl, proxyConf.proxyApiKey);
        const personalOneDriveRequest = isPersonalOneDriveRequest(toolArgs);
        const execProgress: IProgressTrackerData = {
          kind: 'progress-tracker',
          label: `Executing ${toolName}...`,
          progress: 60,
          status: 'running',
          detail: `Preparing ${serverName}`
        };
        currentStore.updateBlock(progressBlock.id, { data: execProgress });
        const execution = await executeCatalogMcpTool({
          serverId,
          serverName,
          serverUrl,
          toolName,
          rawArgs: toolArgs,
          connections: currentStore.mcpConnections,
          getConnections: () => useGrimoireStore.getState().mcpConnections,
          mcpClient,
          sessionHelpers: {
            findExistingSession,
            connectToM365Server
          },
          getToken: currentStore.getToken,
          sourceContext: deps.sourceContext,
          taskContext: hybridInteractionEngine.getCurrentTaskContext(),
          artifacts: hybridInteractionEngine.getCurrentArtifacts(),
          currentSiteUrl: currentStore.userContext?.currentSiteUrl,
          webGroundingEnabled: !!currentStore.copilotWebGroundingEnabled,
          enrichResolvedArgs: async ({ args: resolvedArgs, targetContext, targetSource }) => {
            if (serverId !== 'mcp_ODSPRemoteServer' || toolName !== 'getFolderChildren') {
              return { args: resolvedArgs };
            }

            if (pickFirstNonEmptyString(resolvedArgs.documentLibraryId, resolvedArgs.driveId)) {
              return { args: resolvedArgs };
            }

            if (personalOneDriveRequest) {
              const personalRoot = await resolvePersonalOneDriveRootLocation(
                currentStore,
                mcpClient,
                targetContext,
                {
                  currentSiteUrl: currentStore.userContext?.currentSiteUrl
                }
              );
              return {
                args: {
                  ...resolvedArgs,
                  documentLibraryId: personalRoot.documentLibraryId
                },
                recoverySteps: [`resolved personal OneDrive document library via ${targetSource}`]
              };
            }

            const documentLibraryId = await resolveDefaultDocumentLibraryId(
              currentStore,
              mcpClient,
              targetContext,
              {
                currentSiteUrl: currentStore.userContext?.currentSiteUrl
              }
            );
            if (!documentLibraryId) {
              return { args: resolvedArgs };
            }

            return {
              args: {
                ...resolvedArgs,
                documentLibraryId
              },
              recoverySteps: [`resolved document library via ${targetSource}`]
            };
          }
        });
        const latestStore = useGrimoireStore.getState();
        latestStore.removeBlock(progressBlock.id);
        hybridInteractionEngine.onBlockRemoved(progressBlock.id);

        if (execution.success && execution.mcpResult) {
          let renderedContent = execution.mcpResult.content;
          if (personalOneDriveRequest && execution.realToolName.toLowerCase() === 'findfileorfolder') {
            const personalSiteUrl = derivePersonalOneDriveSiteUrl({
              siteUrl: execution.targetContext?.siteUrl,
              currentSiteUrl: latestStore.userContext?.currentSiteUrl,
              loginName: latestStore.userContext?.loginName,
              email: latestStore.userContext?.email
            });
            if (!personalSiteUrl) {
              return completeWithInfoCard(
                'Personal OneDrive MCP Limitation',
                'I could not derive your personal OneDrive site URL well enough to safely scope the ODSP search results, so I did not show potentially broader results.',
                deps,
                'use_m365_capability'
              );
            }

            const scopedContent = scopePersonalOneDriveSearchContent(renderedContent, personalSiteUrl);
            if (scopedContent.error || !scopedContent.content) {
              return completeWithInfoCard(
                'Personal OneDrive MCP Limitation',
                scopedContent.error || 'I could not safely scope the ODSP search results to your personal OneDrive.',
                deps,
                'use_m365_capability'
              );
            }
            renderedContent = scopedContent.content;
          }

          const summary = mapMcpResultToBlock(
            serverId,
            toolName,
            renderedContent,
            latestStore.pushBlock,
            (block) => hybridInteractionEngine.onBlockCreated(block, deps.sourceContext)
          );
          logService.debug('mcp', 'MCP execution trace', JSON.stringify({
            ...execution.trace,
            finalSummary: typeof summary === 'string' ? summary : JSON.stringify(summary)
          }));
          trackToolCompletion('use_m365_capability', '', true, 1, deps);
          logService.info('mcp', `M365 tool ${toolName} succeeded`);
          return completeOutcome(JSON.stringify({ success: true, toolName, server: serverName, summary }));
        }

        const alreadyExistsOutcome = tryHandleAlreadyExistsConflict(
          toolName,
          toolArgs,
          execution,
          deps,
          'use_m365_capability',
          { serverName, pushBlock: latestStore.pushBlock }
        );
        if (alreadyExistsOutcome) {
          return alreadyExistsOutcome;
        }

        logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));
        const errMsg = execution.error || 'Tool execution failed';
        const errBlock = createBlock('error', `MCP Error: ${toolName}`, {
          kind: 'error', message: errMsg
        } as IErrorData);
        trackCreatedBlock(latestStore, errBlock, deps);
        trackToolCompletion('use_m365_capability', '', false, 0, deps);
        hybridInteractionEngine.sendToolError(toolName, errMsg);
        logService.error('mcp', `M365 tool ${toolName} failed: ${errMsg}`);
        return errorOutcome(JSON.stringify({
          success: false,
          error: errMsg,
          realToolName: execution.realToolName,
          requiredParams: execution.requiredFields,
          schema: execution.schemaProps
        }));
      };

      const asyncResult = executeM365Tool().catch((err: Error) => {
        const currentStore = useGrimoireStore.getState();
        currentStore.removeBlock(progressBlock.id);
        hybridInteractionEngine.onBlockRemoved(progressBlock.id);
        const errBlock = createBlock('error', `MCP Error: ${toolName}`, {
          kind: 'error', message: err.message
        } as IErrorData);
        trackCreatedBlock(currentStore, errBlock, deps);
        trackToolCompletion('use_m365_capability', '', false, 0, deps);
        hybridInteractionEngine.sendToolError(toolName, err.message);
        logService.error('mcp', `M365 capability error: ${err.message}`);
        return errorOutcome(JSON.stringify({ success: false, error: err.message }));
      });

      if (awaitAsync) return asyncResult;
      return completeOutcome(JSON.stringify({ success: true, message: `Executing ${toolName} on ${serverName}... Results will appear shortly.` }));
    }
  };
}
