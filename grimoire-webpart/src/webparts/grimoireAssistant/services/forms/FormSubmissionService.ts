/**
 * FormSubmissionService — Executes MCP tools from form data.
 * Reuses the same MCP infrastructure as use_m365_capability.
 */

import type { IFormData } from '../../models/IBlock';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { McpClientService } from '../mcp/McpClientService';
import { logService } from '../logging/LogService';
import { getCatalogEntry, resolveServerUrl } from '../../models/McpServerCatalog';
import { findExistingSession, connectToM365Server } from '../tools/ToolRuntimeSharedHelpers';
import type { IFunctionCallStore } from '../tools/ToolRuntimeContracts';
import { extractMcpTextParts } from '../mcp/mcpUtils';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { createCorrelationId } from '../hie/HAEContracts';
import type { IHieSourceContext } from '../hie/HIETypes';
import {
  executeCatalogMcpTool,
  extractStructuredMcpPayload
} from '../mcp/McpExecutionAdapter';
import type { IMcpTargetContext } from '../mcp/McpTargetContext';
import {
  pickFirstNonEmptyString,
  resolveDefaultDocumentLibraryId,
  resolveSiteIdForTargetContext
} from '../mcp/OdspLocationResolver';
import { mapMcpResultToBlock } from '../mcp/McpResultMapper';
import { buildDocxBase64FromText } from './DocxBuilder';
import { sanitizeCreatedFileName } from '../tools/FormCreateHelpers';

export interface IFormSubmissionResult {
  success: boolean;
  message: string;
  content?: unknown;
}

function setNestedArgValue(target: Record<string, unknown>, path: string, value: unknown): void {
  const segments = path
    .split('.')
    .map((segment) => segment.trim())
    .filter((segment) => segment.length > 0);

  if (segments.length === 0) {
    return;
  }

  if (segments.length === 1) {
    target[segments[0]] = value;
    return;
  }

  let cursor: Record<string, unknown> = target;
  for (let i = 0; i < segments.length - 1; i++) {
    const segment = segments[i];
    const existing = cursor[segment];
    if (!existing || typeof existing !== 'object' || Array.isArray(existing)) {
      cursor[segment] = {};
    }
    cursor = cursor[segment] as Record<string, unknown>;
  }

  cursor[segments[segments.length - 1]] = value;
}

function ensureCreatedFileName(fileName: string | undefined, extension: string): string {
  return sanitizeCreatedFileName(fileName, extension);
}

async function enrichCreateLocationArgs(
  toolName: string,
  serverId: string,
  mcpArgs: Record<string, unknown>,
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  targetContext?: IMcpTargetContext
): Promise<Record<string, unknown>> {
  if (serverId !== 'mcp_ODSPRemoteServer') {
    return mcpArgs;
  }

  if (!['createFolder', 'createSmallTextFile', 'createSmallBinaryFile'].includes(toolName)) {
    return mcpArgs;
  }

  if (pickFirstNonEmptyString(mcpArgs.documentLibraryId, mcpArgs.driveId)) {
    return mcpArgs;
  }

  const documentLibraryId = await resolveDefaultDocumentLibraryId(store, mcpClient, targetContext);
  if (!documentLibraryId) {
    return mcpArgs;
  }

  logService.info('mcp', `Form submission auto-resolved documentLibraryId for ${toolName}`);
  return {
    ...mcpArgs,
    documentLibraryId
  };
}

async function enrichSharePointListArgs(
  toolName: string,
  serverId: string,
  mcpArgs: Record<string, unknown>,
  store: IFunctionCallStore,
  mcpClient: McpClientService,
  targetContext?: IMcpTargetContext,
  originalArgs?: Record<string, unknown>
): Promise<Record<string, unknown>> {
  if (serverId !== 'mcp_SharePointListsTools') {
    return mcpArgs;
  }

  const currentSiteId = pickFirstNonEmptyString(mcpArgs.siteId);
  const originalSiteId = pickFirstNonEmptyString(originalArgs?.siteId);
  if (currentSiteId && !originalSiteId) {
    return mcpArgs;
  }

  const resolvedSiteId = await resolveSiteIdForTargetContext(store, mcpClient, targetContext);
  if (!resolvedSiteId) {
    return mcpArgs;
  }

  if (currentSiteId === resolvedSiteId) {
    return mcpArgs;
  }

  logService.info('mcp', `Form submission refreshed siteId for ${toolName}`);
  return {
    ...mcpArgs,
    siteId: resolvedSiteId
  };
}

function emitFormSubmissionEvent(
  eventName: string,
  correlationId: string,
  formData: IFormData,
  submissionTarget: IFormData['submissionTarget'],
  message: string | undefined,
  sourceContext: IHieSourceContext | undefined,
  success?: boolean
): void {
  hybridInteractionEngine.emitEvent({
    eventName,
    source: 'form',
    surface: 'action-panel',
    correlationId,
    payload: {
      preset: formData.preset,
      formDescription: formData.description,
      toolName: submissionTarget.toolName,
      serverId: submissionTarget.serverId,
      message,
      success,
      sourceBlockId: sourceContext?.sourceBlockId,
      sourceBlockType: sourceContext?.sourceBlockType,
      sourceBlockTitle: sourceContext?.sourceBlockTitle,
      sourceArtifactId: sourceContext?.sourceArtifactId,
      sourceTaskKind: sourceContext?.sourceTaskKind,
      sourceEventName: sourceContext?.sourceEventName,
      sourceCorrelationId: sourceContext?.correlationId,
      sourceTurnId: sourceContext?.sourceTurnId,
      sourceRootTurnId: sourceContext?.sourceRootTurnId,
      sourceParentTurnId: sourceContext?.sourceParentTurnId,
      selectedItems: sourceContext?.selectedItems,
      targetContext: submissionTarget.targetContext || sourceContext?.targetContext
    },
    exposurePolicy: { mode: 'store-only', relevance: 'contextual' },
    turnId: sourceContext?.sourceTurnId,
    rootTurnId: sourceContext?.sourceRootTurnId,
    parentTurnId: sourceContext?.sourceParentTurnId
  });
}

function prepareSubmissionArgs(
  formData: IFormData,
  toolName: string,
  rawArgs: Record<string, unknown>
): Record<string, unknown> {
  if (formData.preset !== 'word-document-create' || toolName !== 'createSmallBinaryFile') {
    return rawArgs;
  }

  if (typeof rawArgs.base64Content === 'string' && rawArgs.base64Content.trim()) {
    return rawArgs;
  }

  const content = pickFirstNonEmptyString(rawArgs.contentText, rawArgs.contentInHtml, rawArgs.content);
  const preparedArgs: Record<string, unknown> = {
    ...rawArgs,
    filename: ensureCreatedFileName(pickFirstNonEmptyString(rawArgs.filename), 'docx'),
    base64Content: buildDocxBase64FromText(content || '')
  };

  delete preparedArgs.contentText;
  delete preparedArgs.contentInHtml;
  delete preparedArgs.content;

  return preparedArgs;
}

function getSharePointColumnTypeLabel(record: Record<string, unknown>): string | undefined {
  const textSettings = record.text;
  if (textSettings && typeof textSettings === 'object' && !Array.isArray(textSettings)) {
    return (textSettings as Record<string, unknown>).allowMultipleLines === true
      ? 'Multiple lines of text'
      : 'Text';
  }
  if (record.number && typeof record.number === 'object') return 'Number';
  if (record.choice && typeof record.choice === 'object') return 'Choice';
  if (record.boolean && typeof record.boolean === 'object') return 'Yes/No';
  if (record.dateTime && typeof record.dateTime === 'object') return 'Date and time';
  if (record.personOrGroup && typeof record.personOrGroup === 'object') return 'Person or Group';
  if (record.hyperlinkOrPicture && typeof record.hyperlinkOrPicture === 'object') return 'Link or Picture';
  if (record.lookup && typeof record.lookup === 'object') return 'Lookup';
  return undefined;
}

function shouldRenderSuccessResultBlock(serverId: string, toolName: string): boolean {
  if (serverId !== 'mcp_SharePointListsTools') {
    return false;
  }

  const normalizedToolName = toolName.toLowerCase();
  return normalizedToolName === 'createlistcolumn' || normalizedToolName === 'editlistcolumn';
}

/**
 * Execute the MCP tool targeted by a form submission.
 */
export async function executeFormSubmission(
  formData: IFormData,
  fieldValues: Record<string, string>,
  emailTags: Record<string, string[]>,
  store: IFunctionCallStore,
  sourceContext?: IHieSourceContext
): Promise<IFormSubmissionResult> {
  const { submissionTarget } = formData;
  const { toolName, serverId, staticArgs, fieldToParamMap, targetContext: explicitTargetContext } = submissionTarget;
  const submissionCorrelationId = createCorrelationId('formsubmit');

  if (!toolName || !serverId) {
    emitFormSubmissionEvent(
      'form.execution.failed',
      submissionCorrelationId,
      formData,
      submissionTarget,
      'Form submission target is not configured.',
      sourceContext,
      false
    );
    return { success: false, message: 'Form submission target is not configured.' };
  }

  emitFormSubmissionEvent(
    'form.execution.started',
    submissionCorrelationId,
    formData,
    submissionTarget,
    undefined,
    sourceContext
  );

  // Build MCP arguments from form values
  const mcpArgs: Record<string, unknown> = { ...staticArgs };
  const fieldKeys = Object.keys(fieldValues);
  for (let i = 0; i < fieldKeys.length; i++) {
    const fieldKey = fieldKeys[i];
    const paramName = (fieldToParamMap && fieldToParamMap[fieldKey]) || fieldKey;
    const raw = fieldValues[fieldKey];

    // Find the field definition to determine type
    const fieldDef = formData.fields.find((f) => f.key === fieldKey);
    if (!fieldDef) {
      // Include hidden or unmapped values directly
      if (raw !== undefined && raw !== '') {
        setNestedArgValue(mcpArgs, paramName, raw);
      }
      continue;
    }

    // Type coercion based on field type
    switch (fieldDef.type) {
      case 'email-list':
      case 'people-picker': {
        const tags = emailTags[fieldKey] || [];
        if (tags.length > 0) {
          // Send as native array — MCP tools expect array type, not JSON string
          setNestedArgValue(mcpArgs, paramName, tags);
        }
        break;
      }
      case 'toggle':
        setNestedArgValue(mcpArgs, paramName, raw === 'true');
        break;
      case 'number':
        if (raw !== '') {
          setNestedArgValue(mcpArgs, paramName, Number(raw));
        }
        break;
      case 'hidden':
        if (raw !== '') {
          setNestedArgValue(mcpArgs, paramName, raw);
        }
        break;
      default:
        if (raw !== '') {
          setNestedArgValue(mcpArgs, paramName, raw);
        }
        break;
    }
  }

  const preparedArgs = prepareSubmissionArgs(formData, toolName, mcpArgs);

  logService.info('mcp', `Form submission: ${toolName} on ${serverId}`);

  // Resolve server
  const server = getCatalogEntry(serverId);
  if (!server) {
    return { success: false, message: `Server "${serverId}" not found in M365 catalog.` };
  }

  const envId = store.mcpEnvironmentId;
  if (!envId) {
    return { success: false, message: 'MCP Environment ID is not configured.' };
  }

  const proxyConf = store.proxyConfig;
  if (!proxyConf) {
    return { success: false, message: 'No proxy config available.' };
  }

  const serverUrl = resolveServerUrl(serverId, envId);
  const mcpClient = new McpClientService(proxyConf.proxyUrl, proxyConf.proxyApiKey);

  try {
    const currentStore = useGrimoireStore.getState();
    const execution = await executeCatalogMcpTool({
      serverId,
      serverName: server.name,
      serverUrl,
      toolName,
      rawArgs: preparedArgs,
      connections: currentStore.mcpConnections,
      getConnections: () => useGrimoireStore.getState().mcpConnections,
      mcpClient,
      sessionHelpers: {
        findExistingSession,
        connectToM365Server
      },
      getToken: currentStore.getToken,
      explicitTargetContext,
      sourceContext,
      taskContext: hybridInteractionEngine.getCurrentTaskContext(),
      artifacts: hybridInteractionEngine.getCurrentArtifacts(),
      currentSiteUrl: currentStore.userContext?.currentSiteUrl,
      enrichResolvedArgs: async ({ args, targetContext, targetSource }) => {
        const sharePointArgs = await enrichSharePointListArgs(
          toolName,
          serverId,
          args,
          store,
          mcpClient,
          targetContext,
          preparedArgs
        );
        const enrichedArgs = await enrichCreateLocationArgs(
          toolName,
          serverId,
          sharePointArgs,
          store,
          mcpClient,
          targetContext
        );
        const extraRecoverySteps: string[] = [];
        if (
          pickFirstNonEmptyString(sharePointArgs.siteId)
          && pickFirstNonEmptyString(args.siteId) !== pickFirstNonEmptyString(sharePointArgs.siteId)
        ) {
          extraRecoverySteps.push(`resolved siteId via ${targetSource}`);
        }
        if (
          !pickFirstNonEmptyString(sharePointArgs.documentLibraryId, sharePointArgs.driveId)
          && pickFirstNonEmptyString(enrichedArgs.documentLibraryId, enrichedArgs.driveId)
        ) {
          extraRecoverySteps.push(`resolved document library via ${targetSource}`);
        }
        return {
          args: enrichedArgs,
          recoverySteps: extraRecoverySteps
        };
      }
    });

    if (execution.success && execution.mcpResult) {
      const successMsg = extractSuccessMessage(
        toolName,
        execution.mcpResult.content,
        {
          ...preparedArgs,
          ...execution.resolvedArgs
        }
      ) || 'Operation completed successfully.';
      if (shouldRenderSuccessResultBlock(serverId, toolName)) {
        mapMcpResultToBlock(
          serverId,
          toolName,
          execution.mcpResult.content,
          store.pushBlock,
          (block) => hybridInteractionEngine.onBlockCreated?.(block, sourceContext)
        );
      }
      logService.debug('mcp', 'MCP execution trace', JSON.stringify({
        ...execution.trace,
        finalSummary: successMsg
      }));
      logService.info('mcp', `Form submission ${toolName} succeeded`);
      emitFormSubmissionEvent(
        'form.execution.completed',
        submissionCorrelationId,
        formData,
        submissionTarget,
        successMsg,
        sourceContext,
        true
      );
      return { success: true, message: successMsg, content: execution.mcpResult.content };
    }

    const errMsg = execution.error || 'Tool execution failed';
    logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));
    logService.warning('mcp', `Form submission ${toolName} failed: ${errMsg}`);
    emitFormSubmissionEvent(
      'form.execution.failed',
      submissionCorrelationId,
      formData,
      submissionTarget,
      errMsg,
      sourceContext,
      false
    );
    return { success: false, message: errMsg };
  } catch (err) {
    const errMsg = (err as Error).message || 'Unknown error';
    logService.error('mcp', `Form submission error: ${errMsg}`);
    emitFormSubmissionEvent(
      'form.execution.failed',
      submissionCorrelationId,
      formData,
      submissionTarget,
      errMsg,
      sourceContext,
      false
    );
    return { success: false, message: errMsg };
  }
}

/**
 * Extract a human-readable message from an MCP response.
 * Looks for a `message` field in the JSON text parts.
 */
export function extractSuccessMessage(
  toolName: string,
  content: Array<{ type: string; text?: string }>,
  submittedArgs?: Record<string, unknown>
): string | undefined {
  const payload = extractStructuredMcpPayload(content).payload;
  if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
    const record = payload as Record<string, unknown>;
    const genericMessage = typeof record.message === 'string' ? record.message.trim() : undefined;
    const reply = typeof record.reply === 'string' ? record.reply.trim() : undefined;
    const name = pickFirstNonEmptyString(record.name, record.displayName, record.title);
    const webUrl = pickFirstNonEmptyString(record.webUrl, record.url);

    if (toolName === 'createFolder') {
      const folderName = name || pickFirstNonEmptyString(submittedArgs?.folderName);
      if (folderName && webUrl) {
        return `Folder "${folderName}" created: ${webUrl}`;
      }
      if (folderName) {
        return `Folder "${folderName}" created successfully.`;
      }
    }

    if (toolName === 'createSmallTextFile' || toolName === 'createSmallBinaryFile') {
      const fileName = name || pickFirstNonEmptyString(submittedArgs?.filename);
      if (fileName && webUrl) {
        return `File "${fileName}" created: ${webUrl}`;
      }
      if (fileName) {
        return `File "${fileName}" created successfully.`;
      }
    }

    if (toolName === 'renameFileOrFolder') {
      const renamedTo = name || pickFirstNonEmptyString(submittedArgs?.newFileOrFolderName);
      const renamedFrom = pickFirstNonEmptyString(submittedArgs?.oldFileOrFolderName);
      if (renamedFrom && renamedTo && renamedFrom !== renamedTo) {
        return webUrl
          ? `Renamed "${renamedFrom}" to "${renamedTo}": ${webUrl}`
          : `Renamed "${renamedFrom}" to "${renamedTo}".`;
      }
      if (renamedTo) {
        return webUrl
          ? `Renamed to "${renamedTo}": ${webUrl}`
          : `Renamed to "${renamedTo}".`;
      }
    }

    if (toolName === 'createListColumn') {
      const columnName = name || pickFirstNonEmptyString(submittedArgs?.displayName, submittedArgs?.name);
      const listName = pickFirstNonEmptyString(submittedArgs?.listName);
      const typeLabel = getSharePointColumnTypeLabel(record);
      if (columnName && listName && typeLabel) {
        return `Added ${typeLabel} column "${columnName}" to list "${listName}".`;
      }
      if (columnName && listName) {
        return `Added column "${columnName}" to list "${listName}".`;
      }
      if (columnName && typeLabel) {
        return `Added ${typeLabel} column "${columnName}".`;
      }
      if (columnName) {
        return `Added column "${columnName}".`;
      }
    }

    if (toolName === 'editListColumn') {
      const columnName = name || pickFirstNonEmptyString(
        submittedArgs?.displayName,
        submittedArgs?.name,
        submittedArgs?.columnId
      );
      if (columnName) {
        return `Updated column "${columnName}".`;
      }
    }

    if (genericMessage && !/^graph tool executed successfully\.?$/i.test(genericMessage)) {
      return genericMessage;
    }
    if (reply) {
      return reply;
    }
  }

  const parts = extractMcpTextParts(content);
  for (let i = 0; i < parts.length; i++) {
    try {
      const parsed = JSON.parse(parts[i]);
      if (parsed && typeof parsed.message === 'string' && parsed.message && !/^graph tool executed successfully\.?$/i.test(parsed.message)) {
        return parsed.message;
      }
      if (parsed && typeof parsed.reply === 'string' && parsed.reply) {
        return parsed.reply;
      }
    } catch { /* not JSON — skip */ }
  }
  return undefined;
}
