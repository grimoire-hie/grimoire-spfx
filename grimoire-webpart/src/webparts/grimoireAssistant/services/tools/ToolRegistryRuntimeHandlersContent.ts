import { useGrimoireStore } from '../../store/useGrimoireStore';
import type {
  IDocumentLibraryData,
  IErrorData,
  IFilePreviewData,
  IListItemsData,
  ISiteInfoData
} from '../../models/IBlock';
import { createBlock } from '../../models/IBlock';
import { ASSISTANT_SUMMARY_TARGET_TEXT } from '../../config/assistantLengthLimits';
import { logService } from '../logging/LogService';
import { CopilotChatService, type ICopilotChatReference, isM365ContextUri } from '../copilot/CopilotChatService';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type { ToolRuntimeHandler, ToolRuntimeHandlerResult } from './ToolRuntimeHandlerTypes';
import type { ContentRuntimeToolName } from './ToolRuntimeHandlerPartitions';
import { trackCreatedBlock, trackToolCompletion, trackUpdatedBlock } from './ToolRuntimeHieHelpers';
import { completeOutcome, errorOutcome } from './ToolRuntimeOutcomeHelpers';
import { extractSiteUrl, resolveUrlFromBlocks } from './ToolRuntimeUrlHelpers';
import { parseRenderHints } from './ToolRuntimeContentHelpers';

type IContentRuntimeHelpers = Record<string, unknown>;

type ReadMode = 'summarize' | 'full' | 'answer';

function normalizeTextArg(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  return trimmed || undefined;
}

function normalizeStringArrayArg(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  const cleaned: string[] = [];
  for (let i = 0; i < value.length; i++) {
    if (typeof value[i] !== 'string') continue;
    const trimmed = value[i].trim();
    if (trimmed) cleaned.push(trimmed);
  }
  return cleaned;
}

function normalizeReadMode(value: unknown, fallback: ReadMode): ReadMode {
  const mode = normalizeTextArg(value)?.toLowerCase();
  if (mode === 'summarize' || mode === 'full' || mode === 'answer') {
    return mode;
  }
  return fallback;
}

function formatReferences(refs: ICopilotChatReference[] | undefined): Array<{ title?: string; url?: string; source?: string }> {
  return (refs || []).map((r) => ({
    title: r.title,
    url: r.url,
    source: r.attributionSource || r.attributionType
  }));
}

function getModeSpecificHeading(mode: ReadMode, fallback: string): string {
  if (mode === 'summarize') return `Summary: ${fallback}`;
  if (mode === 'answer') return `Answer: ${fallback}`;
  return `Content: ${fallback}`;
}

function normalizeSummaryText(text: string): string {
  return text
    .replace(/\r/g, '')
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/__(.*?)__/g, '$1')
    .replace(/`{1,3}([^`]+)`{1,3}/g, '$1')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function buildSummarizePrompt(target: string): string {
  return `Summarize ${target} for a business user in ${ASSISTANT_SUMMARY_TARGET_TEXT}. Focus on key points, decisions, and action items. Return plain text only in the source language. Do not use markdown, headings, or introductory phrases.`;
}

export function buildContentRuntimeHandlers(
  _helpers: IContentRuntimeHelpers
): Pick<Record<RuntimeHandledToolName, ToolRuntimeHandler>, ContentRuntimeToolName> {
  return {
  browse_document_library: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, sitesService } = deps;
    const rawSiteUrl = args.site_url as string;
    const resolvedBrowseUrl = resolveUrlFromBlocks(rawSiteUrl, undefined) || rawSiteUrl;
    const siteUrl = extractSiteUrl(resolvedBrowseUrl) || resolvedBrowseUrl;
    const folderPath = args.folder_path as string | undefined;
    const renderHints = parseRenderHints(args);
    logService.info('graph', `Browse library: ${siteUrl}`);
    store.setExpression('thinking');

    if (!sitesService) {
      logService.warning('graph', 'AadHttpClient not available');
      return errorOutcome(JSON.stringify({ success: false, error: 'SharePoint connection not available. Please ensure you are signed in.' }));
    }

    const asyncResult = sitesService.browseDrive(siteUrl, folderPath).then((resp) => {
      const currentStore = useGrimoireStore.getState();
      if (resp.success && resp.data) {
        const libData: IDocumentLibraryData = {
          kind: 'document-library',
          siteName: siteUrl,
          libraryName: 'Documents',
          items: resp.data.map((item) => ({
            name: item.name,
            type: item.type,
            url: item.url,
            documentLibraryId: item.documentLibraryId,
            fileOrFolderId: item.fileOrFolderId,
            size: item.size,
            lastModified: item.lastModified,
            author: item.author,
            fileType: item.fileType
          })),
          breadcrumb: folderPath ? folderPath.split('/') : []
        };
        const libBlock = createBlock('document-library', 'Document Library', libData, true, renderHints);
        trackCreatedBlock(currentStore, libBlock, deps);
        trackToolCompletion('browse_document_library', libBlock.id, true, libData.items.length, deps);
        return completeOutcome(JSON.stringify({ success: true, itemCount: libData.items.length, items: libData.items.map((i) => ({ name: i.name, type: i.type })) }));
      }
      trackToolCompletion('browse_document_library', '', false, 0, deps);
      logService.error('graph', `Browse failed: ${resp.error}`);
      return errorOutcome(JSON.stringify({ success: false, error: resp.error }));
    }).catch((err: Error) => {
      trackToolCompletion('browse_document_library', '', false, 0, deps);
      logService.error('graph', `Browse error: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Browsing document library at ${siteUrl}...` }));
  },

  show_file_details: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, sitesService } = deps;
    const rawFileUrl = args.file_url as string;
    const fileUrl = resolveUrlFromBlocks(rawFileUrl, args.file_name as string | undefined) || rawFileUrl;
    const renderHints = parseRenderHints(args);
    logService.info('graph', `File details: ${fileUrl}`);
    store.setExpression('thinking');

    try { window.open(fileUrl, '_blank', 'noopener,noreferrer'); } catch { /* popup blocked */ }

    let fileName = 'File';
    try { fileName = decodeURIComponent(fileUrl.split('/').pop() || 'File'); } catch { /* keep default */ }
    const fileType = fileName.includes('.') ? fileName.split('.').pop() || '' : '';

    const fileData: IFilePreviewData = {
      kind: 'file-preview',
      fileName,
      fileUrl,
      fileType,
      metadata: {}
    };
    const fileBlock = createBlock('file-preview', fileName, fileData, true, renderHints);
    trackCreatedBlock(store, fileBlock, deps);

    let asyncResult: Promise<ToolRuntimeHandlerResult> | undefined;
    if (sitesService) {
      asyncResult = sitesService.getFileDetails(fileUrl).then((resp) => {
        const currentStore = useGrimoireStore.getState();
        if (resp.success && resp.data) {
          const updatedData: IFilePreviewData = {
            kind: 'file-preview',
            fileName: resp.data.fileName,
            fileUrl: resp.data.fileUrl,
            fileType: resp.data.fileType,
            size: resp.data.size,
            lastModified: resp.data.lastModified,
            author: resp.data.author,
            metadata: resp.data.metadata
          };
          trackUpdatedBlock(currentStore, fileBlock.id, { data: updatedData }, { ...fileBlock, data: updatedData }, deps);
          trackToolCompletion('show_file_details', fileBlock.id, true, 1, deps);
          return completeOutcome(JSON.stringify({
            success: true,
            displayed: true,
            opened: true,
            fileName: resp.data.fileName,
            fileType: resp.data.fileType,
            size: resp.data.size,
            author: resp.data.author,
            lastModified: resp.data.lastModified,
            note: 'A file preview card with full metadata is already displayed to the user AND the file has been opened in a new browser tab. Do NOT create a markdown or info card — the user can already see all details.'
          }));
        }
        trackToolCompletion('show_file_details', fileBlock.id, false, 0, deps);
        logService.warning('graph', `File details fetch failed: ${resp.error}`);
        return errorOutcome(JSON.stringify({ success: false, error: resp.error }));
      }).catch((err: Error) => {
        trackToolCompletion('show_file_details', fileBlock.id, false, 0, deps);
        logService.warning('graph', `File details error: ${err.message}`);
        return errorOutcome(JSON.stringify({ success: false, error: err.message }));
      });
    }

    if (awaitAsync && asyncResult) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, displayed: true, opened: true, fileName, fileUrl, note: 'File preview card is displayed and file opened in new tab. Do NOT create additional cards.' }));
  },

  show_site_info: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, sitesService } = deps;
    const siteUrl = args.site_url as string;
    logService.info('graph', `Site info: ${siteUrl}`);
    store.setExpression('thinking');

    if (!sitesService) {
      logService.warning('graph', 'AadHttpClient not available');
      return errorOutcome(JSON.stringify({ success: false, error: 'SharePoint connection not available. Please ensure you are signed in.' }));
    }

    const asyncResult = sitesService.getSiteInfo(siteUrl).then((resp) => {
      const currentStore = useGrimoireStore.getState();
      if (resp.success && resp.data) {
        const siteData: ISiteInfoData = {
          kind: 'site-info',
          siteName: resp.data.siteName,
          siteUrl: resp.data.siteUrl,
          description: resp.data.description,
          created: resp.data.created,
          lastModified: resp.data.lastModified,
          libraries: resp.data.libraries,
          lists: resp.data.lists
        };
        const siteBlock = createBlock('site-info', resp.data.siteName, siteData);
        trackCreatedBlock(currentStore, siteBlock, deps);
        trackToolCompletion('show_site_info', siteBlock.id, true, 1, deps);
        return completeOutcome(JSON.stringify({
          success: true,
          siteName: resp.data.siteName,
          description: resp.data.description,
          libraries: resp.data.libraries?.length,
          lists: resp.data.lists?.length
        }));
      }
      trackToolCompletion('show_site_info', '', false, 0, deps);
      logService.error('graph', `Site info failed: ${resp.error}`);
      return errorOutcome(JSON.stringify({ success: false, error: resp.error }));
    }).catch((err: Error) => {
      trackToolCompletion('show_site_info', '', false, 0, deps);
      logService.error('graph', `Site info error: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Loading site info for ${siteUrl}...` }));
  },

  show_list_items: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, sitesService } = deps;
    const listName = args.list_name as string;
    const siteUrl = args.site_url as string;
    const renderHints = parseRenderHints(args);
    logService.info('graph', `List items: ${listName}`);
    store.setExpression('thinking');

    if (!sitesService || !siteUrl) {
      const reason = siteUrl ? 'SharePoint connection not available' : 'Missing site_url parameter';
      logService.warning('graph', reason);
      return errorOutcome(JSON.stringify({ success: false, error: `${reason}. Please ensure you are signed in.` }));
    }

    const asyncResult = sitesService.getListItems(siteUrl, listName).then((resp) => {
      const currentStore = useGrimoireStore.getState();
      if (resp.success && resp.data) {
        const listData: IListItemsData = {
          kind: 'list-items',
          listName,
          columns: resp.data.columns,
          items: resp.data.items,
          totalCount: resp.data.totalCount
        };
        const listBlock = createBlock('list-items', listName, listData, true, renderHints);
        trackCreatedBlock(currentStore, listBlock, deps);
        trackToolCompletion('show_list_items', listBlock.id, true, listData.totalCount, deps);
        return completeOutcome(JSON.stringify({ success: true, listName, totalCount: listData.totalCount, columns: listData.columns?.slice(0, 5) }));
      }
      trackToolCompletion('show_list_items', '', false, 0, deps);
      logService.error('graph', `List items failed: ${resp.error}`);
      return errorOutcome(JSON.stringify({ success: false, error: resp.error }));
    }).catch((err: Error) => {
      trackToolCompletion('show_list_items', '', false, 0, deps);
      logService.error('graph', `List items error: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Loading items from "${listName}"...` }));
  },

  read_file_content: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, aadClient } = deps;
    const rawFileUrl = normalizeTextArg(args.file_url);
    const rawFileUrls = normalizeStringArrayArg(args.file_urls);
    const fileName = normalizeTextArg(args.file_name);
    const mode = normalizeReadMode(args.mode, 'full');
    const question = normalizeTextArg(args.question);
    const resolvedUrls: string[] = [];

    for (let i = 0; i < rawFileUrls.length; i++) {
      const resolved = resolveUrlFromBlocks(rawFileUrls[i], undefined) || rawFileUrls[i];
      const trimmed = resolved.trim();
      if (trimmed) resolvedUrls.push(trimmed);
    }
    if (resolvedUrls.length === 0 && rawFileUrl) {
      const resolvedSingle = (resolveUrlFromBlocks(rawFileUrl, fileName) || rawFileUrl).trim();
      if (resolvedSingle) resolvedUrls.push(resolvedSingle);
    }

    const fileUris: string[] = [];
    const dedupe = new Set<string>();
    for (let i = 0; i < resolvedUrls.length; i++) {
      const key = resolvedUrls[i].toLowerCase();
      if (dedupe.has(key)) continue;
      dedupe.add(key);
      fileUris.push(resolvedUrls[i]);
    }

    if (fileUris.length === 0) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'read_file_content requires "file_url" (single file) or "file_urls" (multiple files).'
      }));
    }

    logService.info('graph', `Read file content via Copilot (${mode}): ${fileUris.length} file(s)`);
    store.setExpression('thinking');

    if (!aadClient) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'Microsoft Graph connection not available. Please ensure you are signed in.'
      }));
    }

    const executeReadFile = async (): Promise<ToolRuntimeHandlerResult> => {
      for (let i = 0; i < fileUris.length; i++) {
        if (!isM365ContextUri(fileUris[i])) {
          throw new Error('Only SharePoint/OneDrive URLs are supported for Copilot file processing.');
        }
      }
      if (mode === 'full' && fileUris.length > 1) {
        throw new Error('read_file_content with mode "full" supports a single file only. Use mode "summarize" or "answer" for multiple files.');
      }
      if (mode === 'answer' && !question) {
        throw new Error('read_file_content with mode "answer" requires the "question" parameter.');
      }

      const primaryFileUrl = fileUris[0];
      const displayName = fileUris.length > 1
        ? `${fileUris.length} files`
        : (fileName || decodeURIComponent(primaryFileUrl.split('/').pop() || '') || 'Document');
      let prompt: string;
      if (mode === 'summarize') {
        if (fileUris.length > 1) {
          prompt = `Summarize and compare these files for a business user in ${ASSISTANT_SUMMARY_TARGET_TEXT}. Focus on key similarities, differences, decisions, and action items. Return plain text only in the source language. Do not use markdown, headings, or introductory phrases.`;
        } else {
          prompt = buildSummarizePrompt('this file');
        }
      } else if (mode === 'answer') {
        prompt = fileUris.length > 1
          ? `Answer this question using all provided file contents: ${question}. When files differ, call out the differences clearly.`
          : `Answer this question using the file content: ${question}`;
      } else {
        prompt = 'Retrieve the complete file content, preserving headings, structure, and tabular information when present.';
      }

      const copilot = new CopilotChatService(aadClient);
      const result = await copilot.chat({
        prompt,
        fileUris,
        enableWebGrounding: !!store.copilotWebGroundingEnabled
      });
      if (!result.success || !result.text) {
        throw new Error(result.error?.message || 'Copilot file processing failed');
      }

      const heading = getModeSpecificHeading(mode, displayName);
      const references = formatReferences(result.references);
      const content = mode === 'summarize' ? normalizeSummaryText(result.text) : result.text.trim();
      const note = mode === 'summarize'
        ? `Copilot already generated the final summary text in "content". Call show_info_card with heading "${heading}" and body equal to "content". Do NOT re-summarize, do NOT call read_file_content again, and keep chat acknowledgment brief.`
        : `Copilot already generated the requested ${mode === 'answer' ? 'answer' : 'content'} in "content". Use it directly without extra retrieval calls.`;

      trackToolCompletion('read_file_content', '', true, 1, deps);
      return completeOutcome(JSON.stringify({
        success: true,
        fileName: displayName,
        fileUrl: primaryFileUrl,
        fileUrls: fileUris,
        fileCount: fileUris.length,
        mode,
        contentReadable: true,
        content,
        references,
        conversationId: result.conversationId,
        note
      }));
    };

    const asyncResult = executeReadFile().catch((err: Error) => {
      const errBlock = createBlock('error', 'File Processing Error', {
        kind: 'error',
        message: err.message
      } as IErrorData);
      trackCreatedBlock(useGrimoireStore.getState(), errBlock, deps);
      trackToolCompletion('read_file_content', '', false, 0, deps);
      logService.error('graph', `Read file content error: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: 'Processing file with Copilot...' }));
  },

  read_email_content: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, aadClient } = deps;
    const emailSubject = normalizeTextArg(args.subject);
    const emailSender = normalizeTextArg(args.sender);
    const emailDateHint = normalizeTextArg(args.date_hint);
    const mode = normalizeReadMode(args.mode, 'full');
    const emailQuestion = normalizeTextArg(args.question);

    if (!emailSubject && !emailSender && !emailDateHint) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'read_email_content requires at least one identifier: subject, sender, or date_hint.'
      }));
    }
    if (mode === 'answer' && !emailQuestion) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'read_email_content with mode "answer" requires the "question" parameter.'
      }));
    }
    if (!aadClient) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'Microsoft Graph connection not available. Please ensure you are signed in.'
      }));
    }

    logService.info('graph', `Read email content via Copilot (${mode}): "${emailSubject || '(no subject)'}"`);
    store.setExpression('thinking');

    const executeReadEmail = async (): Promise<ToolRuntimeHandlerResult> => {
      let target = 'the email';
      if (emailSubject) target += ` with subject "${emailSubject}"`;
      if (emailSender) target += `${emailSubject ? ' from' : ' from'} ${emailSender}`;
      if (emailDateHint) target += ` around ${emailDateHint}`;

      let prompt: string;
      if (mode === 'summarize') {
        prompt = `${buildSummarizePrompt(target)} Include sender, recipients, sent time, and key decisions/action items.`;
      } else if (mode === 'answer') {
        prompt = `Using ${target}, answer this question: ${emailQuestion}`;
      } else {
        prompt = `Retrieve the complete content of ${target}. Include full body text, recipients, sent date, and listed attachments.`;
      }

      const copilot = new CopilotChatService(aadClient);
      const result = await copilot.chat({
        prompt,
        enableWebGrounding: !!store.copilotWebGroundingEnabled
      });
      if (!result.success || !result.text) {
        throw new Error(result.error?.message || 'Copilot email processing failed');
      }

      const heading = getModeSpecificHeading(mode, emailSubject || 'Email');
      const references = formatReferences(result.references);
      const content = mode === 'summarize' ? normalizeSummaryText(result.text) : result.text.trim();
      const note = mode === 'summarize'
        ? `Copilot already generated the final summary text in "content". Call show_info_card with heading "${heading}" and body equal to "content". Do NOT re-summarize and keep chat acknowledgment brief.`
        : `Copilot already generated the requested ${mode === 'answer' ? 'answer' : 'content'} in "content". Use it directly without extra retrieval calls.`;

      trackToolCompletion('read_email_content', '', true, 1, deps);
      return completeOutcome(JSON.stringify({
        success: true,
        subject: emailSubject || '',
        sender: emailSender || '',
        mode,
        contentReadable: true,
        content,
        references,
        conversationId: result.conversationId,
        note
      }));
    };

    const emailAsyncResult = executeReadEmail().catch((err: Error) => {
      const errBlock = createBlock('error', 'Email Processing Error', {
        kind: 'error',
        message: err.message
      } as IErrorData);
      trackCreatedBlock(useGrimoireStore.getState(), errBlock, deps);
      logService.error('graph', `Read email content error: ${err.message}`);
      trackToolCompletion('read_email_content', '', false, 0, deps);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return emailAsyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: 'Processing email with Copilot...' }));
  },

  read_teams_messages: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, aadClient } = deps;
    const chatOrChannel = normalizeTextArg(args.chat_or_channel);
    const teamsTopic = normalizeTextArg(args.topic);
    const teamsQuestion = normalizeTextArg(args.question);
    const mode = normalizeReadMode(args.mode, 'full');

    if (!chatOrChannel) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'read_teams_messages requires chat_or_channel.'
      }));
    }
    if (mode === 'answer' && !teamsQuestion) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'read_teams_messages with mode "answer" requires the "question" parameter.'
      }));
    }
    if (!aadClient) {
      return errorOutcome(JSON.stringify({
        success: false,
        error: 'Microsoft Graph connection not available. Please ensure you are signed in.'
      }));
    }

    logService.info('graph', `Read Teams messages via Copilot (${mode}): "${chatOrChannel}"`);
    store.setExpression('thinking');

    const executeReadTeams = async (): Promise<ToolRuntimeHandlerResult> => {
      let target = `the Teams chat/channel "${chatOrChannel}"`;
      if (teamsTopic) target += ` with focus on "${teamsTopic}"`;

      let prompt: string;
      if (mode === 'summarize') {
        prompt = `${buildSummarizePrompt(`recent messages from ${target}`)} Include key participants, major points, and action items.`;
      } else if (mode === 'answer') {
        prompt = `Using recent messages from ${target}, answer this question: ${teamsQuestion}`;
      } else {
        prompt = `Retrieve recent messages from ${target}. Include senders, timestamps, and message content.`;
      }

      const copilot = new CopilotChatService(aadClient);
      const result = await copilot.chat({
        prompt,
        enableWebGrounding: !!store.copilotWebGroundingEnabled
      });
      if (!result.success || !result.text) {
        throw new Error(result.error?.message || 'Copilot Teams processing failed');
      }

      const heading = getModeSpecificHeading(mode, chatOrChannel);
      const references = formatReferences(result.references);
      const content = mode === 'summarize' ? normalizeSummaryText(result.text) : result.text.trim();
      const note = mode === 'summarize'
        ? `Copilot already generated the final summary text in "content". Call show_info_card with heading "${heading}" and body equal to "content". Do NOT re-summarize and keep chat acknowledgment brief.`
        : `Copilot already generated the requested ${mode === 'answer' ? 'answer' : 'content'} in "content". Use it directly without extra retrieval calls.`;

      trackToolCompletion('read_teams_messages', '', true, 1, deps);
      return completeOutcome(JSON.stringify({
        success: true,
        chatOrChannel,
        mode,
        contentReadable: true,
        content,
        references,
        conversationId: result.conversationId,
        note
      }));
    };

    const teamsAsyncResult = executeReadTeams().catch((err: Error) => {
      const errBlock = createBlock('error', 'Teams Processing Error', {
        kind: 'error',
        message: err.message
      } as IErrorData);
      trackCreatedBlock(useGrimoireStore.getState(), errBlock, deps);
      logService.error('graph', `Read Teams messages error: ${err.message}`);
      trackToolCompletion('read_teams_messages', '', false, 0, deps);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return teamsAsyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: 'Processing Teams messages with Copilot...' }));
  },
  };
}
