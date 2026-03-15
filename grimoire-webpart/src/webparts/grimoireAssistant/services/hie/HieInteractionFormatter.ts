import { serializeUntrustedData } from '../realtime/PromptSafety';
import { ASSISTANT_SUMMARY_TARGET_TEXT } from '../../config/assistantLengthLimits';
import type { IBlockInteraction } from './HIETypes';
import type { BlockTracker } from './BlockTracker';
import type { IHiePromptMessage } from './HiePromptProtocol';

const DOCUMENT_DATA_FIELDS = ['title', 'name', 'fileName', 'fileType', 'url', 'author', 'date', 'site', 'summary'] as const;
const EMAIL_DATA_FIELDS = ['Subject', 'title', 'From', 'itemId'] as const;
const PERSON_DATA_FIELDS = ['displayName', 'email', 'jobTitle', 'department'] as const;
const PERMISSION_DATA_FIELDS = ['principal', 'role', 'inherited', 'targetName', 'targetUrl'] as const;
const GENERIC_DATA_FIELDS = ['label', 'message', 'name', 'fileName', 'url', 'target'] as const;
const LIST_ROW_FIELDS = ['displayName', 'name', 'title', 'description', 'id', 'webUrl', 'url', 'template'] as const;
const SELECTION_ITEM_FIELDS = ['id', 'label', 'description', 'url', 'target'] as const;
const SELECTED_ITEM_FIELDS = ['index', 'title', 'url'] as const;
const FORM_DATA_FIELDS = ['preset', 'toolName', 'success', 'message', 'fields', 'retryAction'] as const;

export class HieInteractionFormatter {
  private tracker: BlockTracker;

  constructor(tracker: BlockTracker) {
    this.tracker = tracker;
  }

  public format(interaction: IBlockInteraction): IHiePromptMessage | undefined {
    const p = interaction.payload;

    switch (interaction.action) {
      case 'click-result':
        return this.buildInteractionMessage(
          'The user opened a search result in the browser.',
          'The file is already open. Do not show file details or preview unless the user explicitly asks.',
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      case 'click-folder':
        return this.buildInteractionMessage(
          'The user clicked a folder in the document library.',
          'Acknowledge the navigation only if needed and use the folder metadata from Untrusted data for any follow-up.',
          this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
        );
      case 'click-file':
        return this.buildInteractionMessage(
          'The user clicked a file in the document library.',
          'Acknowledge the navigation only if needed and use the file metadata from Untrusted data for any follow-up.',
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      case 'click-user': {
        const intentHint = typeof p.userTextPreview === 'string' && p.userTextPreview.trim()
          ? ` The user's original request was: "${p.userTextPreview}".`
          : '';
        const instructions = intentHint
          ? 'The person data is already resolved — do NOT call search_people again or show a selection list. If the original request implies using this person (e.g., as attendee, recipient, email target), proceed directly with that action using the person\'s email from Untrusted data. Otherwise use the person data for any follow-up response.'
          : 'Use the person data from Untrusted data for any follow-up response.';
        return this.buildInteractionMessage(
          `The user clicked a person entry.${intentHint}`,
          instructions,
          this.pickAllowedFields(p, PERSON_DATA_FIELDS)
        );
      }
      case 'click-list-row':
        return this.buildInteractionMessage(
          'The user clicked a list row.',
          'Use the row metadata from Untrusted data for any follow-up response.',
          this.buildListRowInteractionData(p)
        );
      case 'click-permission':
        return this.buildInteractionMessage(
          'The user clicked a permission entry.',
          'Use the permission principal, role, and target from Untrusted data for any follow-up response.',
          this.pickAllowedFields(p, PERMISSION_DATA_FIELDS)
        );
      case 'click-activity':
        return this.buildInteractionMessage(
          'The user clicked an activity entry.',
          'Use the activity target from Untrusted data for any follow-up response.',
          this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
        );
      case 'confirm':
        return this.buildInteractionMessage(
          'The user confirmed a pending action.',
          'Proceed based on the confirmed message in Untrusted data.',
          this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
        );
      case 'cancel':
        return this.buildInteractionMessage(
          'The user cancelled a pending action.',
          'Respect the cancellation and do not continue the cancelled action.',
          this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
        );
      case 'select':
        return this.buildInteractionMessage(
          'The user selected an option from a selection list.',
          'Use the selected label from Untrusted data to continue the flow.',
          this.buildSelectionInteractionData(p)
        );
      case 'open-external':
        return this.buildInteractionMessage(
          'The user opened a file in the browser.',
          'The item is already open externally. Use the file metadata from Untrusted data for any follow-up response.',
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      case 'navigate':
        return this.buildInteractionMessage(
          'The user navigated within the action panel.',
          'Use the navigation target from Untrusted data for any follow-up response.',
          this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
        );
      case 'retry':
        return this.buildInteractionMessage(
          'The user clicked retry for a prior action.',
          'Use the retry target from Untrusted data to decide the next step.',
          this.pickAllowedFields(p, [...GENERIC_DATA_FIELDS, ...FORM_DATA_FIELDS])
        );
      case 'submit-form':
        return this.buildInteractionMessage(
          'The user submitted a compose form and the tool already executed.',
          'Do not call use_m365_capability or any other tool to repeat this action. Briefly acknowledge the result using the status and message from Untrusted data.',
          this.pickAllowedFields(p, FORM_DATA_FIELDS)
        );
      case 'cancel-form':
        return this.buildInteractionMessage(
          'The user cancelled a compose form.',
          'Acknowledge the cancellation only if useful and do not continue the cancelled form flow.',
          this.pickAllowedFields(p, FORM_DATA_FIELDS)
        );
      case 'look':
        return this.formatLookAction(interaction);
      case 'summarize':
        return this.formatSummarizeAction(interaction);
      case 'chat-about': {
        const focusedItems = this.resolveFocusedFileContexts(p);
        const selectedItems = focusedItems.length > 0
          ? focusedItems.map((item, index) => ({ index: index + 1, title: item.title, url: item.url }))
          : this.buildSelectedItemsData(p.selectedItems);
        const untrustedData: Record<string, unknown> = {
          ...this.pickAllowedFields(p, [...DOCUMENT_DATA_FIELDS, ...PERSON_DATA_FIELDS, ...GENERIC_DATA_FIELDS]),
          ...(selectedItems ? { selectedItems } : {})
        };

        if (selectedItems && selectedItems.length > 1) {
          return this.buildInteractionMessage(
            'The user wants to discuss multiple selected documents.',
            'Set the documents listed in selectedItems in Untrusted data as the conversation focus. Acknowledge briefly and ask what they want to know, such as compare, summarize differences, or answer specific questions. For follow-up questions involving multiple documents, call read_file_content with mode "answer" using the URLs from selectedItems in Untrusted data. If the user narrows to one document, use only that document URL. Use mode "full" only if the user explicitly asks for full content. Do not call read_file_content yet.',
            untrustedData
          );
        }

        if (selectedItems && selectedItems.length === 1) {
          return this.buildInteractionMessage(
            'The user wants to discuss a specific document.',
            'Set the document listed in selectedItems in Untrusted data as the conversation focus. Acknowledge briefly and ask what they want to know. For follow-up questions about that document, call read_file_content with mode "answer" using the document URL from selectedItems in Untrusted data and include the document title when available. Use mode "full" only if the user explicitly asks for full content. Do not call read_file_content yet.',
            untrustedData
          );
        }

        return this.buildInteractionMessage(
          'The user wants to discuss a selected item.',
          'Set the item described in Untrusted data as conversation focus. Acknowledge and ask what they want to know.',
          untrustedData
        );
      }
      case 'click-item':
        if (p.Subject || p.From) {
          return this.buildInteractionMessage(
            'The user clicked a specific email item.',
            'Retrieve the full content of this one email. Use read_email_content with mode "full" and use the subject and sender from Untrusted data when available. Do not summarize, do not search broadly, and do not call unrelated tools. Just retrieve and display this one email.',
            {
              ...this.pickAllowedFields(p, EMAIL_DATA_FIELDS),
              ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
            }
          );
        }
        if (typeof p.url === 'string' && p.url.trim()) {
          return this.buildInteractionMessage(
            'The user clicked a specific content item.',
            'Retrieve the full content of this one item. Use read_file_content with mode "full" and the file URL from Untrusted data. Include the title when available. Do not summarize, do not search broadly, and do not call unrelated tools. Just retrieve and display this one item.',
            {
              ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS),
              ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
            }
          );
        }
        return this.buildInteractionMessage(
          'The user clicked a specific content item.',
          'Retrieve the full content of this one item. Do not summarize, do not search broadly, and do not call unrelated tools. Just retrieve and display this one item.',
          {
            ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS),
            ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
          }
        );
      case 'dismiss':
        return undefined;
      default:
        return undefined;
    }
  }

  private pickAllowedFields(payload: Record<string, unknown>, allowedFields: readonly string[]): Record<string, unknown> {
    const picked: Record<string, unknown> = {};

    allowedFields.forEach((field) => {
      const value = payload[field];
      if (value === undefined || value === null || value === '') return;
      picked[field] = value;
    });

    return picked;
  }

  private buildSelectedItemsData(
    source: unknown
  ): Array<{ index?: number; title?: string; url?: string }> | undefined {
    if (!Array.isArray(source)) return undefined;

    const selectedItems = source
      .filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object')
      .map((entry) => {
        const picked = this.pickAllowedFields(entry, SELECTED_ITEM_FIELDS);
        return {
          index: typeof picked.index === 'number' ? picked.index : undefined,
          title: typeof picked.title === 'string' ? picked.title : undefined,
          url: typeof picked.url === 'string' ? picked.url : undefined
        };
      })
      .filter((entry) => entry.index !== undefined || entry.title || entry.url);

    return selectedItems.length > 0 ? selectedItems : undefined;
  }

  private buildListRowInteractionData(payload: Record<string, unknown>): Record<string, unknown> {
    const rowData = payload.rowData && typeof payload.rowData === 'object' && !Array.isArray(payload.rowData)
      ? payload.rowData as Record<string, unknown>
      : undefined;

    const picked = rowData
      ? this.pickAllowedFields(rowData, LIST_ROW_FIELDS)
      : this.pickAllowedFields(payload, [...GENERIC_DATA_FIELDS, ...LIST_ROW_FIELDS]);

    const webUrl = typeof picked.webUrl === 'string' ? picked.webUrl : undefined;
    if (webUrl && !picked.url) {
      picked.url = webUrl;
    }

    const index = typeof payload.index === 'number' ? payload.index : undefined;
    return {
      ...(index !== undefined ? { index } : {}),
      ...picked,
      ...(rowData ? { rowData } : {})
    };
  }

  private buildSelectionInteractionData(payload: Record<string, unknown>): Record<string, unknown> {
    const picked = this.pickAllowedFields(payload, [...GENERIC_DATA_FIELDS, 'prompt', 'selectedIds']);
    const selectedItems = Array.isArray(payload.selectedItems)
      ? payload.selectedItems
        .filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object' && !Array.isArray(entry))
        .map((entry) => {
          const item = this.pickAllowedFields(entry, SELECTION_ITEM_FIELDS);
          const description = typeof item.description === 'string' ? item.description : undefined;
          if (!item.url && description && /^https?:\/\//i.test(description)) {
            item.url = description;
          }
          return item;
        })
        .filter((entry) => Object.keys(entry).length > 0)
      : undefined;

    return {
      ...picked,
      ...(selectedItems && selectedItems.length > 0 ? { selectedItems } : {})
    };
  }

  private buildInteractionMessage(
    trustedAction: string,
    trustedInstructions: string,
    untrustedData: Record<string, unknown>
  ): IHiePromptMessage {
    return {
      kind: 'interaction',
      body: [
      `Trusted action: ${trustedAction}`,
      `Trusted instructions: ${trustedInstructions}`,
      `Untrusted data: ${serializeUntrustedData(untrustedData)}`,
      ].join('\n')
    };
  }

  private formatLookAction(interaction: IBlockInteraction): IHiePromptMessage {
    const p = interaction.payload;

    switch (interaction.blockType) {
      case 'search-results': {
        const focused = this.resolveFocusedFileContext(p);
        const untrustedData: Record<string, unknown> = {
          ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        };
        if (focused?.url) {
          untrustedData.selectedItems = [{ index: 1, title: focused.title, url: focused.url }];
        }

        return this.buildInteractionMessage(
          'The user clicked Focus on a search result.',
          focused?.url
            ? 'Describe the document using the metadata in Untrusted data in one complete response. Do not ask follow-up questions. Do not call any tools right now. Set this document as focused context. For any follow-up question about the same document, call read_file_content with mode "answer" using the document URL from selectedItems in Untrusted data and include the document title when available. Use mode "full" only if the user explicitly asks for full content.'
            : 'Describe the document using the metadata in Untrusted data in one complete response. Do not ask follow-up questions. Do not call any tools.',
          untrustedData
        );
      }
      case 'document-library':
        return this.buildInteractionMessage(
          'The user clicked Focus on a document library item.',
          'Describe the file or folder using the metadata in Untrusted data in one complete response. Do not ask follow-up questions. Do not call any tools.',
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      case 'markdown':
        if (p.itemId || p.Subject || p.From) {
          return this.buildInteractionMessage(
            'The user clicked Focus on an email card.',
            'Retrieve and display the full email content. Call read_email_content with mode "full" using the subject and sender from Untrusted data when available. Do not describe the card in chat; fetch the real content.',
            this.pickAllowedFields(p, EMAIL_DATA_FIELDS)
          );
        }
        return this.buildInteractionMessage(
          'The user clicked Focus on an item card.',
          'Describe the item using the fields in Untrusted data in one complete response. Do not offer further actions.',
          {
            ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS),
            ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
          }
        );
      case 'list-items':
        return this.buildInteractionMessage(
          'The user clicked Focus on a list row.',
          'Describe the row using the fields in Untrusted data in one complete response. Do not ask follow-up questions. Do not call any tools.',
          this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
        );
      case 'user-card':
        return this.buildInteractionMessage(
          'The user clicked Focus on a person card.',
          'Describe the person using the profile data in Untrusted data in one complete response. Do not ask follow-up questions. Do not call any tools.',
          this.pickAllowedFields(p, PERSON_DATA_FIELDS)
        );
      default:
        return this.buildInteractionMessage(
          'The user clicked Focus on an item.',
          'Describe the item using the data in Untrusted data in one complete response. Do not ask follow-up questions. Do not call any tools.',
          {
            ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS),
            ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS),
            ...this.pickAllowedFields(p, PERSON_DATA_FIELDS)
          }
        );
    }
  }

  private formatSummarizeAction(interaction: IBlockInteraction): IHiePromptMessage {
    const p = interaction.payload;

    switch (interaction.blockType) {
      case 'search-results':
        return this.buildInteractionMessage(
          'The user clicked Summarize on a search result.',
          `Process the document referenced in Untrusted data with Copilot. Use read_file_content with mode "summarize" and the document URL from Untrusted data. Include the document title when available. The tool result already contains the final summary text. Then call show_info_card to place that summary in the action panel with heading "Summary: <document title from Untrusted data>". Target summary body length: ${ASSISTANT_SUMMARY_TARGET_TEXT}. Do not re-summarize or rewrite from scratch. After that, send only a brief one-sentence acknowledgment in chat. Do not paste the full summary in chat.`,
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      case 'document-library':
        return this.buildInteractionMessage(
          'The user clicked Summarize on a document library item.',
          `Process the file referenced in Untrusted data with Copilot. Use read_file_content with mode "summarize" and the document URL from Untrusted data. Include the file name when available. The tool result already contains the final summary text. Then call show_info_card to place that summary in the action panel with heading "Summary: <document name from Untrusted data>". Target summary body length: ${ASSISTANT_SUMMARY_TARGET_TEXT}. Do not re-summarize or rewrite from scratch. After that, send only a brief one-sentence acknowledgment in chat. Do not paste the full summary in chat.`,
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      case 'markdown':
        if (p.itemId || p.Subject || p.From) {
          return this.buildInteractionMessage(
            'The user clicked Summarize on an email.',
            `Process the email referenced in Untrusted data with Copilot. Call read_email_content with mode "summarize" using the subject and sender from Untrusted data when available. The tool result already contains the final summary text. Then call show_info_card to place that summary in the action panel with heading "Summary: <email subject from Untrusted data>". Target summary body length: ${ASSISTANT_SUMMARY_TARGET_TEXT}. Do not re-summarize or rewrite from scratch. After that, send only a brief one-sentence acknowledgment in chat. Do not paste the full summary in chat.`,
            this.pickAllowedFields(p, EMAIL_DATA_FIELDS)
          );
        }
        return this.buildInteractionMessage(
          'The user clicked Summarize on an item.',
          `Retrieve and summarize the content referenced in Untrusted data with Copilot using mode "summarize". Then call show_info_card to place the summary in the action panel. Target summary body length: ${ASSISTANT_SUMMARY_TARGET_TEXT}. After that, send only a brief one-sentence acknowledgment in chat. Do not paste the full summary in chat.`,
          {
            ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS),
            ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
          }
        );
      case 'file-preview':
        return this.buildInteractionMessage(
          'The user clicked Summarize on a file preview.',
          `Process the file referenced in Untrusted data with Copilot. Use read_file_content with mode "summarize" and the document URL from Untrusted data. Include the file name when available. The tool result already contains the final summary text. Then call show_info_card to place that summary in the action panel with heading "Summary: <file name from Untrusted data>". Target summary body length: ${ASSISTANT_SUMMARY_TARGET_TEXT}. Do not re-summarize or rewrite from scratch. After that, send only a brief one-sentence acknowledgment in chat. Do not paste the full summary in chat.`,
          this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS)
        );
      default:
        return this.buildInteractionMessage(
          'The user clicked Summarize on an item.',
          `Retrieve and summarize the content referenced in Untrusted data. Show the summary in an action-panel info card via show_info_card with a target body length of ${ASSISTANT_SUMMARY_TARGET_TEXT}, then acknowledge briefly in chat.`,
          {
            ...this.pickAllowedFields(p, DOCUMENT_DATA_FIELDS),
            ...this.pickAllowedFields(p, GENERIC_DATA_FIELDS)
          }
        );
    }
  }

  private normalizeFocusedLabel(value: unknown): string | undefined {
    if (typeof value !== 'string') return undefined;
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    return trimmed
      .replace(/^"(.*)"$/, '$1')
      .replace(/^(summary|content|answer)\s*:\s*/i, '')
      .trim() || undefined;
  }

  private resolveFocusedFileContext(payload: Record<string, unknown>): { title?: string; url: string } | undefined {
    const directUrl = typeof payload.url === 'string' ? payload.url.trim() : '';
    const candidates = [
      this.normalizeFocusedLabel(payload.title),
      this.normalizeFocusedLabel(payload.name),
      this.normalizeFocusedLabel(payload.fileName),
      this.normalizeFocusedLabel(payload.heading)
    ].filter((v): v is string => !!v);

    if (directUrl) {
      return { title: candidates[0], url: directUrl };
    }

    if (candidates.length === 0) return undefined;
    const candidateSet = new Set(candidates.map((c) => c.toLowerCase()));
    const allRefs = this.tracker.getAllReferences();

    for (let i = 0; i < allRefs.length; i++) {
      const refs = allRefs[i].references;
      for (let j = 0; j < refs.length; j++) {
        const ref = refs[j];
        const refUrl = typeof ref.url === 'string' ? ref.url.trim() : '';
        const refTitle = this.normalizeFocusedLabel(ref.title);
        const refTitleLower = refTitle?.toLowerCase();
        if (!refUrl || !refTitleLower) continue;
        if (candidateSet.has(refTitleLower)) {
          return { title: ref.title, url: refUrl };
        }
      }
    }

    return undefined;
  }

  private resolveFocusedFileContexts(payload: Record<string, unknown>): Array<{ title?: string; url: string }> {
    const rawSelection = payload.selectedItems;
    const resolved: Array<{ title?: string; url: string }> = [];
    const dedupe = new Set<string>();

    if (Array.isArray(rawSelection)) {
      for (let i = 0; i < rawSelection.length; i++) {
        const entry = rawSelection[i];
        if (!entry || typeof entry !== 'object') continue;
        const candidate = entry as Record<string, unknown>;
        const url = typeof candidate.url === 'string' ? candidate.url.trim() : '';
        if (!url) continue;
        const dedupeKey = url.toLowerCase();
        if (dedupe.has(dedupeKey)) continue;
        dedupe.add(dedupeKey);
        const title = this.normalizeFocusedLabel(candidate.title)
          || this.normalizeFocusedLabel(candidate.name)
          || this.normalizeFocusedLabel(candidate.fileName);
        resolved.push({ title, url });
      }
    }

    if (resolved.length > 0) return resolved;

    const single = this.resolveFocusedFileContext(payload);
    return single ? [single] : [];
  }
}
