/**
 * McpResultMapper
 * Maps MCP tool results to typed UI blocks using schema inference.
 *
 * Instead of hardcoded tool-name → block-type lookup tables,
 * this uses SchemaInferrer to detect the best block type from data shape,
 * with optional catalog blockHint overrides for ambiguous cases.
 */

import type { IBlock, BlockType } from '../../models/IBlock';
import type { IMcpContent } from '../../models/IMcpTypes';
import { M365_MCP_CATALOG } from '../../models/McpServerCatalog';
import type { IMcpCatalogTool } from '../../models/McpServerCatalog';
import { infer } from './SchemaInferrer';
import { build } from './GenericBlockBuilder';
import { extractMcpText } from './mcpUtils';
import { prepareAgentReplyForDisplay } from './EntityParsers';
import { logService } from '../logging/LogService';

// ─── Constants ──────────────────────────────────────────────────

/** Friendly field labels for email single-resource rendering — ordered for readability. */
const EMAIL_FRIENDLY_FIELDS: ReadonlyArray<[string, string]> = [
  ['subject', 'Subject'],
  ['from', 'From'],
  ['toRecipients', 'To'],
  ['ccRecipients', 'Cc'],
  ['receivedDateTime', 'Date'],
  ['importance', 'Importance'],
  ['hasAttachments', 'Attachments'],
];

/** Fields that identify a { data } response as email-shaped (vs calendar, Teams, etc.) */
const EMAIL_MARKER_FIELDS = ['from', 'toRecipients', 'subject'];
const CALENDAR_EVENT_LIST_TOOLS: ReadonlySet<string> = new Set([
  'listcalendarview',
  'listevents'
]);

interface ICalendarEventRow {
  event: string;
  date?: string;
  start?: string;
  end?: string;
  location?: string;
  organizer?: string;
  openUrl?: string;
}

interface ICopilotReplyCandidate {
  text: string;
  source: string;
}

// ─── Helpers ────────────────────────────────────────────────────

/**
 * Truncate content and wrap with the LLM instruction prefix.
 */
function formatLlmReturn(content: string, maxLen: number = 2000): string {
  const truncated = content.length > maxLen ? content.substring(0, maxLen) + '...' : content;
  return `[Results already displayed to user in a visual block. DO NOT call show_markdown to re-display this data. Summarize or answer based on the following actual data:]\n${truncated}`;
}

/**
 * Strip HTML tags and decode common entities, returning plain text.
 */
function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#\d+;/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
  return value as Record<string, unknown>;
}

function asObjectArray(value: unknown): Record<string, unknown>[] {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is Record<string, unknown> =>
    !!item && typeof item === 'object' && !Array.isArray(item)
  );
}

function firstString(obj: Record<string, unknown>, keys: string[]): string | undefined {
  for (let i = 0; i < keys.length; i++) {
    const raw = obj[keys[i]];
    if (typeof raw === 'string') {
      const trimmed = raw.trim();
      if (trimmed.length > 0) return trimmed;
    }
  }
  return undefined;
}

function extractDateTimeString(value: unknown): string | undefined {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  const record = asRecord(value);
  if (!record) return undefined;

  return firstString(record, ['dateTime', 'startDateTime', 'endDateTime', 'date'])
    || extractDateTimeString(record.start)
    || extractDateTimeString(record.end);
}

function formatCalendarDateTime(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) {
    return raw.replace('T', ' ').replace(/\.\d+Z?$/, '');
  }

  const pad = (n: number): string => n.toString().padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function extractLocationLabel(value: unknown): string | undefined {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  const record = asRecord(value);
  if (!record) return undefined;

  const direct = firstString(record, ['displayName', 'name', 'address']);
  if (direct) return direct;

  const address = asRecord(record.address);
  if (address) {
    const parts = [
      firstString(address, ['street']),
      firstString(address, ['city']),
      firstString(address, ['state']),
      firstString(address, ['countryOrRegion'])
    ].filter((part): part is string => !!part);
    if (parts.length > 0) return parts.join(', ');
  }

  return undefined;
}

function extractOrganizerLabel(value: unknown): string | undefined {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  const record = asRecord(value);
  if (!record) return undefined;

  const direct = firstString(record, ['displayName', 'name', 'address', 'email']);
  if (direct) return direct;

  const emailAddress = asRecord(record.emailAddress);
  if (emailAddress) {
    const name = firstString(emailAddress, ['name']);
    const address = firstString(emailAddress, ['address']);
    if (name && address) return `${name} <${address}>`;
    if (address) return address;
    if (name) return name;
  }

  return undefined;
}

function getCalendarItems(parsed: unknown): Record<string, unknown>[] | undefined {
  if (Array.isArray(parsed)) {
    return asObjectArray(parsed);
  }

  const root = asRecord(parsed);
  if (!root) return undefined;

  const candidates = [root.value, root.events, root.items, root.results];
  for (let i = 0; i < candidates.length; i++) {
    if (Array.isArray(candidates[i])) {
      return asObjectArray(candidates[i]);
    }
  }

  return undefined;
}

function extractCalendarRows(parsed: unknown): ICalendarEventRow[] | undefined {
  const items = getCalendarItems(parsed);
  if (!items) return undefined;

  return items
    .map((item) => {
      const slot = asRecord(item.meetingTimeSlot);
      const event = firstString(item, ['subject', 'title', 'meetingTitle']) || 'Calendar event';
      const startRaw = extractDateTimeString(item.start) || extractDateTimeString(item.startDateTime) || extractDateTimeString(slot?.start);
      const endRaw = extractDateTimeString(item.end) || extractDateTimeString(item.endDateTime) || extractDateTimeString(slot?.end);
      const start = formatCalendarDateTime(startRaw);
      const end = formatCalendarDateTime(endRaw);
      const date = start || end;
      const location = extractLocationLabel(item.location) || extractLocationLabel(slot?.location);
      const organizer = extractOrganizerLabel(item.organizer);
      const openUrl = firstString(item, ['webLink', 'joinWebUrl', 'url']);

      return {
        event,
        date,
        start,
        end,
        location,
        organizer,
        openUrl
      } as ICalendarEventRow;
    })
    .filter((row) => !!row.event);
}

function isCalendarNoResultsText(text: string): boolean {
  const normalized = text.toLowerCase();
  return normalized.indexOf('no calendar events found') !== -1
    || normalized.indexOf('no events found for the given criteria') !== -1;
}

function mapCalendarEventResult(
  serverId: string,
  toolName: string,
  parsed: unknown,
  rawText: string
): { block: IBlock; llmText: string } | undefined {
  if (serverId !== 'mcp_CalendarTools') return undefined;
  if (!CALENDAR_EVENT_LIST_TOOLS.has(toolName.toLowerCase())) return undefined;

  const rows = extractCalendarRows(parsed);
  if (!rows && !isCalendarNoResultsText(rawText)) return undefined;

  if (!rows || rows.length === 0) {
    const body = 'No calendar events found for the selected time range.';
    const block = build(
      { blockType: 'info-card', confidence: 'high', fieldMap: { heading: 'heading', body: 'body' } },
      { heading: 'Calendar', body },
      body,
      toolName
    );
    block.title = 'Calendar';
    return { block, llmText: body };
  }

  const lines: string[] = [];
  rows.forEach((row, index) => {
    lines.push(`${index + 1}. **Event:** ${row.event}`);
    if (row.date) lines.push(`   **Date:** ${row.date}`);
    if (row.start) lines.push(`   **Start:** ${row.start}`);
    if (row.end) lines.push(`   **End:** ${row.end}`);
    if (row.location) lines.push(`   **Location:** ${row.location}`);
    if (row.organizer) lines.push(`   **Organizer:** ${row.organizer}`);
    if (row.openUrl) lines.push(`   **Open:** [Open in Outlook](${row.openUrl})`);
  });

  const content = lines.join('\n');
  const block = build(
    { blockType: 'markdown', confidence: 'high', fieldMap: {} },
    undefined,
    content,
    toolName
  );
  block.title = 'Calendar Events';
  return { block, llmText: content };
}

/**
 * Check if a data object looks like an email (has email-specific marker fields).
 */
function isEmailShaped(data: Record<string, unknown>): boolean {
  return EMAIL_MARKER_FIELDS.some((f) => data[f] !== undefined);
}

/**
 * Try to parse text content as JSON, returning undefined on failure.
 */
function tryParseJson(text: string): unknown | undefined {
  try {
    return JSON.parse(text);
  } catch {
    const trimmed = text.trim();
    if (!trimmed) return undefined;

    const starts: number[] = [];
    const firstObject = trimmed.indexOf('{');
    const firstArray = trimmed.indexOf('[');
    if (firstObject >= 0) starts.push(firstObject);
    if (firstArray >= 0) starts.push(firstArray);
    starts.sort((a, b) => a - b);

    for (let i = 0; i < starts.length; i++) {
      const candidate = extractBalancedJson(trimmed, starts[i]);
      if (!candidate) continue;
      try {
        return JSON.parse(candidate);
      } catch {
        // Continue trying other candidates.
      }
    }
    return undefined;
  }
}

function unwrapEmbeddedPayload(parsed: unknown): unknown {
  let current = parsed;

  for (let depth = 0; depth < 3; depth++) {
    const record = asRecord(current);
    if (!record) return current;

    const candidates: unknown[] = [record.response, record.result, record.payload];
    let unwrapped: unknown | undefined;

    for (let i = 0; i < candidates.length; i++) {
      const candidate = candidates[i];
      if (typeof candidate === 'string') {
        const parsedCandidate = tryParseJson(candidate);
        if (parsedCandidate !== undefined) {
          unwrapped = parsedCandidate;
          break;
        }
      } else if (candidate && typeof candidate === 'object') {
        unwrapped = candidate;
        break;
      }
    }

    if (unwrapped === undefined) return current;
    current = unwrapped;
  }

  return current;
}

function extractBalancedJson(text: string, start: number): string | undefined {
  const first = text[start];
  if (first !== '{' && first !== '[') return undefined;

  const stack: string[] = [first === '{' ? '}' : ']'];
  let inString = false;
  let escaping = false;

  for (let i = start + 1; i < text.length; i++) {
    const ch = text[i];

    if (inString) {
      if (escaping) {
        escaping = false;
      } else if (ch === '\\') {
        escaping = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') {
      stack.push('}');
      continue;
    }
    if (ch === '[') {
      stack.push(']');
      continue;
    }
    if (ch === '}' || ch === ']') {
      const expected = stack.pop();
      if (expected !== ch) return undefined;
      if (stack.length === 0) {
        return text.slice(start, i + 1);
      }
    }
  }

  return undefined;
}

function stripCopilotMetadata(text: string): string {
  return text
    .replace(/\s*;\s*CorrelationId:\s*[^\n]+/gi, '')
    .replace(/\s*;\s*TimeStamp:\s*[^\n]+/gi, '')
    .replace(/(?:^|\n)\s*(?:CorrelationId|TimeStamp)\s*:\s*[^\n]*/gi, '')
    .trim();
}

function looksLikeDisplayText(text: string): boolean {
  const trimmed = stripCopilotMetadata(text);
  if (!trimmed || trimmed.length < 8) return false;
  if (!/[A-Za-z]/.test(trimmed)) return false;
  if (/^(?:\{|\[)/.test(trimmed)) return false;
  if (/^[A-Za-z0-9_-]+$/.test(trimmed)) return false;
  if (/^https?:\/\/\S+$/i.test(trimmed)) return false;
  if (/^[0-9a-f]{8}-[0-9a-f-]{20,}$/i.test(trimmed)) return false;
  if (/^(?:ok|success|completed)$/i.test(trimmed)) return false;
  return true;
}

function extractCopilotText(value: unknown, depth: number = 0): string | undefined {
  if (depth > 8 || value === undefined || value === null) return undefined;

  if (typeof value === 'string') {
    const nested = tryParseJson(value);
    if (nested !== undefined && nested !== value) {
      const parsedCandidate = extractCopilotText(nested, depth + 1);
      if (parsedCandidate) return parsedCandidate;
    }

    const cleaned = stripCopilotMetadata(value);
    return looksLikeDisplayText(cleaned) ? cleaned : undefined;
  }

  if (Array.isArray(value)) {
    for (let i = value.length - 1; i >= 0; i--) {
      const candidate = extractCopilotText(value[i], depth + 1);
      if (candidate) return candidate;
    }
    return undefined;
  }

  const record = asRecord(value);
  if (!record) return undefined;

  const role = firstString(record, ['role', 'author', 'sender', 'from']);
  if (role && /user/i.test(role)) {
    return undefined;
  }

  const priorityKeys = ['reply', 'text', 'content', 'body', 'summary', 'answer', 'response', 'message'];
  for (let i = 0; i < priorityKeys.length; i++) {
    const candidate = extractCopilotText(record[priorityKeys[i]], depth + 1);
    if (candidate) return candidate;
  }

  const containerKeys = ['messages', 'choices', 'responses', 'output', 'data', 'result', 'value', 'items'];
  for (let i = 0; i < containerKeys.length; i++) {
    const candidate = extractCopilotText(record[containerKeys[i]], depth + 1);
    if (candidate) return candidate;
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const candidate = extractCopilotText(record[keys[i]], depth + 1);
    if (candidate) return candidate;
  }

  return undefined;
}

function extractCopilotReplyCandidate(payload: Record<string, unknown>): ICopilotReplyCandidate | undefined {
  const candidates: Array<[string, unknown]> = [
    ['reply', payload.reply],
    ['rawResponse', payload.rawResponse],
    ['response', payload.response],
    ['message', payload.message],
    ['text', payload.text],
    ['content', payload.content]
  ];

  for (let i = 0; i < candidates.length; i++) {
    const [source, value] = candidates[i];
    const text = extractCopilotText(value);
    if (text) {
      return { text, source };
    }
  }

  return undefined;
}

// ─── Catalog Lookup ─────────────────────────────────────────────

/**
 * Find a catalog tool entry by server ID and tool name.
 * Falls back to searching all servers if serverId doesn't match.
 */
function findCatalogTool(serverId: string, toolName: string): IMcpCatalogTool | undefined {
  const lower = toolName.toLowerCase();
  // Try specific server first
  const server = M365_MCP_CATALOG.find((s) => s.id === serverId);
  if (server) {
    const tool = server.tools.find((t) => t.name.toLowerCase() === lower);
    if (tool) return tool;
  }
  // Fall back to searching all servers
  let found: IMcpCatalogTool | undefined;
  M365_MCP_CATALOG.forEach((s) => {
    if (!found) {
      const t = s.tools.find((tool) => tool.name.toLowerCase() === lower);
      if (t) found = t;
    }
  });
  return found;
}

// ─── Public API ─────────────────────────────────────────────────

/**
 * Map an MCP tool result to a typed UI block and push it to the action panel.
 *
 * Uses schema inference to detect the best block type from data shape,
 * with optional catalog blockHint overrides.
 *
 * @param serverId - The MCP server ID (used to look up catalog hints)
 * @param toolName - The MCP tool name that produced the result
 * @param content - The MCP response content array
 * @param pushBlock - Store action to push a block to the action panel
 * @returns A short summary string for the LLM response
 */
export function mapMcpResultToBlock(
  serverId: string,
  toolName: string,
  content: IMcpContent[],
  pushBlock: (block: IBlock) => void,
  onBlockCreated?: (block: IBlock) => void
): string {
  const text = extractMcpText(content);
  if (!text) {
    return 'No content returned from MCP tool.';
  }

  // Try parsing each content item individually first (Agent 365 often
  // returns JSON in the first item and metadata like CorrelationId in subsequent items).
  // Falling back to parsing the full joined text handles single-item responses.
  let parsed: unknown | undefined;
  for (let i = 0; i < content.length; i++) {
    if (content[i].type === 'text' && content[i].text) {
      const candidate = tryParseJson(content[i].text!);
      if (candidate !== undefined) {
        parsed = candidate;
        break;
      }
    }
  }
  if (parsed === undefined) {
    parsed = tryParseJson(text);
  }

  // 0. Check catalog for block hint (used in both reply and general paths)
  const catalogTool = findCatalogTool(serverId, toolName);
  const hint: BlockType | undefined = catalogTool?.blockHint;

  // Agent 365 responses often have a { reply: "markdown...", message: "status" } shape.
  // Render as cleaned markdown — HIE extracts numbered references for voice interaction.
  if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
    const obj = parsed as Record<string, unknown>;
    if (typeof obj.reply === 'string') {
      const replyCandidate = extractCopilotReplyCandidate(obj);
      if (!replyCandidate) {
        logService.warning(
          'mcp',
          `Empty Copilot reply for ${toolName}; no user-facing text recovered from payload keys: ${Object.keys(obj).join(', ')}`
        );
        return 'Copilot returned an empty reply and no user-facing text could be recovered from the payload. Avoid claiming that the target URL has no public results; the tool may simply not support this lookup.';
      }
      const reply = replyCandidate.text;
      logService.debug('mcp', `Recovered reply from ${replyCandidate.source} (${reply.length} chars):\n${reply.substring(0, 200)}`);
      const { cleanedReply, itemIds } = prepareAgentReplyForDisplay(reply);
      const hasItemIds = Object.keys(itemIds).length > 0;
      const block = build(
        { blockType: 'markdown', confidence: 'high', fieldMap: {} },
        undefined,
        cleanedReply,
        toolName
      );
      if (hasItemIds) {
        (block.data as { itemIds?: Record<number, string> }).itemIds = itemIds;
      }
      pushBlock(block);
      if (onBlockCreated) onBlockCreated(block);
      return formatLlmReturn(cleanedReply);
    }
  }

  // Agent 365 single email response: { message: "...", data: { subject, from, body, ... } }
  // Only matches email-shaped data (has from/toRecipients/subject) to avoid swallowing
  // calendar events, Teams messages, etc. that also use { message, data }.
  if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
    const obj = parsed as Record<string, unknown>;
    if (typeof obj.data === 'object' && obj.data !== null && !Array.isArray(obj.data)) {
      const data = obj.data as Record<string, unknown>;
      if (isEmailShaped(data)) {
        const lines: string[] = [];

        EMAIL_FRIENDLY_FIELDS.forEach(([key, label]) => {
          const val = data[key];
          if (val === undefined || val === null || val === '') return;
          if (key === 'hasAttachments' && val === false) return;
          if (key === 'hasAttachments' && val === true) { lines.push(`**${label}:** Yes`); return; }
          if (key === 'importance' && val === 'Normal') return;
          if (Array.isArray(val)) {
            if (val.length > 0) lines.push(`**${label}:** ${val.join(', ')}`);
          } else if (key === 'receivedDateTime' && typeof val === 'string') {
            try { lines.push(`**${label}:** ${new Date(val).toLocaleString()}`); } catch { lines.push(`**${label}:** ${val}`); }
          } else {
            lines.push(`**${label}:** ${String(val)}`);
          }
        });

        // Body: strip HTML, show plain text
        const body = data.body;
        if (typeof body === 'string' && body.length > 0) {
          const plain = stripHtml(body);
          if (plain.length > 0) {
            lines.push('---');
            lines.push(plain.substring(0, 2000));
          }
        } else if (typeof data.bodyPreview === 'string' && data.bodyPreview.length > 0) {
          lines.push('---');
          lines.push(data.bodyPreview as string);
        }
        if (lines.length > 0) {
          const cleaned = lines.join('\n\n');
          const block = build(
            { blockType: hint || 'markdown', confidence: 'high', fieldMap: {} },
            undefined,
            cleaned,
            toolName
          );
          pushBlock(block);
          if (onBlockCreated) onBlockCreated(block);
          return formatLlmReturn(cleaned);
        }
      }
    }
  }

  const dataPayload = unwrapEmbeddedPayload(parsed);

  const calendarMapped = mapCalendarEventResult(serverId, toolName, dataPayload, text);
  if (calendarMapped) {
    pushBlock(calendarMapped.block);
    if (onBlockCreated) onBlockCreated(calendarMapped.block);
    return formatLlmReturn(calendarMapped.llmText);
  }

  // 1. Use pre-fetched catalog hint

  // 2. Schema inference (uses hint as override if present)
  const inference = infer(dataPayload !== undefined ? dataPayload : text, hint);

  // 3. Build block from inference
  const block: IBlock = build(inference, dataPayload, text, toolName);
  pushBlock(block);
  if (onBlockCreated) onBlockCreated(block);

  return formatLlmReturn(text);
}
