/**
 * EntityParsers — Utilities for cleaning Agent 365 AI markdown replies.
 *
 * Agent 365 MCP tools often return `{ reply: "AI markdown..." }` with emoji,
 * analysis sections, and broken URLs that should be stripped before display.
 */

import { logService } from '../logging/LogService';

interface IPreparedAgentReply {
  cleanedReply: string;
  itemIds: Record<number, string>;
}

interface INumberedSection {
  number: number;
  raw: string;
}

const OUTLOOK_ITEM_LINK_PATTERN = /\[(\d+)\]\(https?:\/\/outlook\.office(?:365)?\.com\/[^)]*ItemID=([^&)]+)/g;
const FIRST_OUTLOOK_ITEM_LINK_PATTERN = /https?:\/\/outlook\.office(?:365)?\.com\/[^)\s]*ItemID=([^&)]+)/;
const NUMBERED_ITEM_START_PATTERN = /^\s*(\d+)[.)]\s*(.*)$/;

// ─── cleanAgentReply ─────────────────────────────────────────────

function normalizeNumberedHeadings(text: string): string {
  return text
    .replace(/^#{1,3}\s*\*{0,2}(\d+)[.)]\*{0,2}\s*/gm, '$1. ')
    .replace(/^\s*\*{1,2}(\d+)[.)]\*{1,2}\s*/gm, '$1. ');
}

function splitIntoNumberedSections(content: string): { preamble: string; sections: INumberedSection[] } {
  const normalized = normalizeNumberedHeadings(content).replace(/\r\n/g, '\n');
  const lines = normalized.split('\n');
  const preamble: string[] = [];
  const sections: INumberedSection[] = [];
  let currentNumber = 0;
  let currentLines: string[] | undefined;

  const pushCurrent = (): void => {
    if (!currentLines || currentLines.length === 0) return;
    sections.push({
      number: currentNumber,
      raw: currentLines.join('\n').trim()
    });
    currentLines = undefined;
  };

  lines.forEach((line) => {
    const match = NUMBERED_ITEM_START_PATTERN.exec(line.trim());
    if (match) {
      pushCurrent();
      currentNumber = parseInt(match[1], 10);
      currentLines = [`${currentNumber}. ${match[2]}`.trimEnd()];
      return;
    }

    if (currentLines) {
      currentLines.push(line);
      return;
    }

    preamble.push(line);
  });

  pushCurrent();
  return {
    preamble: preamble.join('\n').trim(),
    sections
  };
}

function sanitizeDedupedPreamble(preamble: string): string {
  if (!preamble.trim()) return '';

  const filtered = preamble
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => !!line)
    .filter((line) => !/\b(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+)\b.*\bemails?\b/i.test(line))
    .filter((line) => !/\bonly\b.*\bfound\b/i.test(line));

  return filtered.join('\n\n');
}

function renumberSection(section: string, nextNumber: number): string {
  return normalizeNumberedHeadings(section).replace(/^\s*\d+[.)]\s*/, `${nextNumber}. `);
}

function dedupeOutlookReply(rawReply: string): { reply: string; itemIds: Record<number, string> } {
  const itemIds = extractOutlookItemIds(rawReply);
  if (Object.keys(itemIds).length < 2) {
    return { reply: rawReply, itemIds };
  }

  const { preamble, sections } = splitIntoNumberedSections(rawReply);
  if (sections.length < 2) {
    return { reply: rawReply, itemIds };
  }

  const seenItemIds = new Set<string>();
  const keptSections: string[] = [];
  const dedupedItemIds: Record<number, string> = {};
  let removedAny = false;

  sections.forEach((section) => {
    const itemId = itemIds[section.number] || FIRST_OUTLOOK_ITEM_LINK_PATTERN.exec(section.raw)?.[1];
    if (itemId && seenItemIds.has(itemId)) {
      removedAny = true;
      return;
    }

    const nextNumber = keptSections.length + 1;
    keptSections.push(renumberSection(section.raw, nextNumber));
    if (itemId) {
      seenItemIds.add(itemId);
      dedupedItemIds[nextNumber] = itemId;
    }
  });

  if (!removedAny) {
    return { reply: rawReply, itemIds };
  }

  const sanitizedPreamble = sanitizeDedupedPreamble(preamble);
  const rebuiltReply = [
    sanitizedPreamble,
    keptSections.join('\n\n---\n\n')
  ].filter((part) => !!part).join('\n\n');

  return {
    reply: rebuiltReply,
    itemIds: dedupedItemIds
  };
}

/** Emoji ranges using surrogate pairs (ES5-compatible, no 'u' flag).
 * Covers: Miscellaneous Symbols (U+2600-26FF), Dingbats (U+2700-27BF),
 * and supplementary emoji via surrogate pairs (U+1F300-1F9FF). */
// eslint-disable-next-line no-misleading-character-class
const EMOJI_PATTERN = /[\u2600-\u26FF\u2700-\u27BF\uFE00-\uFE0F\u200D\u20E3]|\uD83C[\uDF00-\uDFFF]|\uD83D[\uDC00-\uDEFF]|\uD83E[\uDD00-\uDDFF]/g;

/**
 * Clean Agent 365 AI markdown replies:
 * - Strip "What You Can Do Next" / "Summary" / "Actions" tail sections
 * - Remove emoji characters
 * - Strip Outlook and Teams web URLs
 * - Collapse excess whitespace
 */
export function cleanAgentReply(text: string): string {
  return normalizeNumberedHeadings(text)
    // Strip tail sections: common AI analysis/suggestion headings and everything after (en/fr/it/de/es)
    .replace(/\n+\s*(?:#{1,3}\s*)?(?:\*\*)?(?:What You Can Do|Summary|Actions? Available|Next Steps|Suggested Actions|Available Actions|High[‑-]Level Patterns|Patterns Noticed|Key Observations|Additional Notes|Notes?|Ce que vous pouvez faire|(?:Résum[eé]|R[eé]sumé)|Actions? (?:disponibles|sugg[eé]r[eé]es)|Prochaines [eé]tapes|Cosa puoi fare|Riepilogo|Azioni (?:disponibili|suggerite)|Prossimi passi|Was Sie tun k[oö]nnen|Zusammenfassung|(?:Verf[uü]gbare |Vorgeschlagene )?Aktionen|N[aä]chste Schritte|Lo que puedes hacer|Resumen|Acciones (?:disponibles|sugeridas)|Pr[oó]ximos pasos)(?:\*\*)?[^\n]*(?:\n[\s\S]*)?$/i, '')
    // Strip "If you'd like" / "Would you like" / "Let me know" tail paragraphs (en/fr/it/de/es)
    .replace(/\n+\s*(?:If you(?:'d| would)? (?:like|want|need|prefer)|Would you like|Let me know|Feel free to|Si vous (?:souhaitez|voulez)|Souhaitez-vous|Dites-moi|N'h[eé]sitez pas|Se (?:desideri|vuoi)|Fammi sapere|Non esitare|Wenn Sie m[oö]chten|M[oö]chten Sie|Lassen Sie mich wissen|Z[oö]gern Sie nicht|Si (?:desea|quiere)|Le gustar[ií]a|H[aá]game saber|No dude en)[^\n]*(?:\n[\s\S]*)?$/i, '')
    // Strip "Total results:" / "Results:" summary lines (en/fr/it/de/es)
    .replace(/\n*\s*(?:Total results|Results|R[eé]sultats(?: totaux)?|Risultati(?: totali)?|(?:Gesamt)?[Ee]rgebnisse|Resultados(?: totales)?):?\s*[^\n]*/gi, '')
    // Remove emoji
    .replace(EMOJI_PATTERN, '')
    // Strip Outlook [N](url) reference links entirely (Agent 365 appends 20+ footnote-style links)
    .replace(/\[\d+\]\(https?:\/\/outlook\.office(?:365)?\.com\/[^)]*\)/g, '')
    // Strip Teams URLs (not useful in SharePoint context)
    .replace(/https?:\/\/teams\.microsoft\.com\/[^\s)>\]]+/g, '')
    // Clean up empty markdown links: []() or [n]()
    .replace(/\[\d*\]\(\s*\)/g, '')
    // Collapse leftover large numbers on their own line
    .replace(/\n\s*\d{4,}\s*\n/g, '\n')
    // Collapse excessive blank lines
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ─── extractOutlookItemIds ────────────────────────────────────

/**
 * Scan raw Agent 365 reply for Outlook links like [N](https://outlook.office365.com/owa/?ItemID=AAMk...)
 * and extract the ItemID values (which are valid Graph message IDs for GetMessage).
 * Returns a map of { cardNumber → itemId }.
 */
export function extractOutlookItemIds(rawReply: string): Record<number, string> {
  const result: Record<number, string> = {};
  OUTLOOK_ITEM_LINK_PATTERN.lastIndex = 0;
  const pattern = OUTLOOK_ITEM_LINK_PATTERN;
  let m = pattern.exec(rawReply);
  while (m !== null) {
    const num = parseInt(m[1], 10);
    // Keep URL-encoded form: %2F stays as %2F so the MCP server won't
    // split on '/' when constructing the Graph API URL path.
    const itemId = m[2];
    if (num > 0 && itemId) {
      logService.debug('mcp', 'Extracted ItemID #' + num + ' len=' + itemId.length + ' id=' + itemId.substring(0, 80) + (itemId.length > 80 ? '...' : ''));
      result[num] = itemId;
    }
    m = pattern.exec(rawReply);
  }
  return result;
}

export function prepareAgentReplyForDisplay(rawReply: string): IPreparedAgentReply {
  const deduped = dedupeOutlookReply(rawReply);
  return {
    cleanedReply: cleanAgentReply(deduped.reply),
    itemIds: deduped.itemIds
  };
}
