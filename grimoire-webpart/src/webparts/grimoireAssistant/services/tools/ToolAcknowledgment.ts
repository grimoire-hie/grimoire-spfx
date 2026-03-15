/**
 * ToolAcknowledgment
 * Deterministic, local acknowledgments for tool dispatch start.
 */

import { normalizeConversationLanguage } from '../context/ConversationLanguage';

function readQuery(args: Record<string, unknown>): string | undefined {
  const raw = typeof args.query === 'string' ? args.query.trim() : '';
  if (!raw || raw === '*') return undefined;
  return raw;
}

function quote(value: string): string {
  return `"${value}"`;
}

function readTargetUrl(args: Record<string, unknown>): string | undefined {
  const raw = typeof args.target_url === 'string' ? args.target_url.trim() : '';
  return raw || undefined;
}

function countFileTargets(args: Record<string, unknown>): number {
  const raw = args.file_urls;
  if (!Array.isArray(raw)) return 0;
  let count = 0;
  for (let i = 0; i < raw.length; i++) {
    if (typeof raw[i] !== 'string') continue;
    if (raw[i].trim()) count++;
  }
  return count;
}

export function getToolAckText(
  toolName: string,
  args: Record<string, unknown>,
  language?: string
): string | undefined {
  const query = readQuery(args);
  const targetUrl = readTargetUrl(args);
  const locale = normalizeConversationLanguage(language) || 'en';

  switch (toolName) {
    case 'research_public_web':
      if (locale === 'fr') {
        return targetUrl
          ? `D'accord, je vérifie ${quote(targetUrl)}.`
          : query
            ? `D'accord, je recherche sur le web public ${quote(query)}.`
            : `D'accord, je recherche sur le web public.`;
      }
      if (locale === 'it') {
        return targetUrl
          ? `Va bene, controllo ${quote(targetUrl)}.`
          : query
            ? `Va bene, cerco sul web pubblico ${quote(query)}.`
            : `Va bene, cerco sul web pubblico.`;
      }
      if (locale === 'de') {
        return targetUrl
          ? `In Ordnung, ich prüfe ${quote(targetUrl)}.`
          : query
            ? `In Ordnung, ich recherchiere im öffentlichen Web nach ${quote(query)}.`
            : `In Ordnung, ich recherchiere im öffentlichen Web.`;
      }
      if (locale === 'es') {
        return targetUrl
          ? `De acuerdo, reviso ${quote(targetUrl)}.`
          : query
            ? `De acuerdo, busco en la web pública ${quote(query)}.`
            : `De acuerdo, busco en la web pública.`;
      }
      return targetUrl
        ? `Okay, I'm checking ${quote(targetUrl)}.`
        : query
          ? `Okay, I'm researching the public web for ${quote(query)}.`
          : `Okay, I'm researching the public web.`;
    case 'search_sharepoint':
      if (locale === 'fr') {
        return query
          ? `D'accord, je cherche dans SharePoint ${quote(query)}.`
          : `D'accord, je cherche dans SharePoint.`;
      }
      if (locale === 'it') {
        return query
          ? `Va bene, cerco in SharePoint ${quote(query)}.`
          : `Va bene, cerco in SharePoint.`;
      }
      if (locale === 'de') {
        return query
          ? `In Ordnung, ich suche in SharePoint nach ${quote(query)}.`
          : `In Ordnung, ich suche in SharePoint.`;
      }
      if (locale === 'es') {
        return query
          ? `De acuerdo, busco en SharePoint ${quote(query)}.`
          : `De acuerdo, busco en SharePoint.`;
      }
      return query
        ? `Okay, I'm searching SharePoint for ${quote(query)}.`
        : `Okay, I'm searching SharePoint.`;
    case 'search_emails':
      if (locale === 'fr') {
        return query
          ? `D'accord, je récupère vos e-mails pour ${quote(query)}.`
          : `D'accord, je récupère vos e-mails.`;
      }
      if (locale === 'it') {
        return query
          ? `Va bene, recupero le tue e-mail per ${quote(query)}.`
          : `Va bene, recupero le tue e-mail.`;
      }
      if (locale === 'de') {
        return query
          ? `In Ordnung, ich hole deine E-Mails zu ${quote(query)}.`
          : `In Ordnung, ich hole deine E-Mails.`;
      }
      if (locale === 'es') {
        return query
          ? `De acuerdo, recupero tus correos sobre ${quote(query)}.`
          : `De acuerdo, recupero tus correos.`;
      }
      return query
        ? `Okay, I'm pulling your emails for ${quote(query)}.`
        : `Okay, I'm pulling your emails.`;
    case 'search_people':
      if (locale === 'fr') {
        return query
          ? `D'accord, je recherche des personnes pour ${quote(query)}.`
          : `D'accord, je recherche des personnes.`;
      }
      if (locale === 'it') {
        return query
          ? `Va bene, cerco persone per ${quote(query)}.`
          : `Va bene, cerco persone.`;
      }
      if (locale === 'de') {
        return query
          ? `In Ordnung, ich suche Personen zu ${quote(query)}.`
          : `In Ordnung, ich suche Personen.`;
      }
      if (locale === 'es') {
        return query
          ? `De acuerdo, busco personas para ${quote(query)}.`
          : `De acuerdo, busco personas.`;
      }
      return query
        ? `Okay, I'm looking up people for ${quote(query)}.`
        : `Okay, I'm looking up people.`;
    case 'search_sites':
      if (locale === 'fr') {
        return query
          ? `D'accord, je cherche des sites pour ${quote(query)}.`
          : `D'accord, je cherche des sites.`;
      }
      if (locale === 'it') {
        return query
          ? `Va bene, cerco siti per ${quote(query)}.`
          : `Va bene, cerco siti.`;
      }
      if (locale === 'de') {
        return query
          ? `In Ordnung, ich suche Websites zu ${quote(query)}.`
          : `In Ordnung, ich suche Websites.`;
      }
      if (locale === 'es') {
        return query
          ? `De acuerdo, busco sitios para ${quote(query)}.`
          : `De acuerdo, busco sitios.`;
      }
      return query
        ? `Okay, I'm searching sites for ${quote(query)}.`
        : `Okay, I'm searching sites.`;
    case 'read_file_content':
      if (locale === 'fr') {
        return countFileTargets(args) > 1
          ? `D'accord, je lis ces fichiers.`
          : `D'accord, je lis ce fichier.`;
      }
      if (locale === 'it') {
        return countFileTargets(args) > 1
          ? `Va bene, leggo quei file.`
          : `Va bene, leggo quel file.`;
      }
      if (locale === 'de') {
        return countFileTargets(args) > 1
          ? `In Ordnung, ich lese diese Dateien.`
          : `In Ordnung, ich lese diese Datei.`;
      }
      if (locale === 'es') {
        return countFileTargets(args) > 1
          ? `De acuerdo, leo esos archivos.`
          : `De acuerdo, leo ese archivo.`;
      }
      return countFileTargets(args) > 1
        ? `Okay, I'm reading those files.`
        : `Okay, I'm reading that file.`;
    case 'read_email_content':
      if (locale === 'fr') return `D'accord, je lis cet e-mail.`;
      if (locale === 'it') return `Va bene, leggo quella e-mail.`;
      if (locale === 'de') return `In Ordnung, ich lese diese E-Mail.`;
      if (locale === 'es') return `De acuerdo, leo ese correo.`;
      return `Okay, I'm reading that email.`;
    case 'read_teams_messages':
      if (locale === 'fr') return `D'accord, je lis les messages Teams.`;
      if (locale === 'it') return `Va bene, leggo i messaggi di Teams.`;
      if (locale === 'de') return `In Ordnung, ich lese Teams-Nachrichten.`;
      if (locale === 'es') return `De acuerdo, leo los mensajes de Teams.`;
      return `Okay, I'm reading Teams messages.`;
    default:
      return undefined;
  }
}

const EXPLICIT_SELECTION_LIST_PATTERNS: ReadonlyArray<RegExp> = [
  /\bshow options\b/i,
  /\bshow me options\b/i,
  /\boptions list\b/i,
  /\bselection list\b/i,
  /\bradio buttons?\b/i,
  /\blist to choose\b/i,
  /\bchoose from (a )?list\b/i,
  /\bgive me (a )?list to choose\b/i
];

export function isExplicitSelectionListRequest(userText: string): boolean {
  if (!userText || !userText.trim()) return false;
  return EXPLICIT_SELECTION_LIST_PATTERNS.some((pattern) => pattern.test(userText));
}
