/**
 * ToolCompletionAcknowledgment
 * Deterministic, local acknowledgments for async tool completion.
 */

import { detectCapabilityFocus } from '../../models/McpServerCatalog';
import { normalizeConversationLanguage } from '../context/ConversationLanguage';

const LOCAL_COMPLETION_ACK_TOOLS: ReadonlySet<string> = new Set([
  'research_public_web',
  'search_sharepoint',
  'search_people',
  'search_sites',
  'search_emails',
  'browse_document_library',
  'show_file_details',
  'show_site_info',
  'show_list_items',
  'list_m365_servers',
  'get_my_profile',
  'get_recent_documents',
  'get_trending_documents',
  'recall_notes'
]);

const SEARCH_LIKE_COMPLETION_TOOLS: ReadonlySet<string> = new Set([
  'research_public_web',
  'search_sharepoint',
  'search_people',
  'search_sites',
  'search_emails',
  'browse_document_library',
  'get_recent_documents',
  'get_trending_documents',
  'recall_notes'
]);

const CHAINING_HINTS: ReadonlyArray<RegExp> = [
  // English
  /\band\b/i,
  /\bthen\b/i,
  /\balso\b/i,
  /\bafter that\b/i,
  /\bafterwards\b/i,
  // French
  /\bet\b/i,
  /\bpuis\b/i,
  /\baussi\b/i,
  /\bensuite\b/i,
  /\bapr[eè]s\b/i,
  // Italian
  /\bpoi\b/i,
  /\banche\b/i,
  /\bdopo\b/i,
  /\bquindi\b/i,
  // German
  /\bund\b/i,
  /\bdann\b/i,
  /\bauch\b/i,
  /\bdanach\b/i,
  /\banschlie[sß]end\b/i,
  // Spanish
  /\bluego\b/i,
  /\btambi[eé]n\b/i,
  /\bdespu[eé]s\b/i
];

const POST_SEARCH_ACTION_HINTS: ReadonlyArray<RegExp> = [
  // English
  /\bsummarize\b/i,
  /\bsummary\b/i,
  /\bpreview\b/i,
  /\bopen\b/i,
  /\bread\b/i,
  /\bcompare\b/i,
  /\banaly[sz]e\b/i,
  /\bexplain\b/i,
  /\bshow details?\b/i,
  /\btell me about\b/i,
  // French
  /\br[eé]sumer\b/i,
  /\br[eé]sum[eé]\b/i,
  /\baper[cç]u\b/i,
  /\bouvrir\b/i,
  /\blire\b/i,
  /\bcomparer\b/i,
  /\banalyser\b/i,
  /\bexpliquer\b/i,
  /\bafficher\b/i,
  // Italian
  /\briassumere\b/i,
  /\briassunto\b/i,
  /\banteprima\b/i,
  /\baprire\b/i,
  /\bleggere\b/i,
  /\bconfrontare\b/i,
  /\banalizzare\b/i,
  /\bspiegare\b/i,
  /\bmostrare\b/i,
  // German
  /\bzusammenfassen\b/i,
  /\bzusammenfassung\b/i,
  /\bvorschau\b/i,
  /\b[oö]ffnen\b/i,
  /\blesen\b/i,
  /\bvergleichen\b/i,
  /\banalysieren\b/i,
  /\berkl[aä]ren\b/i,
  /\banzeigen\b/i,
  // Spanish
  /\bresumi[r]\b/i,
  /\bresumen\b/i,
  /\bvista previa\b/i,
  /\babrir\b/i,
  /\bleer\b/i,
  /\bcomparar\b/i,
  /\banalizar\b/i,
  /\bexplicar\b/i,
  /\bmostrar\b/i
];

const COUNT_KEYS: ReadonlyArray<string> = [
  'itemCount',
  'count',
  'displayedResults',
  'totalCount',
  'permissionCount',
  'activityCount',
  'deleted'
];

function pluralize(count: number, singular: string, plural: string): string {
  return `${count} ${count === 1 ? singular : plural}`;
}

function matchesAny(text: string, patterns: ReadonlyArray<RegExp>): boolean {
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(text)) return true;
  }
  return false;
}

function readNumericValue(value: unknown): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === 'string' && /^\d+$/.test(value.trim())) {
    return parseInt(value, 10);
  }
  return undefined;
}

function extractCompletionItemCount(payload: Record<string, unknown>): number {
  for (let i = 0; i < COUNT_KEYS.length; i++) {
    const value = readNumericValue(payload[COUNT_KEYS[i]]);
    if (value !== undefined) {
      return value;
    }
  }

  const arrayKeys: ReadonlyArray<string> = ['results', 'items', 'people', 'notes', 'permissions', 'activities'];
  for (let i = 0; i < arrayKeys.length; i++) {
    const value = payload[arrayKeys[i]];
    if (Array.isArray(value)) {
      return value.length;
    }
  }

  return 0;
}

export function getToolCompletionAckText(
  toolName: string,
  itemCount: number,
  language?: string
): string | undefined {
  const locale = normalizeConversationLanguage(language) || 'en';

  switch (toolName) {
    case 'research_public_web':
      if (locale === 'fr') return 'J’ai résumé les résultats du web public dans le panneau.';
      if (locale === 'it') return 'Ho riassunto i risultati del web pubblico nel pannello.';
      if (locale === 'de') return 'Ich habe die Ergebnisse aus dem öffentlichen Web im Panel zusammengefasst.';
      if (locale === 'es') return 'He resumido los resultados de la web pública en el panel.';
      return 'I summarized the public web results in the panel.';
    case 'search_sharepoint':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai trouvé ${pluralize(itemCount, 'document', 'documents')}. Ils sont dans le panneau.`
          : 'Je n’ai trouvé aucun document correspondant.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho trovato ${pluralize(itemCount, 'documento', 'documenti')}. Sono nel pannello.`
          : 'Non ho trovato documenti corrispondenti.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Dokument', 'Dokumente')} gefunden. Sie sind im Panel.`
          : 'Ich habe keine passenden Dokumente gefunden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `Encontré ${pluralize(itemCount, 'documento', 'documentos')}. Están en el panel.`
          : 'No encontré documentos coincidentes.';
      }
      return itemCount > 0
        ? `I found ${pluralize(itemCount, 'document', 'documents')}. They are in the panel.`
        : 'I did not find matching documents.';
    case 'search_people':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai trouvé ${pluralize(itemCount, 'personne', 'personnes')}. Les cartes sont dans le panneau.`
          : 'Je n’ai trouvé aucune personne correspondante.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho trovato ${pluralize(itemCount, 'persona', 'persone')}. Le schede sono nel pannello.`
          : 'Non ho trovato persone corrispondenti.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Person', 'Personen')} gefunden. Die Karten sind im Panel.`
          : 'Ich habe keine passenden Personen gefunden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `Encontré ${pluralize(itemCount, 'persona', 'personas')}. Las tarjetas están en el panel.`
          : 'No encontré personas coincidentes.';
      }
      return itemCount > 0
        ? `I found ${pluralize(itemCount, 'person', 'people')}. The cards are in the panel.`
        : 'I did not find matching people.';
    case 'search_sites':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai trouvé ${pluralize(itemCount, 'site', 'sites')}. Ils sont dans le panneau.`
          : 'Je n’ai trouvé aucun site correspondant.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho trovato ${pluralize(itemCount, 'sito', 'siti')}. Sono nel pannello.`
          : 'Non ho trovato siti corrispondenti.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Website', 'Websites')} gefunden. Sie sind im Panel.`
          : 'Ich habe keine passenden Websites gefunden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `Encontré ${pluralize(itemCount, 'sitio', 'sitios')}. Están en el panel.`
          : 'No encontré sitios coincidentes.';
      }
      return itemCount > 0
        ? `I found ${pluralize(itemCount, 'site', 'sites')}. They are in the panel.`
        : 'I did not find matching sites.';
    case 'search_emails':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai trouvé ${pluralize(itemCount, 'e-mail', 'e-mails')}. Ils sont dans le panneau.`
          : 'Je n’ai trouvé aucun e-mail correspondant.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho trovato ${pluralize(itemCount, 'e-mail', 'e-mail')}. Sono nel pannello.`
          : 'Non ho trovato e-mail corrispondenti.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'E-Mail', 'E-Mails')} gefunden. Sie sind im Panel.`
          : 'Ich habe keine passenden E-Mails gefunden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `Encontré ${pluralize(itemCount, 'correo', 'correos')}. Están en el panel.`
          : 'No encontré correos coincidentes.';
      }
      return itemCount > 0
        ? `I found ${pluralize(itemCount, 'email', 'emails')}. They are in the panel.`
        : 'I did not find matching emails.';
    case 'browse_document_library':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai chargé ${pluralize(itemCount, 'élément', 'éléments')} de la bibliothèque.`
          : 'J’ai chargé la bibliothèque, mais elle semble vide.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho caricato ${pluralize(itemCount, 'elemento', 'elementi')} dalla raccolta.`
          : 'Ho caricato la raccolta, ma sembra vuota.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Element', 'Elemente')} aus der Bibliothek geladen.`
          : 'Ich habe die Bibliothek geladen, aber sie wirkt leer.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `He cargado ${pluralize(itemCount, 'elemento', 'elementos')} de la biblioteca.`
          : 'He cargado la biblioteca, pero parece vacía.';
      }
      return itemCount > 0
        ? `I loaded ${pluralize(itemCount, 'item', 'items')} from the library.`
        : 'I loaded the library, but it looks empty.';
    case 'show_file_details':
      if (locale === 'fr') return 'J’ai chargé les détails du fichier dans le panneau.';
      if (locale === 'it') return 'Ho caricato i dettagli del file nel pannello.';
      if (locale === 'de') return 'Ich habe die Dateidetails im Panel geladen.';
      if (locale === 'es') return 'He cargado los detalles del archivo en el panel.';
      return 'I loaded the file details in the panel.';
    case 'show_site_info':
      if (locale === 'fr') return 'J’ai chargé les détails du site dans le panneau.';
      if (locale === 'it') return 'Ho caricato i dettagli del sito nel pannello.';
      if (locale === 'de') return 'Ich habe die Sitedetails im Panel geladen.';
      if (locale === 'es') return 'He cargado los detalles del sitio en el panel.';
      return 'I loaded the site details in the panel.';
    case 'show_list_items':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai chargé ${pluralize(itemCount, 'élément de liste', 'éléments de liste')} dans le panneau.`
          : 'J’ai chargé la liste, mais aucun élément n’a été renvoyé.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho caricato ${pluralize(itemCount, 'elemento di elenco', 'elementi di elenco')} nel pannello.`
          : 'Ho caricato l’elenco, ma non è stato restituito alcun elemento.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Listenelement', 'Listenelemente')} im Panel geladen.`
          : 'Ich habe die Liste geladen, aber es wurden keine Elemente zurückgegeben.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `He cargado ${pluralize(itemCount, 'elemento de lista', 'elementos de lista')} en el panel.`
          : 'He cargado la lista, pero no se devolvió ningún elemento.';
      }
      return itemCount > 0
        ? `I loaded ${pluralize(itemCount, 'list item', 'list items')} in the panel.`
        : 'I loaded the list, but no items were returned.';
    case 'list_m365_servers':
      if (locale === 'fr') return 'J’ai ouvert une vue d’ensemble des capacités dans le panneau.';
      if (locale === 'it') return 'Ho aperto una panoramica delle capacità nel pannello.';
      if (locale === 'de') return 'Ich habe eine Funktionsübersicht im Panel geöffnet.';
      if (locale === 'es') return 'He abierto un resumen de capacidades en el panel.';
      return 'I opened a capability overview in the panel.';
    case 'get_my_profile':
      if (locale === 'fr') return 'J’ai chargé votre profil dans le panneau.';
      if (locale === 'it') return 'Ho caricato il tuo profilo nel pannello.';
      if (locale === 'de') return 'Ich habe dein Profil im Panel geladen.';
      if (locale === 'es') return 'He cargado tu perfil en el panel.';
      return 'I loaded your profile in the panel.';
    case 'get_recent_documents':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai chargé ${pluralize(itemCount, 'document récent', 'documents récents')}.`
          : 'Je ne trouve pas de documents récents pour le moment.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho caricato ${pluralize(itemCount, 'documento recente', 'documenti recenti')}.`
          : 'Non riesco a trovare documenti recenti in questo momento.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'aktuelles Dokument', 'aktuelle Dokumente')} geladen.`
          : 'Ich konnte im Moment keine aktuellen Dokumente finden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `He cargado ${pluralize(itemCount, 'documento reciente', 'documentos recientes')}.`
          : 'No pude encontrar documentos recientes ahora mismo.';
      }
      return itemCount > 0
        ? `I loaded your ${pluralize(itemCount, 'recent document', 'recent documents')}.`
        : 'I could not find recent documents right now.';
    case 'get_trending_documents':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai chargé ${pluralize(itemCount, 'document tendance', 'documents tendance')} dans le panneau.`
          : 'Je ne trouve pas de documents tendance pour le moment.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho caricato ${pluralize(itemCount, 'documento di tendenza', 'documenti di tendenza')} nel pannello.`
          : 'Non riesco a trovare documenti di tendenza in questo momento.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Trenddokument', 'Trenddokumente')} im Panel geladen.`
          : 'Ich konnte im Moment keine Trenddokumente finden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `He cargado ${pluralize(itemCount, 'documento en tendencia', 'documentos en tendencia')} en el panel.`
          : 'No pude encontrar documentos en tendencia ahora mismo.';
      }
      return itemCount > 0
        ? `I loaded ${pluralize(itemCount, 'trending document', 'trending documents')} in the panel.`
        : 'I could not find trending documents right now.';
    case 'recall_notes':
      if (locale === 'fr') {
        return itemCount > 0
          ? `J’ai trouvé ${pluralize(itemCount, 'note', 'notes')}. Elles sont dans le panneau.`
          : 'Je n’ai trouvé aucune note correspondante.';
      }
      if (locale === 'it') {
        return itemCount > 0
          ? `Ho trovato ${pluralize(itemCount, 'nota', 'note')}. Sono nel pannello.`
          : 'Non ho trovato note corrispondenti.';
      }
      if (locale === 'de') {
        return itemCount > 0
          ? `Ich habe ${pluralize(itemCount, 'Notiz', 'Notizen')} gefunden. Sie sind im Panel.`
          : 'Ich habe keine passenden Notizen gefunden.';
      }
      if (locale === 'es') {
        return itemCount > 0
          ? `Encontré ${pluralize(itemCount, 'nota', 'notas')}. Están en el panel.`
          : 'No encontré notas coincidentes.';
      }
      return itemCount > 0
        ? `I found ${pluralize(itemCount, 'note', 'notes')}. They are in the panel.`
        : 'I did not find matching notes.';
    default:
      return undefined;
  }
}

function isSuccessfulToolOutput(output: string): boolean {
  try {
    const parsed = JSON.parse(output) as Record<string, unknown>;
    return parsed.success === true;
  } catch {
    return false;
  }
}

function readComposePreset(args: Record<string, unknown>): string {
  return typeof args.preset === 'string' ? args.preset.trim() : '';
}

export function getVoiceToolCompletionAckText(
  toolName: string,
  args: Record<string, unknown>,
  output: string,
  language?: string
): string | undefined {
  if (!isSuccessfulToolOutput(output)) {
    return undefined;
  }

  if (toolName !== 'show_compose_form') {
    return undefined;
  }

  const locale = normalizeConversationLanguage(language) || 'en';
  const preset = readComposePreset(args);
  switch (preset) {
    case 'email-compose':
    case 'email-reply':
    case 'email-forward':
    case 'email-reply-all-thread':
      if (locale === 'fr') return 'J’ai ouvert le brouillon d’e-mail dans le panneau.';
      if (locale === 'it') return 'Ho aperto la bozza dell’e-mail nel pannello.';
      if (locale === 'de') return 'Ich habe den E-Mail-Entwurf im Panel geöffnet.';
      if (locale === 'es') return 'He abierto el borrador del correo en el panel.';
      return 'I opened the email draft in the panel.';
    case 'share-teams-chat':
    case 'teams-message':
      if (locale === 'fr') return 'J’ai ouvert le brouillon du chat Teams dans le panneau.';
      if (locale === 'it') return 'Ho aperto la bozza della chat di Teams nel pannello.';
      if (locale === 'de') return 'Ich habe den Teams-Chat-Entwurf im Panel geöffnet.';
      if (locale === 'es') return 'He abierto el borrador del chat de Teams en el panel.';
      return 'I opened the Teams chat draft in the panel.';
    case 'share-teams-channel':
    case 'teams-channel-message':
      if (locale === 'fr') return 'J’ai ouvert le brouillon du canal Teams dans le panneau.';
      if (locale === 'it') return 'Ho aperto la bozza del canale di Teams nel pannello.';
      if (locale === 'de') return 'Ich habe den Teams-Kanal-Entwurf im Panel geöffnet.';
      if (locale === 'es') return 'He abierto el borrador del canal de Teams en el panel.';
      return 'I opened the Teams channel draft in the panel.';
    case 'event-create':
    case 'event-update':
    case 'event-cancel':
      if (locale === 'fr') return 'J’ai ouvert le formulaire de calendrier dans le panneau.';
      if (locale === 'it') return 'Ho aperto il modulo del calendario nel pannello.';
      if (locale === 'de') return 'Ich habe das Kalenderformular im Panel geöffnet.';
      if (locale === 'es') return 'He abierto el formulario del calendario en el panel.';
      return 'I opened the calendar form in the panel.';
    default:
      if (locale === 'fr') return 'J’ai ouvert le formulaire dans le panneau.';
      if (locale === 'it') return 'Ho aperto il modulo nel pannello.';
      if (locale === 'de') return 'Ich habe das Formular im Panel geöffnet.';
      if (locale === 'es') return 'He abierto el formulario en el panel.';
      return 'I opened the form in the panel.';
  }
}

export function getToolCompletionAckFromOutput(
  toolName: string,
  output: string,
  language?: string
): string | undefined {
  let parsed: unknown;
  try {
    parsed = JSON.parse(output);
  } catch {
    return undefined;
  }

  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return undefined;
  }

  const payload = parsed as Record<string, unknown>;
  if (payload.success === false) {
    return undefined;
  }

  return getToolCompletionAckText(toolName, extractCompletionItemCount(payload), language);
}

export function shouldUseImmediateCompletionAck(
  toolName: string,
  latestUserText: string
): boolean {
  if (!LOCAL_COMPLETION_ACK_TOOLS.has(toolName)) {
    return false;
  }

  if (toolName === 'list_m365_servers' && detectCapabilityFocus(latestUserText) !== undefined) {
    return false;
  }

  if (!SEARCH_LIKE_COMPLETION_TOOLS.has(toolName)) {
    return true;
  }

  const normalizedText = latestUserText.trim();
  if (!normalizedText) {
    return true;
  }

  return !(matchesAny(normalizedText, CHAINING_HINTS) && matchesAny(normalizedText, POST_SEARCH_ACTION_HINTS));
}
