import type { IMcpTargetContext } from '../mcp/McpTargetContext';
import type { IBlock, ISiteInfoData } from '../../models/IBlock';

export type ContextualVisibleAction = 'summarize' | 'preview' | 'permissions';
export type ContextualMutationAction = 'rename';
export type ContextualContainerAction = 'browse-library' | 'show-list-items' | 'list-lists';

export interface IContextualContainerTarget {
  siteUrl?: string;
  siteName?: string;
  libraryName?: string;
  listName?: string;
}

export const CONTEXTUAL_COMPOSE_ACTION_HINTS: ReadonlyArray<RegExp> = [
  /\bsend\b/i,
  /\bshare\b/i,
  /\bemail\b/i,
  /\bmail\b/i,
  /\bpost\b/i,
  /\bcompose\b/i,
  // fr
  /\benvoyer\b/i,
  /\bpartager\b/i,
  /\bcourrier\b/i,
  /\br[eé]diger\b/i,
  // it
  /\binviare\b/i,
  /\bcondividere\b/i,
  /\bposta\b/i,
  /\bcomporre\b/i,
  // de
  /\bsenden\b/i,
  /\bteilen\b/i,
  /\bverfassen\b/i,
  // es
  /\benviar\b/i,
  /\bcompartir\b/i,
  /\bcorreo\b/i,
  /\bredactar\b/i
];

export const CONTEXTUAL_EMAIL_HINTS: ReadonlyArray<RegExp> = [
  /\bemail\b/i,
  /\bmail\b/i,
  /\be-mail\b/i,
  /\bcourrier\b/i,       // fr
  /\bcourriel\b/i,       // fr
  /\bposta\b/i,          // it
  /\bcorreo\b/i          // es
];

export const CONTEXTUAL_TEAMS_HINTS: ReadonlyArray<RegExp> = [
  /\bteams?\b/i
];

export const CONTEXTUAL_CHANNEL_HINTS: ReadonlyArray<RegExp> = [
  /\bchannel(?:s)?\b/i,
  /\bcanal\b/i,          // fr/es
  /\bcanaux\b/i,         // fr plural
  /\bcanale\b/i,         // it
  /\bcanali\b/i,         // it plural
  /\bkanal\b/i,          // de
  /\bkan[aä]le\b/i,      // de plural
  /\bcanales\b/i         // es plural
];

export const CONTEXTUAL_CHAT_HINTS: ReadonlyArray<RegExp> = [
  /\bchat\b/i,
  /\bdm\b/i,
  /\bdirect message\b/i,
  /\bmessagerie\b/i,         // fr
  /\bdiscussion\b/i,         // fr
  /\bmessaggio diretto\b/i,  // it
  /\bdirekt(?:nachricht|e nachricht)\b/i, // de
  /\bmensaje directo\b/i     // es
];

export const CONTEXTUAL_SHARE_TARGET_HINTS: ReadonlyArray<RegExp> = [
  /\bresult(?:s)?\b/i,
  /\bdocument(?:s)?\b/i,
  /\bdoc(?:s)?\b/i,
  /\bfile(?:s)?\b/i,
  /\bpdf(?:s)?\b/i,
  /\bitem(?:s)?\b/i,
  /\bthis\b/i,
  /\bthat\b/i,
  /\bthese\b/i,
  /\bthem\b/i,
  // fr
  /\br[eé]sultat(?:s)?\b/i,
  /\bfichier(?:s)?\b/i,
  /\b[eé]l[eé]ment(?:s)?\b/i,
  /\bceci\b/i,
  /\bcela\b/i,
  // it
  /\brisultat[oi]\b/i,
  /\bdocument[oi]\b/i,
  /\belement[oi]\b/i,
  /\bquest[oi]\b/i,
  // de
  /\bergebnis(?:se)?\b/i,
  /\bdokument(?:e)?\b/i,
  /\bdatei(?:en)?\b/i,
  /\bdiese?\b/i,
  // es
  /\bresultado(?:s)?\b/i,
  /\barchivo(?:s)?\b/i,
  /\belemento(?:s)?\b/i,
  /\best[oe]s?\b/i
];

export const CONTEXTUAL_SUMMARIZE_HINTS: ReadonlyArray<RegExp> = [
  /\bsummarize\b/i,
  /\bsummary\b/i,
  /\brecap\b/i,
  /\br[eé]sumer\b/i,           // fr
  /\br[eé]sum[eé]\b/i,         // fr
  /\briassumere\b/i,           // it
  /\briassunto\b/i,            // it
  /\bzusammenfassen\b/i,       // de
  /\bzusammenfassung\b/i,      // de
  /\bresumi[r]\b/i,            // es
  /\bresumen\b/i               // es
];

export const CONTEXTUAL_PREVIEW_HINTS: ReadonlyArray<RegExp> = [
  /\bpreview\b/i,
  /\bopen\b/i,
  /\bread\b/i,
  /\bshow\b/i,
  /\bdetails?\b/i,
  // fr
  /\baper[cç]u\b/i,
  /\bouvrir\b/i,
  /\blire\b/i,
  /\bafficher\b/i,
  /\bd[eé]tails?\b/i,
  // it
  /\banteprima\b/i,
  /\baprire\b/i,
  /\bleggere\b/i,
  /\bmostrare\b/i,
  /\bdettagli[o]?\b/i,
  // de
  /\bvorschau\b/i,
  /\b[oö]ffnen\b/i,
  /\blesen\b/i,
  /\banzeigen\b/i,
  // es
  /\bvista previa\b/i,
  /\babrir\b/i,
  /\bleer\b/i,
  /\bmostrar\b/i,
  /\bdetalle(?:s)?\b/i
];

export const CONTEXTUAL_PERMISSION_HINTS: ReadonlyArray<RegExp> = [
  /\bpermissions?\b/i,
  /\baccess\b/i,
  /\bwho can\b/i,
  /\bsharing\b/i,
  // fr
  /\bautorisations?\b/i,
  /\bacc[eè]s\b/i,
  /\bqui peut\b/i,
  /\bpartage\b/i,
  // it
  /\bpermess[oi]\b/i,
  /\baccesso\b/i,
  /\bchi pu[oò]\b/i,
  /\bcondivisione\b/i,
  // de
  /\bberechtigungen?\b/i,
  /\bzugriff\b/i,
  /\bwer (?:kann|darf)\b/i,
  /\bfreigabe\b/i,
  // es
  /\bpermisos?\b/i,
  /\bacceso\b/i,
  /\bqui[eé]n puede\b/i
];

export const CONTEXTUAL_RENAME_HINTS: ReadonlyArray<RegExp> = [
  /\brename\b/i,
  /\bchange (?:the )?name\b/i,
  /\brenommer\b/i,             // fr
  /\bchanger le nom\b/i,       // fr
  /\brinominare\b/i,           // it
  /\bcambiare (?:il )?nome\b/i, // it
  /\bumbenennen\b/i,           // de
  /\brenombrar\b/i,            // es
  /\bcambiar (?:el )?nombre\b/i // es
];

export const CONTEXTUAL_COLUMN_CREATE_HINTS: ReadonlyArray<RegExp> = [
  /\badd\b/i,
  /\bcreate\b/i,
  /\bnew\b/i,
  /\binsert\b/i,
  // fr
  /\bajouter\b/i,
  /\bcr[eé]er\b/i,
  /\bnouvelle?\b/i,
  // it
  /\baggiungere\b/i,
  /\bcreare\b/i,
  /\bnuov[oa]\b/i,
  // de
  /\bhinzuf[uü]gen\b/i,
  /\berstellen\b/i,
  /\bneue?[nrs]?\b/i,
  // es
  /\bagregar\b/i,
  /\ba[nñ]adir\b/i,
  /\bcrear\b/i,
  /\bnuev[oa]\b/i
];

export const CONTEXTUAL_COLUMN_NOUN_HINTS: ReadonlyArray<RegExp> = [
  /\bcolumns?\b/i,
  /\bfield(?:s)?\b/i,
  // fr
  /\bcolonnes?\b/i,
  /\bchamp(?:s)?\b/i,
  // it
  /\bcolonn[ae]\b/i,
  /\bcamp[oi]\b/i,
  // de
  /\bspalten?\b/i,
  /\bfeld(?:er)?\b/i,
  // es
  /\bcolumnas?\b/i,
  /\bcampos?\b/i
];

export const CONTEXTUAL_SELECTED_REFERENCE_HINTS: ReadonlyArray<RegExp> = [
  /\bthe one i selected\b/i,
  /\bthe (?:file|folder|document|item) i selected\b/i,
  /\bselected (?:file|folder|document|item)\b/i,
  /\bthat one\b/i,
  // fr
  /\bcelui que j'ai s[eé]lectionn[eé]\b/i,
  /\bcelui-l[aà]\b/i,
  // it
  /\bquello (?:che ho )?selezionat[oi]\b/i,
  /\bquello l[aà]\b/i,
  // de
  /\b(?:die|das|den) ausgew[aä]hlte\b/i,
  /\bdieses?\b/i,
  // es
  /\bel que seleccion[eé]\b/i,
  /\b[eé]se\b/i
];

export const CONTEXTUAL_RENAME_CLARIFICATION_HINTS: ReadonlyArray<RegExp> = [
  /\bwhich item should i rename\b/i,
  /\btell me the item number\b/i,
  /\bwhat new name\b/i
];

export const CONTEXTUAL_LIST_DISCOVERY_HINTS: ReadonlyArray<RegExp> = [
  /\ball lists?\b/i,
  /\bshow (?:me )?(?:all )?lists?\b/i,
  /\blist (?:all )?lists?\b/i,
  /\bwhat lists?\b/i,
  // fr
  /\btoutes? les listes?\b/i,
  /\bafficher les listes?\b/i,
  /\bquelles? listes?\b/i,
  // it
  /\btutt[ie] (?:gli |le )?elenchi?\b/i,
  /\bmostrare (?:gli |le )?elenchi?\b/i,
  /\bquali elenchi?\b/i,
  // de
  /\balle listen?\b/i,
  /\blisten? anzeigen\b/i,
  /\bwelche listen?\b/i,
  // es
  /\btodas? las listas?\b/i,
  /\bmostrar (?:las )?listas?\b/i,
  /\bqu[eé] listas?\b/i
];

export const CONTEXTUAL_LIST_CONTENT_HINTS: ReadonlyArray<RegExp> = [
  /\bitems?\b/i,
  /\brows?\b/i,
  /\bentries\b/i,
  /\brecords\b/i,
  /\bcontent\b/i,
  // fr
  /\b[eé]l[eé]ments?\b/i,
  /\blignes?\b/i,
  /\bentr[eé]es?\b/i,
  /\bcontenu\b/i,
  // it
  /\belement[oi]\b/i,
  /\brigh[ei]\b/i,
  /\bvoc[ie]\b/i,
  /\bcontenut[oi]\b/i,
  // de
  /\belemente?\b/i,
  /\bzeilen?\b/i,
  /\beintr[aä]ge?\b/i,
  /\binhalt\b/i,
  // es
  /\belementos?\b/i,
  /\bfilas?\b/i,
  /\bregistros?\b/i,
  /\bcontenido\b/i
];

export const CONTEXTUAL_LIBRARY_HINTS: ReadonlyArray<RegExp> = [
  /\bdocument library\b/i,
  /\blibrary\b/i,
  /\bdocuments\b/i,
  /\bfiles?\b/i,
  /\bfolders?\b/i,
  /\bcontent\b/i,
  // fr
  /\bbiblioth[eè]que\b/i,
  /\bfichiers?\b/i,
  /\bdossiers?\b/i,
  // it
  /\braccolta\b/i,
  /\bdocument[oi]\b/i,
  /\bcartell[ae]\b/i,
  // de
  /\bbibliothek\b/i,
  /\bdokumente?\b/i,
  /\bdateien?\b/i,
  /\bordner\b/i,
  // es
  /\bbiblioteca\b/i,
  /\barchivos?\b/i,
  /\bcarpetas?\b/i
];

export const WHICH_ONE_PROMPT_HINTS: ReadonlyArray<RegExp> = [
  /\bwhich one\b/i,
  /\bwhich (?:result|document|doc|file)\b/i,
  /\bpick one\b/i,
  /\bchoose one\b/i,
  // fr
  /\blequel\b/i,
  /\blaquelle\b/i,
  /\bquel (?:r[eé]sultat|document|fichier)\b/i,
  // it
  /\bquale\b/i,
  /\bquale (?:risultato|documento|file)\b/i,
  // de
  /\bwelche[rs]?\b/i,
  /\bwelche[rs]? (?:ergebnis|dokument|datei)\b/i,
  // es
  /\bcu[aá]l\b/i,
  /\bcu[aá]l (?:resultado|documento|archivo)\b/i
];

export const ORDINAL_INDEX_PATTERNS: ReadonlyArray<{ pattern: RegExp; index: number }> = [
  { pattern: /\b(?:first|premi(?:er|[eè]re)|prim[oa]|erste[rs]?|primer[oa]?)\b/i, index: 1 },
  { pattern: /\b(?:second|deuxi[eè]me|second[oa]|zweite[rs]?|segund[oa]?)\b/i, index: 2 },
  { pattern: /\b(?:third|troisi[eè]me|terz[oa]|dritte[rs]?|tercer[oa]?)\b/i, index: 3 },
  { pattern: /\b(?:fourth|quatri[eè]me|quart[oa]|vierte[rs]?|cuart[oa]?)\b/i, index: 4 },
  { pattern: /\b(?:fifth|cinqui[eè]me|quint[oa]|f[uü]nfte[rs]?|quint[oa]?)\b/i, index: 5 },
  { pattern: /\b(?:sixth|sixi[eè]me|sest[oa]|sechste[rs]?|sext[oa]?)\b/i, index: 6 },
  { pattern: /\b(?:seventh|septi[eè]me|settim[oa]|siebte[rs]?|s[eé]ptim[oa]?)\b/i, index: 7 },
  { pattern: /\b(?:eighth|huiti[eè]me|ottav[oa]|achte[rs]?|octav[oa]?)\b/i, index: 8 },
  { pattern: /\b(?:ninth|neuvi[eè]me|non[oa]|neunte[rs]?|noven[oa]?)\b/i, index: 9 },
  { pattern: /\b(?:tenth|dixi[eè]me|decim[oa]|zehnte[rs]?|d[eé]cim[oa]?)\b/i, index: 10 }
];

export interface IContextualVisibleItemMatch {
  block: IBlock;
  index: number;
  title: string;
  url: string;
  targetContext?: IMcpTargetContext;
}

export interface IContextualComposeOverrideDetails {
  description: string;
  staticArgsJson?: string;
}

export function matchesAny(text: string, patterns: ReadonlyArray<RegExp>): boolean {
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(text)) {
      return true;
    }
  }
  return false;
}

export function normalizeVisibleContextText(text: string): string {
  return text.trim().replace(/\s+/g, ' ');
}

export function normalizeVisibleItemTitleText(text: string): string {
  return text
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/[^0-9a-z\u00c0-\u024f\u0400-\u04ff\u0600-\u06ff\u0590-\u05ff\u3040-\u30ff\u4e00-\u9fff\uac00-\ud7af\u0e00-\u0e7f]+/gi, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

export function blockDataAsSiteInfo(block: IBlock): ISiteInfoData | undefined {
  if (block.type !== 'site-info') {
    return undefined;
  }
  return block.data as ISiteInfoData;
}
