/**
 * IntentRoutingPolicy
 * Shared enterprise-first routing guidance plus observability helpers.
 *
 * Text chat uses a fast semantic classifier first and falls back to heuristics.
 * Voice remains on the heuristic/prompt-guided path for now.
 */

import type { IProxyConfig } from '../../store/useGrimoireStore';
import { FAST_INTENT_ROUTER_SYSTEM_PROMPT } from '../../config/promptCatalog';
import { detectCapabilityFocus } from '../../models/McpServerCatalog';
import type { McpCapabilityFocus } from '../../models/McpServerCatalog';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { getNanoService } from '../nano/NanoService';
import { isHiePromptMessage } from '../hie/HiePromptProtocol';
import { classifyExplicitPersonalOneDriveIntent } from './ExplicitPersonalOneDriveIntent';

export type EnterpriseRouteTarget =
  | 'none'
  | 'research_public_web'
  | 'search_sharepoint'
  | 'search_emails'
  | 'search_people'
  | 'search_sites'
  | 'list_m365_servers';

export type FirstTurnRoutingSource = 'fast' | 'heuristic';

export type FirstTurnRoutingOutcome = 'tool_call' | 'clarification' | 'answer_only';
export type FirstTurnRoutingChannel = 'text' | 'voice';

export interface IFirstTurnRoutingObservation {
  userText: string;
  normalizedText: string;
  expectedRoute: EnterpriseRouteTarget;
  expectedToolName?: string;
  expectedToolArgs?: Record<string, unknown>;
  capabilityFocus?: McpCapabilityFocus;
  isGenericEnterpriseSearch: boolean;
  source: FirstTurnRoutingSource;
  confidence: number;
}

export interface IIntentRoutingRegressionFixture {
  utterance: string;
  expectedRoute: EnterpriseRouteTarget;
  expectedToolName: string;
  expectedToolArgs?: Record<string, unknown>;
  expectsClarification: boolean;
}

const VALID_FAST_ROUTES: ReadonlySet<EnterpriseRouteTarget> = new Set<EnterpriseRouteTarget>([
  'none',
  'research_public_web',
  'search_sharepoint',
  'search_emails',
  'search_people',
  'search_sites',
  'list_m365_servers'
]);

const ACTIONABLE_SEARCH_PATTERNS: ReadonlyArray<RegExp> = [
  /\bsearch\b/i,
  /\blook up\b/i,
  /\blookup\b/i,
  /\bfind\b/i,
  /\bresearch\b/i,
  /\bsearching for\b/i,
  /\blooking for\b/i,
  /\binformation about\b/i,
  /\binformations about\b/i,
  /\binfo about\b/i,
  /\bdetails about\b/i,
  /\btell me about\b/i,
  /\bshow me information about\b/i,
  /\bcheck\b/i,
  /\bsummarize\b/i,
  /\binspect\b/i,
  /\breview\b/i,
  /\bread\b/i
];

const CAPABILITY_DISCOVERY_PATTERNS: ReadonlyArray<RegExp> = [
  /\bwhat can you do\b/i,
  /\bwhat do you offer\b/i,
  /\bwhat do you help with\b/i,
  /\bshow (?:me )?(?:your )?(?:capabilities|features)\b/i,
  /\bwhich capabilities\b/i,
  /\bwhat services are available\b/i,
  /\bhow can you help\b/i,
  /\bhelp me understand what you can do\b/i,
  /\bwas kannst du\b/i,
  /\bwas bietest du an\b/i,
  /\bwelche funktionen\b/i,
  /\bwelche möglichkeiten\b/i,
  /\bwobei kannst du helfen\b/i,
  /\bwas kannst du alles\b/i,
  /\bque peux-tu faire\b/i,
  /\bqu(?:e|')est-ce que tu peux faire\b/i,
  /\bque proposes-tu\b/i,
  /\bquelles fonctionnalit(?:e|é)s\b/i,
  /\bcomment peux-tu aider\b/i,
  /\bcosa puoi fare\b/i,
  /\bcosa offri\b/i,
  /\bquali funzionalit(?:a|à)\b/i,
  /\bin cosa puoi aiutare\b/i
];

const EXTERNAL_WEB_HINTS: ReadonlyArray<RegExp> = [
  /\bon the web\b/i,
  /\bweb\b/i,
  /\binternet\b/i,
  /\bpublic\b/i,
  /\bexternal\b/i,
  /\burl\b/i,
  /\bwebsite\b/i,
  /\bgoogle\b/i,
  /\bgithub\b/i,
  /\bwikipedia\b/i,
  /https?:\/\//i,
  /\bwww\./i
];

const EMAIL_HINTS: ReadonlyArray<RegExp> = [
  /\bemail(?:s)?\b/i,
  /\bmail\b/i,
  /\binbox\b/i,
  /\boutlook\b/i,
  /\bsubject\b/i,
  /\bsender\b/i
];

const COMPOSE_ACTION_HINTS: ReadonlyArray<RegExp> = [
  /\bsend\b/i,
  /\bcompose\b/i,
  /\bwrite\b/i,
  /\bdraft\b/i,
  /\breply\b/i,
  /\bforward\b/i,
  /\bshare\b/i
];

const CONTEXT_SHARE_TARGET_HINTS: ReadonlyArray<RegExp> = [
  /\bresult(?:s)?\b/i,
  /\bthis\b/i,
  /\bthat\b/i,
  /\bthese\b/i,
  /\bthem\b/i,
  /\bit\b/i
];

const PEOPLE_HINTS: ReadonlyArray<RegExp> = [
  /\bpeople\b/i,
  /\bperson\b/i,
  /\bemployee(?:s)?\b/i,
  /\bcolleague(?:s)?\b/i,
  /\bcontact(?:s)?\b/i,
  /\bwho is\b/i
];

const SITE_HINTS: ReadonlyArray<RegExp> = [
  /\bsite(?:s)?\b/i,
  /\bsharepoint site(?:s)?\b/i,
  /\bteam site(?:s)?\b/i,
  /\bcommunication site(?:s)?\b/i
];

const INTERNAL_CONTENT_HINTS: ReadonlyArray<RegExp> = [
  /\bsharepoint\b/i,
  /\bonedrive\b/i,
  /\bmicrosoft 365\b/i,
  /\bm365\b/i,
  /\bdocument(?:s)?\b/i,
  /\bfile(?:s)?\b/i,
  /\bpdf(?:s)?\b/i,
  /\bpptx?\b/i,
  /\bxlsx?\b/i,
  /\bdocx?\b/i,
  /\bspreadsheet(?:s)?\b/i,
  /\bpresentation(?:s)?\b/i,
  /\breport(?:s)?\b/i,
  /\blibrary\b/i,
  /\bfolder\b/i
];

const CONTAINER_BROWSE_HINTS: ReadonlyArray<RegExp> = [
  /\bdocument library\b/i,
  /\blibrary\b/i,
  /\bshow (?:me )?(?:all )?(?:the )?files\b/i,
  /\bshow (?:me )?(?:the )?content\b/i,
  /\blist (?:all )?(?:the )?files\b/i,
  /\bbrowse\b/i,
  /\bfolder\b/i
];

const CONTAINER_SCOPE_HINTS: ReadonlyArray<RegExp> = [
  /\bsite\b/i,
  /\blibrary\b/i,
  /\bfolder\b/i,
  /\blist\b/i,
  /https?:\/\/[^\s]*sharepoint\.com/i
];

const CLARIFICATION_PATTERNS: ReadonlyArray<RegExp> = [
  /\bdo you want me to\b/i,
  /\bwould you like me to\b/i,
  /\bcan you clarify\b/i,
  /\bcould you clarify\b/i,
  /\bare you looking for\b/i,
  /\bdo you mean\b/i,
  /\bwhich (?:one|type|kind|source|option|domain)\b/i,
  /\bor search something else\b/i
];

const NON_DECISION_TOOLS: ReadonlySet<string> = new Set([
  'set_expression',
  'show_progress',
  'clear_action_panel'
]);

export const ENTERPRISE_INTENT_REGRESSION_FIXTURES: ReadonlyArray<IIntentRoutingRegressionFixture> = [
  {
    utterance: 'check this url for me https://en.wikipedia.org/wiki/Microsoft',
    expectedRoute: 'research_public_web',
    expectedToolName: 'research_public_web',
    expectsClarification: false
  },
  {
    utterance: 'what do you offer?',
    expectedRoute: 'list_m365_servers',
    expectedToolName: 'list_m365_servers',
    expectsClarification: false
  },
  {
    utterance: 'was kannst du alles?',
    expectedRoute: 'list_m365_servers',
    expectedToolName: 'list_m365_servers',
    expectsClarification: false
  },
  {
    utterance: 'que peux-tu faire ?',
    expectedRoute: 'list_m365_servers',
    expectedToolName: 'list_m365_servers',
    expectsClarification: false
  },
  {
    utterance: 'cosa puoi fare?',
    expectedRoute: 'list_m365_servers',
    expectedToolName: 'list_m365_servers',
    expectsClarification: false
  },
  {
    utterance: 'what can you do for SharePoint?',
    expectedRoute: 'list_m365_servers',
    expectedToolName: 'list_m365_servers',
    expectedToolArgs: { focus: 'sharepoint' },
    expectsClarification: false
  },
  {
    utterance: 'what can you do in Teams?',
    expectedRoute: 'list_m365_servers',
    expectedToolName: 'list_m365_servers',
    expectedToolArgs: { focus: 'teams' },
    expectsClarification: false
  },
  {
    utterance: 'i am searching for documents about animals',
    expectedRoute: 'search_sharepoint',
    expectedToolName: 'search_sharepoint',
    expectsClarification: false
  },
  {
    utterance: 'find emails about animals',
    expectedRoute: 'search_emails',
    expectedToolName: 'search_emails',
    expectsClarification: false
  },
  {
    utterance: 'find people working on animals',
    expectedRoute: 'search_people',
    expectedToolName: 'search_people',
    expectsClarification: false
  },
  {
    utterance: 'find sites about animals',
    expectedRoute: 'search_sites',
    expectedToolName: 'search_sites',
    expectsClarification: false
  }
];

function matchesAny(text: string, patterns: ReadonlyArray<RegExp>): boolean {
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(text)) return true;
  }
  return false;
}

export function hasActionableSearchIntent(text: string): boolean {
  return matchesAny(normalizeText(text), ACTIONABLE_SEARCH_PATTERNS);
}

export function hasExplicitExternalWebHint(text: string): boolean {
  return matchesAny(normalizeText(text), EXTERNAL_WEB_HINTS);
}

export function hasInternalEnterpriseContentHint(text: string): boolean {
  return matchesAny(normalizeText(text), INTERNAL_CONTENT_HINTS);
}

function isComposeOrShareRequest(text: string): boolean {
  if (matchesAny(text, [/reply\b/i, /\bforward\b/i])) {
    return true;
  }

  if (!matchesAny(text, COMPOSE_ACTION_HINTS)) {
    return false;
  }

  if (matchesAny(text, EMAIL_HINTS)) {
    return true;
  }

  return matchesAny(text, CONTEXT_SHARE_TARGET_HINTS) && matchesAny(text, [/by mail\b/i, /\bemail\b/i, /\bmail\b/i]);
}

function isExplicitContainerBrowseRequest(text: string): boolean {
  if (!matchesAny(text, ACTIONABLE_SEARCH_PATTERNS)) {
    return false;
  }

  if (!matchesAny(text, CONTAINER_BROWSE_HINTS)) {
    return false;
  }

  if (!matchesAny(text, CONTAINER_SCOPE_HINTS)) {
    return false;
  }

  return !matchesAny(text, EXTERNAL_WEB_HINTS)
    && !matchesAny(text, EMAIL_HINTS)
    && !matchesAny(text, PEOPLE_HINTS);
}

function normalizeText(text: string): string {
  return text.trim().replace(/\s+/g, ' ');
}

function normalizeJson(raw: string): string {
  const trimmed = raw.trim();
  if (trimmed.startsWith('```')) {
    return trimmed.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  }
  return trimmed;
}

function createObservation(
  userText: string,
  normalizedText: string,
  route: EnterpriseRouteTarget,
  source: FirstTurnRoutingSource,
  confidence: number,
  isGenericEnterpriseSearch: boolean,
  options?: {
    expectedToolArgs?: Record<string, unknown>;
    capabilityFocus?: McpCapabilityFocus;
  }
): IFirstTurnRoutingObservation {
  return {
    userText,
    normalizedText,
    expectedRoute: route,
    expectedToolName: route === 'none' ? undefined : route,
    expectedToolArgs: options?.expectedToolArgs,
    capabilityFocus: options?.capabilityFocus,
    isGenericEnterpriseSearch,
    source,
    confidence
  };
}

function parseFastRoutingResponse(raw: string): { route: EnterpriseRouteTarget; confidence: number } | undefined {
  try {
    const parsed = JSON.parse(normalizeJson(raw)) as {
      route?: string;
      confidence?: number | string;
    };
    const route = typeof parsed.route === 'string' ? parsed.route.trim() as EnterpriseRouteTarget : undefined;
    const confidenceValue = typeof parsed.confidence === 'string'
      ? Number(parsed.confidence)
      : parsed.confidence;
    if (!route || !VALID_FAST_ROUTES.has(route)) {
      return undefined;
    }
    if (typeof confidenceValue !== 'number' || !Number.isFinite(confidenceValue)) {
      return undefined;
    }
    return {
      route,
      confidence: Math.max(0, Math.min(1, confidenceValue))
    };
  } catch {
    return undefined;
  }
}

export function buildEnterpriseIntentRoutingGuidance(): string {
  return [
    '- Treat first-turn capability routing as enterprise-first and action-oriented. Choose the best plausible non-destructive capability family immediately instead of asking the user to pick a domain.',
    '- If the user wants to send, compose, draft, reply to, forward, or share something by email, do not force a search route. Let normal tool planning use `show_compose_form` or other write actions instead.',
    '- If the user asks what you can do, what you offer, or requests a capability overview in any language, call `list_m365_servers` so the action panel shows the overview instead of answering only in chat.',
    '- If the user asks what you can do for a specific workload such as SharePoint, OneDrive, Outlook, Teams, Calendar, Word, Profile, or Copilot, still call `list_m365_servers` but pass `focus` with that workload so the panel shows tool-level detail.',
    '- If the user explicitly asks for external, web, public, internet, website, GitHub, Wikipedia, or URL research, use `research_public_web`. Pass the user request in `query` and include `target_url` when a specific page or URL is provided.',
    '- If the user explicitly asks for emails, inbox messages, Outlook mail, senders, or subjects, use `search_emails`.',
    '- If the user explicitly asks for people, employees, colleagues, contacts, or "who is", use `search_people`.',
    '- If the user explicitly asks for sites, team sites, communication sites, or SharePoint sites, use `search_sites`.',
    '- If the user explicitly asks to browse or list the contents of a specific SharePoint document library, folder, or list in a named site, do not force `search_sharepoint`. Let the direct browse tool planning handle it.',
    '- If the user explicitly asks to browse their personal OneDrive or search it by file name or prefix, do not force `search_sharepoint`. Let the dedicated OneDrive MCP override handle it.',
    '- If the user explicitly asks for documents, files, PDFs, reports, presentations, spreadsheets, SharePoint, or OneDrive content, use `search_sharepoint`.',
    '- If the user uses generic enterprise search phrasing such as "search", "look up", "find", or "information about" without naming a workload or external source, default to `search_sharepoint` even when the topic contains department or subject words such as "marketing".',
    '- Clarify only when no plausible non-destructive capability family fits, or when multiple competing non-search workflows are equally likely.'
  ].join('\n');
}

export function heuristicFirstTurnRouting(userText: string): IFirstTurnRoutingObservation | undefined {
  const normalizedText = normalizeText(userText);
  if (!normalizedText || isHiePromptMessage(normalizedText)) {
    return undefined;
  }
  if (classifyExplicitPersonalOneDriveIntent(normalizedText)) {
    return undefined;
  }
  if (isComposeOrShareRequest(normalizedText)) {
    return undefined;
  }
  if (isExplicitContainerBrowseRequest(normalizedText)) {
    return undefined;
  }
  if (matchesAny(normalizedText, CAPABILITY_DISCOVERY_PATTERNS)) {
    const capabilityFocus = detectCapabilityFocus(normalizedText);
    return createObservation(userText, normalizedText, 'list_m365_servers', 'heuristic', 1, false, capabilityFocus
      ? {
          capabilityFocus,
          expectedToolArgs: { focus: capabilityFocus }
        }
      : undefined);
  }
  if (!matchesAny(normalizedText, ACTIONABLE_SEARCH_PATTERNS)) {
    return undefined;
  }

  if (matchesAny(normalizedText, EXTERNAL_WEB_HINTS)) {
    return createObservation(userText, normalizedText, 'research_public_web', 'heuristic', 1, false);
  }

  if (matchesAny(normalizedText, EMAIL_HINTS)) {
    return createObservation(userText, normalizedText, 'search_emails', 'heuristic', 1, false);
  }

  if (matchesAny(normalizedText, PEOPLE_HINTS)) {
    return createObservation(userText, normalizedText, 'search_people', 'heuristic', 1, false);
  }

  if (matchesAny(normalizedText, SITE_HINTS)) {
    return createObservation(userText, normalizedText, 'search_sites', 'heuristic', 1, false);
  }

  if (matchesAny(normalizedText, INTERNAL_CONTENT_HINTS)) {
    return createObservation(userText, normalizedText, 'search_sharepoint', 'heuristic', 1, false);
  }

  return undefined;
}

export async function classifyFirstTurnRouting(
  userText: string,
  proxyConfig?: IProxyConfig
): Promise<IFirstTurnRoutingObservation | undefined> {
  const normalizedText = normalizeText(userText);
  if (!normalizedText || isHiePromptMessage(normalizedText)) {
    return undefined;
  }
  if (classifyExplicitPersonalOneDriveIntent(normalizedText)) {
    return undefined;
  }
  if (isComposeOrShareRequest(normalizedText)) {
    return undefined;
  }

  const heuristicObservation = heuristicFirstTurnRouting(userText);
  const nano = getNanoService(proxyConfig);
  if (!nano) {
    return heuristicObservation;
  }

  const tuning = getRuntimeTuningConfig().nano;
  const response = await nano.classify(
    FAST_INTENT_ROUTER_SYSTEM_PROMPT,
    normalizedText,
    tuning.intentRoutingTimeoutMs,
    tuning.intentRoutingMaxTokens
  );
  if (!response) {
    return heuristicObservation;
  }

  const parsed = parseFastRoutingResponse(response);
  if (!parsed || parsed.confidence < tuning.intentRoutingConfidenceThreshold) {
    return heuristicObservation;
  }

  return createObservation(
    userText,
    normalizedText,
    parsed.route,
    'fast',
    parsed.confidence,
    heuristicObservation?.isGenericEnterpriseSearch === true && parsed.route === 'search_sharepoint',
    heuristicObservation?.expectedRoute === parsed.route
      ? {
          capabilityFocus: heuristicObservation.capabilityFocus,
          expectedToolArgs: heuristicObservation.expectedToolArgs
        }
      : undefined
  );
}

export function getFirstTurnRoutingObservation(userText: string): IFirstTurnRoutingObservation | undefined {
  return heuristicFirstTurnRouting(userText);
}

export function getForcedFirstToolName(observation?: IFirstTurnRoutingObservation): string | undefined {
  if (!observation) return undefined;
  return observation.expectedToolName;
}

export function getForcedFirstToolArgs(
  observation?: IFirstTurnRoutingObservation
): Record<string, unknown> | undefined {
  if (!observation?.expectedToolArgs) {
    return undefined;
  }
  return { ...observation.expectedToolArgs };
}

export function classifyAssistantFirstTurnOutcome(text: string): FirstTurnRoutingOutcome {
  return isClarificationResponse(text) ? 'clarification' : 'answer_only';
}

export function isClarificationResponse(text: string): boolean {
  const normalizedText = normalizeText(text);
  if (!normalizedText) return false;
  return matchesAny(normalizedText, CLARIFICATION_PATTERNS);
}

export function getObservedFirstToolName(toolNames: ReadonlyArray<string>): string | undefined {
  for (let i = 0; i < toolNames.length; i++) {
    if (!NON_DECISION_TOOLS.has(toolNames[i])) {
      return toolNames[i];
    }
  }
  return undefined;
}

export function buildFirstTurnRoutingLogDetail(
  channel: FirstTurnRoutingChannel,
  outcome: FirstTurnRoutingOutcome,
  observation: IFirstTurnRoutingObservation,
  actualToolName?: string
): string {
  return JSON.stringify({
    channel,
    outcome,
    userText: observation.normalizedText,
    expectedRoute: observation.expectedRoute,
    expectedToolName: observation.expectedToolName,
    expectedToolArgs: observation.expectedToolArgs,
    capabilityFocus: observation.capabilityFocus,
    source: observation.source,
    confidence: observation.confidence,
    genericEnterpriseSearch: observation.isGenericEnterpriseSearch,
    actualToolName
  });
}
