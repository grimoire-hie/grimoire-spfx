/**
 * ContextService — Builds user context from SPFx pageContext + optional Graph /me profile.
 * Injected into SystemPrompt so the LLM always knows who the user is and where they are.
 */

import { GraphService } from '../graph/GraphService';
import { logService } from '../logging/LogService';
import { normalizeError } from '../utils/errorUtils';
import type { AadHttpClient } from '@microsoft/sp-http';

// ─── Types ──────────────────────────────────────────────────────

export interface IUserContext {
  displayName: string;
  email: string;
  loginName: string;
  jobTitle?: string;
  department?: string;
  manager?: string;
  preferredLanguage?: string;
  resolvedLanguage: string;
  currentWebTitle: string;
  currentWebUrl: string;
  currentSiteTitle: string;
  currentSiteUrl: string;
  currentListTitle?: string;
  currentItemId?: number;
}

// ─── Language Resolution ────────────────────────────────────────

const REALTIME_LANGUAGES = new Set([
  'en', 'it', 'af', 'es', 'de', 'fr', 'id', 'ru', 'pl', 'uk',
  'el', 'lv', 'zh', 'ar', 'tr', 'ja', 'sw', 'cy', 'ko', 'is',
  'bn', 'ur', 'ne', 'th', 'pa', 'mr', 'te'
]);

function resolveLanguage(preferredLanguage?: string): string {
  if (!preferredLanguage) return 'en';
  const primary = preferredLanguage.split('-')[0].toLowerCase();
  return REALTIME_LANGUAGES.has(primary) ? primary : 'en';
}

// ─── Graph /me Profile ──────────────────────────────────────────

interface IGraphMeProfile {
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  preferredLanguage?: string;
  manager?: { displayName?: string };
}

/**
 * Fetch enriched profile from Graph /me. Returns partial data — never throws.
 */
async function fetchGraphProfile(aadHttpClient: AadHttpClient): Promise<IGraphMeProfile | undefined> {
  try {
    const graphService = new GraphService(aadHttpClient);
    const result = await graphService.get<IGraphMeProfile>(
      '/me?$select=jobTitle,department,officeLocation,preferredLanguage&$expand=manager($select=displayName)'
    );
    if (result.success && result.data) {
      return result.data;
    }
    logService.debug('graph', `Graph /me enrichment failed: ${result.error || 'no data'}`);
    return undefined;
  } catch (err) {
    const normalizedError = normalizeError(err, 'Graph /me enrichment failed');
    logService.debug('graph', `Graph /me enrichment error: ${normalizedError.message}`);
    return undefined;
  }
}

// ─── PageContext Shape ───────────────────────────────────────────

/**
 * Minimal shape of SPFx pageContext fields we use.
 * Using an interface avoids importing full SPFx types.
 */
export interface IPageContextLike {
  user?: {
    displayName?: string;
    email?: string;
    loginName?: string;
  };
  web?: {
    title?: string;
    absoluteUrl?: string;
  };
  site?: {
    absoluteUrl?: string;
  };
  list?: {
    title?: string;
  };
  listItem?: {
    id?: number;
  };
}

// ─── Build User Context ─────────────────────────────────────────

export function buildBaseUserContext(pageContext: IPageContextLike): IUserContext {
  const user = pageContext.user;
  const web = pageContext.web;
  const site = pageContext.site;

  return {
    displayName: user?.displayName || 'User',
    email: user?.email || '',
    loginName: user?.loginName || '',
    resolvedLanguage: 'en',
    currentWebTitle: web?.title || '',
    currentWebUrl: web?.absoluteUrl || '',
    currentSiteTitle: web?.title || '',
    currentSiteUrl: site?.absoluteUrl || '',
    currentListTitle: pageContext.list?.title,
    currentItemId: pageContext.listItem?.id
  };
}

export async function enrichUserContextWithGraph(
  context: IUserContext,
  aadHttpClient?: AadHttpClient
): Promise<IUserContext> {
  if (!aadHttpClient) {
    return context;
  }

  const nextContext: IUserContext = { ...context };
  const profile = await fetchGraphProfile(aadHttpClient);
  if (profile) {
    nextContext.jobTitle = profile.jobTitle || undefined;
    nextContext.department = profile.department || undefined;
    nextContext.manager = profile.manager?.displayName || undefined;
    nextContext.preferredLanguage = profile.preferredLanguage || undefined;
    nextContext.resolvedLanguage = resolveLanguage(profile.preferredLanguage);
    logService.info('graph', `User context enriched: ${nextContext.displayName}, lang=${nextContext.resolvedLanguage}`);
  }

  return nextContext;
}

/**
 * Build IUserContext from SPFx pageContext + optional Graph enrichment.
 * Graph call is best-effort — returns base context immediately if it fails.
 */
export async function buildUserContext(
  pageContext: IPageContextLike,
  aadHttpClient?: AadHttpClient
): Promise<IUserContext> {
  const baseContext = buildBaseUserContext(pageContext);
  return enrichUserContextWithGraph(baseContext, aadHttpClient);
}
