/**
 * PeopleService
 * People search and profile operations via Microsoft Graph.
 */

import { AadHttpClient } from '@microsoft/sp-http';
import { GraphService } from './GraphService';
import type { IGraphResponse } from './GraphService';

// ─── Graph API response shapes ──────────────────────────────────

interface IGraphPerson {
  id?: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  emailAddresses?: Array<{ address: string }>;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  phones?: Array<{ type: string; number: string }>;
  scoredEmailAddresses?: Array<{ address: string; relevanceScore: number }>;
}

interface IGraphUser {
  id?: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  userType?: string;
}

interface IGraphCollection<T> {
  value: T[];
}

// ─── Public interfaces ──────────────────────────────────────────

export interface IPersonResult {
  displayName: string;
  email: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  phone?: string;
  relevanceScore?: number;
}

type PersonResultSource = 'directory' | 'people';

interface IScoredPersonResult extends IPersonResult {
  id?: string;
  givenName?: string;
  surname?: string;
  source: PersonResultSource;
  userType?: string;
  score?: number;
}

const PEOPLE_SELECT = 'displayName,givenName,surname,mail,jobTitle,department,officeLocation,phones,scoredEmailAddresses,userPrincipalName';
const USER_SELECT = 'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone,userType';

const LEADING_QUERY_PATTERNS: ReadonlyArray<RegExp> = [
  /^(?:please\s+)?show(?:\s+me)?\s+/i,
  /^(?:please\s+)?find\s+/i,
  /^(?:please\s+)?search(?:\s+for)?\s+/i,
  /^(?:please\s+)?look(?:\s+up)?\s+/i,
  /^(?:please\s+)?lookup\s+/i,
  /^(?:please\s+)?open\s+/i,
  /^(?:please\s+)?display\s+/i,
  /^(?:please\s+)?get\s+/i,
  /^(?:please\s+)?give\s+me\s+/i,
  /^(?:please\s+)?fetch\s+/i,
  /^(?:please\s+)?who\s+is\s+/i
];

const TRAILING_QUERY_PATTERNS: ReadonlyArray<RegExp> = [
  /\s+(?:people|person|user|contact|employee|colleague)\s+card$/i,
  /\s+profile\s+card$/i,
  /\s+(?:people|person|user|contact|employee|colleague|profile)$/i,
  /\s+card$/i
];

function collapseWhitespace(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

function normalizeComparable(value: string | undefined): string {
  return collapseWhitespace((value || '').toLowerCase().replace(/[^a-z0-9@._\s-]/g, ' '));
}

function tokenizeComparable(value: string | undefined): string[] {
  return normalizeComparable(value)
    .split(' ')
    .map((token) => token.trim())
    .filter(Boolean);
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function dedupeCandidates(candidates: IScoredPersonResult[]): IScoredPersonResult[] {
  const bestByKey = new Map<string, IScoredPersonResult>();

  candidates.forEach((candidate) => {
    const key = normalizeComparable(candidate.email) || normalizeComparable(candidate.displayName);
    if (!key) return;

    const previous = bestByKey.get(key);
    if (!previous || (candidate.score || 0) > (previous.score || 0)) {
      bestByKey.set(key, candidate);
    }
  });

  return Array.from(bestByKey.values());
}

function mapGraphPerson(person: IGraphPerson): IScoredPersonResult {
  const scoredAddress = person.scoredEmailAddresses && person.scoredEmailAddresses.length > 0
    ? person.scoredEmailAddresses[0]
    : undefined;
  const fallbackAddress = person.emailAddresses && person.emailAddresses.length > 0
    ? person.emailAddresses[0].address
    : undefined;
  const phone = person.phones && person.phones.length > 0 ? person.phones[0].number : undefined;

  return {
    id: person.id,
    displayName: person.displayName || '',
    email: scoredAddress?.address || person.mail || fallbackAddress || person.userPrincipalName || '',
    givenName: person.givenName,
    surname: person.surname,
    jobTitle: person.jobTitle,
    department: person.department,
    officeLocation: person.officeLocation,
    phone,
    relevanceScore: scoredAddress?.relevanceScore,
    source: 'people'
  };
}

function mapGraphUser(user: IGraphUser): IScoredPersonResult {
  const businessPhone = Array.isArray(user.businessPhones)
    ? user.businessPhones.find((phone) => typeof phone === 'string' && phone.trim())
    : undefined;

  return {
    id: user.id,
    displayName: user.displayName || '',
    email: user.mail || user.userPrincipalName || '',
    givenName: user.givenName,
    surname: user.surname,
    jobTitle: user.jobTitle,
    department: user.department,
    officeLocation: user.officeLocation,
    phone: businessPhone || user.mobilePhone,
    source: 'directory',
    userType: user.userType
  };
}

function buildNameParts(query: string): { first: string; rest?: string } | undefined {
  const parts = collapseWhitespace(query).split(' ').filter(Boolean);
  if (parts.length === 0) return undefined;
  return {
    first: parts[0],
    rest: parts.length > 1 ? parts.slice(1).join(' ') : undefined
  };
}

function scoreCandidate(candidate: IScoredPersonResult, query: string): number {
  const normalizedQuery = normalizeComparable(query);
  const queryTokens = tokenizeComparable(query);
  const displayName = normalizeComparable(candidate.displayName);
  const email = normalizeComparable(candidate.email);
  const givenName = normalizeComparable(candidate.givenName);
  const surname = normalizeComparable(candidate.surname);
  const fullName = collapseWhitespace(`${givenName} ${surname}`);
  const combined = [displayName, fullName, email].filter(Boolean).join(' ');
  const nameParts = buildNameParts(query);
  let score = candidate.source === 'directory' ? 220 : 60;

  if (candidate.userType?.toLowerCase() === 'member') {
    score += 40;
  } else if (candidate.userType?.toLowerCase() === 'guest') {
    score -= 40;
  }

  if (email && email === normalizedQuery) {
    score += 400;
  }
  if (displayName && displayName === normalizedQuery) {
    score += 320;
  }
  if (fullName && fullName === normalizedQuery) {
    score += 320;
  }
  if (normalizedQuery && displayName.startsWith(normalizedQuery)) {
    score += 150;
  }
  if (normalizedQuery && fullName.startsWith(normalizedQuery)) {
    score += 150;
  }
  if (normalizedQuery && email.startsWith(normalizedQuery)) {
    score += 120;
  }

  if (queryTokens.length > 0) {
    const matchedTokenCount = queryTokens.filter((token) => combined.includes(token)).length;
    if (matchedTokenCount === queryTokens.length) {
      score += 120;
    }
    score += matchedTokenCount * 22;
  }

  if (nameParts?.first) {
    const firstName = normalizeComparable(nameParts.first);
    const restName = normalizeComparable(nameParts.rest);

    if (firstName && givenName === firstName) {
      score += 90;
    } else if (firstName && givenName.startsWith(firstName)) {
      score += 55;
    }

    if (restName && surname === restName) {
      score += 140;
    } else if (restName && surname.startsWith(restName)) {
      score += 95;
    }
  }

  if (!candidate.email) {
    score -= 25;
  }

  return score;
}

function rankCandidates(candidates: IScoredPersonResult[], query: string, top: number): IPersonResult[] {
  const scored = dedupeCandidates(candidates.map((candidate) => ({
    ...candidate,
    score: scoreCandidate(candidate, query)
  }))).sort((left, right) => {
    const scoreDelta = (right.score || 0) - (left.score || 0);
    if (scoreDelta !== 0) return scoreDelta;

    const sourceDelta = (left.source === 'directory' ? -1 : 1) - (right.source === 'directory' ? -1 : 1);
    if (sourceDelta !== 0) return sourceDelta;

    return left.displayName.localeCompare(right.displayName);
  });

  const strongDirectoryMatch = scored.some((candidate) => candidate.source === 'directory' && (candidate.score || 0) >= 320);
  const filtered = strongDirectoryMatch
    ? scored.filter((candidate) => candidate.source === 'directory')
    : scored;

  return filtered.slice(0, top).map((candidate) => ({
    displayName: candidate.displayName,
    email: candidate.email,
    jobTitle: candidate.jobTitle,
    department: candidate.department,
    officeLocation: candidate.officeLocation,
    phone: candidate.phone,
    relevanceScore: candidate.relevanceScore
  }));
}

export function normalizePeopleSearchQuery(query: string): string {
  let normalized = collapseWhitespace(query);
  if (!normalized) return '';

  normalized = normalized.replace(/^["']+|["']+$/g, '');
  let previous = '';
  while (normalized !== previous) {
    previous = normalized;
    LEADING_QUERY_PATTERNS.forEach((pattern) => {
      normalized = normalized.replace(pattern, '');
    });
    TRAILING_QUERY_PATTERNS.forEach((pattern) => {
      normalized = normalized.replace(pattern, '');
    });
    normalized = collapseWhitespace(normalized);
  }

  return normalized || collapseWhitespace(query);
}

// ─── Service ───────────────────────────────────────────────────

export class PeopleService {
  private graphService: GraphService;

  constructor(client: AadHttpClient) {
    this.graphService = new GraphService(client);
  }

  /**
   * Search for people using the People API (relevance-ranked).
   */
  public async searchPeople(query: string, top: number = 10): Promise<IGraphResponse<IPersonResult[]>> {
    const normalizedQuery = normalizePeopleSearchQuery(query);
    const boundedTop = Math.min(Math.max(top, 1), 25);
    if (!normalizedQuery) {
      return { success: true, data: [], durationMs: 0 };
    }

    const startTime = performance.now();
    const directoryCandidates = await this.searchDirectoryUsers(normalizedQuery, boundedTop);
    const peopleCandidates = directoryCandidates.length >= boundedTop
      ? []
      : await this.searchPeopleApi(normalizedQuery, boundedTop);
    const ranked = rankCandidates([
      ...directoryCandidates,
      ...peopleCandidates
    ], normalizedQuery, boundedTop);

    return {
      success: true,
      data: ranked,
      durationMs: Math.round(performance.now() - startTime)
    };
  }

  private async searchPeopleApi(query: string, top: number): Promise<IScoredPersonResult[]> {
    const resp = await this.graphService.get<IGraphCollection<IGraphPerson>>(
      `/me/people?$search="${encodeURIComponent(query)}"&$top=${top}&$select=${PEOPLE_SELECT}`
    );

    if (!resp.success || !resp.data) {
      return [];
    }

    return resp.data.value.map((person) => mapGraphPerson(person));
  }

  private async searchDirectoryUsers(query: string, top: number): Promise<IScoredPersonResult[]> {
    const filters = this.buildDirectoryFilters(query);
    const candidates: IScoredPersonResult[] = [];

    for (let index = 0; index < filters.length; index++) {
      const resp = await this.graphService.get<IGraphCollection<IGraphUser>>(
        `/users?$filter=${encodeURIComponent(filters[index])}&$top=${top}&$select=${USER_SELECT}`
      );
      if (!resp.success || !resp.data) {
        continue;
      }

      resp.data.value.forEach((user) => {
        if (!user.displayName && !user.mail && !user.userPrincipalName) return;
        candidates.push(mapGraphUser(user));
      });
    }

    return candidates;
  }

  private buildDirectoryFilters(query: string): string[] {
    const filters = new Set<string>();
    const safeQuery = escapeODataString(query);
    const nameParts = buildNameParts(query);

    if (query.includes('@')) {
      filters.add(`mail eq '${safeQuery}' or userPrincipalName eq '${safeQuery}'`);
      return Array.from(filters);
    }

    filters.add(`displayName eq '${safeQuery}'`);
    filters.add(`startswith(displayName,'${safeQuery}')`);

    if (nameParts?.first) {
      const first = escapeODataString(nameParts.first);
      filters.add(`startswith(givenName,'${first}')`);

      if (nameParts.rest) {
        const rest = escapeODataString(nameParts.rest);
        filters.add(`givenName eq '${first}' and surname eq '${rest}'`);
        filters.add(`startswith(givenName,'${first}') and startswith(surname,'${rest}')`);
      } else {
        filters.add(`startswith(surname,'${first}')`);
      }
    }

    return Array.from(filters);
  }
}
