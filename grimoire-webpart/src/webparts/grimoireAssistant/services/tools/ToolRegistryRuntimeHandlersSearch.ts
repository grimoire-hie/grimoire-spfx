import { useGrimoireStore } from '../../store/useGrimoireStore';
import type {
  IErrorData,
  IInfoCardData,
  IMarkdownData,
  ISearchQueryVariantInfo,
  ISearchResult,
  ISearchResultsData,
  ISiteInfoData,
  ISelectionListData,
  IUserCardData
} from '../../models/IBlock';
import type { IMcpConnection } from '../../models/IMcpTypes';
import { createBlock } from '../../models/IBlock';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { logService } from '../logging/LogService';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { CopilotSearchService } from '../search/CopilotSearchService';
import { CopilotRetrievalService } from '../search/CopilotRetrievalService';
import { RRFFusionService } from '../search/RRFFusionService';
import { SearchIntentPlanner } from '../search/SearchIntentPlanner';
import { SharePointSearchService } from '../search/SharePointSearchService';
import { detectQueryLanguage } from '../search/SearchLanguageUtils';
import { resolveServerUrl } from '../../models/McpServerCatalog';
import { McpClientService } from '../mcp/McpClientService';
import { executeCatalogMcpTool } from '../mcp/McpExecutionAdapter';
import { mapMcpResultToBlock } from '../mcp/McpResultMapper';
import type { ICopilotSearchResult, SearchQueryVariantKind } from '../../models/ISearchTypes';
import { PublicWebSearchService, type IPublicWebReference } from '../web/PublicWebSearchService';
import type { RuntimeHandledToolName } from './ToolRuntimeHandlerRegistry';
import type { ToolRuntimeHandler, ToolRuntimeHandlerResult } from './ToolRuntimeHandlerTypes';
import type { SearchRuntimeToolName } from './ToolRuntimeHandlerPartitions';
import { parseRenderHints } from './ToolRuntimeContentHelpers';
import { trackCreatedBlock, trackToolCompletion, trackUpdatedBlock } from './ToolRuntimeHieHelpers';
import { completeOutcome, errorOutcome } from './ToolRuntimeOutcomeHelpers';
import type { IMcpContent } from '../../models/IMcpTypes';
import { normalizePeopleSearchQuery, type IPersonResult } from '../graph/PeopleService';

interface ISearchRuntimeHelpers {
  findExistingSession: (connections: IMcpConnection[], serverUrl: string) => string | undefined;
  connectToM365Server: (
    mcpClient: McpClientService,
    serverUrl: string,
    serverName: string,
    getToken: ((resource: string) => Promise<string>) | undefined
  ) => Promise<string>;
}

const PERSON_PROFILE_SELECT = 'displayName,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone,givenName,surname,id';

function asRecord(value: unknown): Record<string, unknown> | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
  return value as Record<string, unknown>;
}

function firstNonEmptyString(obj: Record<string, unknown>, keys: string[]): string | undefined {
  for (let i = 0; i < keys.length; i++) {
    const value = obj[keys[i]];
    if (typeof value === 'string' && value.trim()) {
      return value.trim();
    }
  }
  return undefined;
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

function parseJsonCandidate(text: string): unknown | undefined {
  const trimmed = stripCopilotMetadata(text.trim());
  if (!trimmed) return undefined;
  try {
    return JSON.parse(trimmed);
  } catch {
    for (let i = 0; i < trimmed.length; i++) {
      const ch = trimmed[i];
      if (ch !== '{' && ch !== '[') continue;
      const jsonSlice = extractBalancedJson(trimmed, i);
      if (!jsonSlice) continue;
      try {
        return JSON.parse(jsonSlice);
      } catch {
        continue;
      }
    }
    return undefined;
  }
}

function extractProfilePayload(content: IMcpContent[]): Record<string, unknown> | undefined {
  const textParts = content
    .filter((item) => item.type === 'text' && item.text)
    .map((item) => item.text || '');

  for (let i = 0; i < textParts.length; i++) {
    const parsed = parseJsonCandidate(textParts[i]);
    const root = asRecord(parsed);
    if (!root) continue;
    const nested = asRecord(root.data)
      || asRecord(root.response)
      || asRecord(root.result)
      || asRecord(root.payload);
    if (nested) return nested;
    return root;
  }

  const merged = textParts.join('\n');
  const parsed = parseJsonCandidate(merged);
  const root = asRecord(parsed);
  if (!root) return undefined;
  return asRecord(root.data)
    || asRecord(root.response)
    || asRecord(root.result)
    || asRecord(root.payload)
    || root;
}

function extractProfilePhone(payload: Record<string, unknown>): string | undefined {
  const businessPhones = payload.businessPhones;
  if (Array.isArray(businessPhones)) {
    const first = businessPhones.find((value) => typeof value === 'string' && value.trim());
    if (typeof first === 'string') {
      return first.trim();
    }
  }

  return firstNonEmptyString(payload, ['mobilePhone', 'phone']);
}

function mergePersonProfile(candidate: IPersonResult, payload: Record<string, unknown>): IPersonResult {
  return {
    displayName: firstNonEmptyString(payload, ['displayName', 'givenName']) || candidate.displayName,
    email: firstNonEmptyString(payload, ['mail', 'userPrincipalName']) || candidate.email,
    jobTitle: firstNonEmptyString(payload, ['jobTitle']) || candidate.jobTitle,
    department: firstNonEmptyString(payload, ['department']) || candidate.department,
    officeLocation: firstNonEmptyString(payload, ['officeLocation']) || candidate.officeLocation,
    phone: extractProfilePhone(payload) || candidate.phone,
    relevanceScore: candidate.relevanceScore
  };
}

function shouldEnrichPeopleSearch(query: string, people: IPersonResult[]): boolean {
  const trimmed = query.trim();
  if (!trimmed) return false;
  if (people.length === 0 || people.length > 3) return false;
  return true;
}

function isGenericPeopleQuery(query: string): boolean {
  const trimmed = query.trim().toLowerCase();
  if (!trimmed) return true;
  if (trimmed === '*' || trimmed === '"*"') return true;

  const normalized = trimmed
    .replace(/[^a-z0-9\s]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return normalized === 'people'
    || normalized === 'person'
    || normalized === 'a person'
    || normalized === 'any person'
    || normalized === 'any people'
    || normalized === 'someone'
    || normalized === 'somebody'
    || normalized === 'anyone'
    || normalized === 'anybody';
}

function safeParseMaxResults(rawValue: string | undefined, fallback: number): number {
  const parsed = Number.parseInt(rawValue || '', 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(Math.max(parsed, 1), 25);
}

function decorateResults(
  results: ICopilotSearchResult[],
  queryText: string,
  variantKind: SearchQueryVariantKind,
  variantLanguage: string
): ICopilotSearchResult[] {
  return results.map((result, index) => ({
    ...result,
    sourceRank: result.sourceRank || (index + 1),
    queryText,
    variantKind,
    variantLanguage
  }));
}

function buildVariantInfo(
  variants: Array<{ kind: SearchQueryVariantKind; query: string; language?: string }>
): ISearchQueryVariantInfo[] | undefined {
  if (variants.length === 0) return undefined;
  const seen = new Set<string>();
  const deduped: ISearchQueryVariantInfo[] = [];

  variants.forEach((variant) => {
    const key = `${variant.kind}:${variant.language || ''}:${variant.query.trim().toLowerCase()}`;
    if (seen.has(key)) return;
    seen.add(key);
    deduped.push({
      kind: variant.kind,
      query: variant.query,
      language: variant.language
    });
  });

  return deduped.length > 0 ? deduped : undefined;
}

function mapFusedResults(results: ReturnType<RRFFusionService['fuseWithContext']>, maxResults: number): ISearchResult[] {
  return results.slice(0, maxResults).map((result) => ({
    title: result.title,
    summary: result.summary,
    url: result.url,
    fileType: result.fileType,
    lastModified: result.lastModified,
    author: result.author,
    siteName: result.siteName,
    sources: result.sources,
    language: result.language
  }));
}

function collectSourceList(results: ISearchResult[]): string[] {
  const allSources = new Set<string>();
  results.forEach((result) => (result.sources || []).forEach((source) => allSources.add(source)));
  return Array.from(allSources);
}

function emptySearchResponse(): { success: true; data: ICopilotSearchResult[] } {
  return {
    success: true,
    data: []
  };
}

function describePublicWebCapabilityIssue(
  capability: 'blocked' | 'unsupported',
  detail?: string
): string {
  if (detail?.trim()) {
    return detail.trim();
  }

  return capability === 'blocked'
    ? 'Public web search is blocked for this tenant, subscription, or deployment.'
    : 'This Azure OpenAI deployment does not support web_search_preview.';
}

function buildPublicWebMarkdown(text: string, references: IPublicWebReference[]): string {
  const sections: string[] = [];
  if (text.trim()) {
    sections.push(text.trim());
  }

  if (references.length > 0) {
    sections.push([
      '### Sources',
      ...references.map((reference) => {
        const label = (reference.title || reference.url).trim();
        return `- [${label}](${reference.url})`;
      })
    ].join('\n'));
  }

  return sections.join('\n\n').trim();
}

function buildPublicWebErrorBlock(
  message: string,
  detail?: string
): { kind: 'error'; message: string; detail?: string } {
  return {
    kind: 'error',
    message,
    detail
  };
}

export function buildSearchRuntimeHandlers(
  helpers: ISearchRuntimeHelpers
): Pick<Record<RuntimeHandledToolName, ToolRuntimeHandler>, SearchRuntimeToolName> {
  const { findExistingSession, connectToM365Server } = helpers;

  return {
  search_sharepoint: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, aadClient } = deps;
    const query = (args.query as string || '').trim();
    const maxResults = safeParseMaxResults(args.max_results as string | undefined, 10);
    const renderHints = parseRenderHints(args);
    logService.info('search', `Search: "${query}"`);

    const loadingData: ISearchResultsData = {
      kind: 'search-results',
      query,
      results: [],
      totalCount: 0,
      source: 'pending'
    };
    const block = createBlock('search-results', `Search: ${query}`, loadingData, true, renderHints);
    trackCreatedBlock(store, block, deps);
    store.setExpression('thinking');

    if (!aadClient) {
      logService.warning('search', 'AadHttpClient not available — cannot search');
      return errorOutcome(JSON.stringify({ success: false, error: 'SharePoint connection not available. Please ensure you are signed in.' }));
    }

    const searchService = new CopilotSearchService(aadClient);
    const retrievalService = new CopilotRetrievalService(aadClient);
    const sharePointSearchService = new SharePointSearchService();
    const planner = new SearchIntentPlanner();
    const rrfService = new RRFFusionService();
    const searchTuning = getRuntimeTuningConfig().search;
    const userLanguage = store.userContext?.resolvedLanguage;
    const currentSiteUrl = store.userContext?.currentSiteUrl;
    const initialQueryLanguage = detectQueryLanguage(query, userLanguage);

    const searchPromise = searchService.search(
      query,
      Math.min(maxResults * searchTuning.primaryResultMultiplier, searchTuning.copilotSearchPageSizeCap)
    );
    const retrievalPromise = retrievalService.search(
      query,
      'sharePoint',
      Math.min(maxResults, searchTuning.copilotRetrievalMaxResultsCap)
    );
    const planPromise = planner.plan(query, { proxyConfig: store.proxyConfig, userLanguage });
    const sharePointPromise = planPromise.then((plan) => {
      const requestQueryText = plan.sharePointLexicalQuery?.trim();
      if (!requestQueryText) {
        logService.info(
          'search',
          `SharePoint Search skipped for "${query}" — planner did not produce a lexical query`
        );
        return emptySearchResponse();
      }

      return sharePointSearchService.search(query, {
        maxResults: Math.min(maxResults * searchTuning.primaryResultMultiplier, searchTuning.sharePointMaxResultsCap),
        queryLanguage: plan.queryLanguage || initialQueryLanguage,
        variantKind: 'raw',
        variantLanguage: plan.queryLanguage || initialQueryLanguage,
        requestQueryText
      });
    });

    const asyncResult = Promise.all([
      searchPromise,
      retrievalPromise,
      sharePointPromise,
      planPromise
    ]).then(async ([searchResp, retrievalResp, sharePointResp, plan]) => {
      const queryLanguage = plan.queryLanguage || initialQueryLanguage;
      const resultSets: ICopilotSearchResult[][] = [
        decorateResults(searchResp.success && searchResp.data ? searchResp.data : [], query, 'raw', queryLanguage),
        decorateResults(retrievalResp.success && retrievalResp.data ? retrievalResp.data : [], query, 'raw', queryLanguage),
        decorateResults(sharePointResp.success && sharePointResp.data ? sharePointResp.data : [], query, 'raw', queryLanguage)
      ].filter((set) => set.length > 0);
      const executedVariants: Array<{ kind: SearchQueryVariantKind; query: string; language?: string }> = [];
      const executedQueryKeys = new Set<string>([query.trim().toLowerCase()]);

      const fuseResults = (): ReturnType<RRFFusionService['fuseWithContext']> => rrfService.fuseWithContext(resultSets, {
        queryText: query,
        queryLanguage,
        currentSiteUrl
      });

      let fused = fuseResults();

      const runVariant = async (
        variantQuery: string,
        variantKind: SearchQueryVariantKind,
        variantLanguage: string,
        includeCopilot: boolean,
        includeSharePoint: boolean = false
      ): Promise<void> => {
        const variantKey = variantQuery.trim().toLowerCase();
        if (!variantKey || executedQueryKeys.has(variantKey)) return;

        const pending: Array<Promise<{ success: boolean; data?: ICopilotSearchResult[] }>> = [];
        if (includeCopilot) {
          pending.push(searchService.search(
            variantQuery,
            Math.min(maxResults * searchTuning.primaryResultMultiplier, searchTuning.copilotSearchPageSizeCap)
          ));
          pending.push(retrievalService.search(
            variantQuery,
            'sharePoint',
            Math.min(maxResults, searchTuning.copilotRetrievalMaxResultsCap)
          ));
        }
        if (includeSharePoint) {
          pending.push(sharePointSearchService.search(variantQuery, {
            maxResults: Math.min(maxResults * searchTuning.primaryResultMultiplier, searchTuning.sharePointMaxResultsCap),
            queryLanguage: variantLanguage,
            variantKind,
            variantLanguage
          }));
        }

        const responses = await Promise.all(pending);
        responses.forEach((response) => {
          if (!response.success || !response.data) return;
          resultSets.push(decorateResults(response.data, variantQuery, variantKind, variantLanguage));
        });

        executedVariants.push({ kind: variantKind, query: variantQuery, language: variantLanguage });
        executedQueryKeys.add(variantKey);
        fused = fuseResults();
      };

      if (
        plan.semanticRewriteQuery
        && plan.semanticRewriteQuery.toLowerCase() !== query.toLowerCase()
      ) {
        await runVariant(plan.semanticRewriteQuery, 'semantic-rewrite', queryLanguage, true, false);
      }

      if (
        fused.length < searchTuning.primaryUniqueResultThreshold
        && plan.correctedQuery
        && plan.correctedQuery.toLowerCase() !== query.toLowerCase()
      ) {
        await runVariant(plan.correctedQuery, 'corrected', queryLanguage, true, false);
      }

      if (
        fused.length < searchTuning.primaryUniqueResultThreshold
        && plan.translationFallbackQuery
        && plan.translationFallbackLanguage
        && plan.translationFallbackQuery.toLowerCase() !== query.toLowerCase()
      ) {
        await runVariant(plan.translationFallbackQuery, 'translation', plan.translationFallbackLanguage, true, false);
      }

      if (
        fused.length < searchTuning.primaryUniqueResultThreshold
        && plan.keywordFallbackQuery
        && plan.keywordFallbackQuery.toLowerCase() !== query.toLowerCase()
      ) {
        await runVariant(plan.keywordFallbackQuery, 'keyword-fallback', queryLanguage, false, true);
      }

      const blockResults = mapFusedResults(fused, maxResults);
      const sourceList = collectSourceList(blockResults);
      const variantInfo = buildVariantInfo(executedVariants);

      const updatedData: ISearchResultsData = {
        kind: 'search-results',
        query,
        results: blockResults,
        totalCount: fused.length,
        source: sourceList.join('+'),
        queryVariants: variantInfo
      };

      const currentStore = useGrimoireStore.getState();
      trackUpdatedBlock(currentStore, block.id, { data: updatedData }, { ...block, data: updatedData }, deps);
      trackToolCompletion('search_sharepoint', block.id, true, blockResults.length, deps);

      logService.info(
        'search',
        `Search complete: ${blockResults.length} results for "${query}" (sources: ${sourceList.join(', ')}${variantInfo ? `, variants: ${variantInfo.map((variant) => variant.kind).join(', ')}` : ''})`
      );
      return completeOutcome(JSON.stringify({
        success: true,
        query,
        displayedResults: blockResults.length,
        totalFound: fused.length,
        variants: variantInfo,
        results: blockResults.slice(0, 10).map((r) => ({ title: r.title, author: r.author, fileType: r.fileType, siteName: r.siteName }))
      }));
    }).catch((err: Error) => {
      const currentStore = useGrimoireStore.getState();
      const errorData: ISearchResultsData = {
        kind: 'search-results',
        query,
        results: [],
        totalCount: 0,
        source: 'error'
      };
      currentStore.updateBlock(block.id, { data: errorData });
      trackToolCompletion('search_sharepoint', block.id, false, 0, deps);
      logService.error('search', `Search failed: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Searching for "${query}"... Results will appear shortly.` }));
  },

  search_people: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, peopleService } = deps;
    const rawQuery = (args.query as string || '').trim();
    const query = normalizePeopleSearchQuery(rawQuery);
    const maxResults = safeParseMaxResults(args.max_results as string | undefined, 5);
    logService.info('search', `People search: "${query || rawQuery}"`);
    store.setExpression('thinking');

    if (isGenericPeopleQuery(query)) {
      const promptBlock = createBlock('info-card', 'People Search', {
        kind: 'info-card',
        heading: 'People Search',
        body: 'Tell me a person\'s name or email and I\'ll look them up.'
      } as IInfoCardData);
      trackCreatedBlock(store, promptBlock, deps);
      trackToolCompletion('search_people', '', true, 0, deps);
      return completeOutcome(JSON.stringify({
        success: true,
        query,
        count: 0,
        message: 'Please provide a person name or email.'
      }));
    }

    if (!peopleService) {
      logService.warning('search', 'AadHttpClient not available — cannot search people');
      return errorOutcome(JSON.stringify({ success: false, error: 'SharePoint connection not available. Please ensure you are signed in.' }));
    }

    const asyncResult = peopleService.searchPeople(query, maxResults).then((resp) => {
      const currentStore = useGrimoireStore.getState();
      const enrichPeople = async (people: IPersonResult[]): Promise<IPersonResult[]> => {
        if (!shouldEnrichPeopleSearch(query, people)) {
          return people;
        }

        const envId = currentStore.mcpEnvironmentId;
        const proxyConf = currentStore.proxyConfig;
        if (!envId || !proxyConf) {
          return people;
        }

        const serverUrl = resolveServerUrl('mcp_MeServer', envId);
        const mcpClient = new McpClientService(proxyConf.proxyUrl, proxyConf.proxyApiKey);

        return Promise.all(people.map(async (person) => {
          const userIdentifier = (person.email || person.displayName || '').trim();
          if (!userIdentifier) return person;

          try {
            const execution = await executeCatalogMcpTool({
              serverId: 'mcp_MeServer',
              serverName: 'User Profile',
              serverUrl,
              toolName: 'GetUserDetails',
              rawArgs: {
                userIdentifier,
                select: PERSON_PROFILE_SELECT
              },
              connections: currentStore.mcpConnections,
              getConnections: () => useGrimoireStore.getState().mcpConnections,
              mcpClient,
              sessionHelpers: {
                findExistingSession,
                connectToM365Server
              },
              getToken: currentStore.getToken,
              sourceContext: deps.sourceContext,
              taskContext: hybridInteractionEngine.getCurrentTaskContext(),
              artifacts: hybridInteractionEngine.getCurrentArtifacts(),
              currentSiteUrl: currentStore.userContext?.currentSiteUrl
            });

            logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));
            if (!execution.success || !execution.mcpResult) {
              return person;
            }

            const payload = extractProfilePayload(execution.mcpResult.content);
            if (!payload) {
              return person;
            }

            return mergePersonProfile(person, payload);
          } catch (error) {
            logService.warning('search', `People profile enrichment failed for "${userIdentifier}": ${(error as Error).message}`);
            return person;
          }
        }));
      };

      if (resp.success && resp.data && resp.data.length > 0) {
        return enrichPeople(resp.data.slice(0, maxResults)).then((people) => {
          people.forEach((person) => {
          const cardData: IUserCardData = {
            kind: 'user-card',
            displayName: person.displayName,
            email: person.email,
            jobTitle: person.jobTitle,
            department: person.department,
            officeLocation: person.officeLocation,
            phone: person.phone
          };
          const cardBlock = createBlock('user-card', person.displayName, cardData);
          trackCreatedBlock(currentStore, cardBlock, deps);
          });
          trackToolCompletion('search_people', '', true, people.length, deps);
          return completeOutcome(JSON.stringify({
            success: true,
            query,
            count: people.length,
            people: people.map((p) => ({ name: p.displayName, title: p.jobTitle, email: p.email }))
          }));
        });
      }

      const noResultBlock = createBlock('info-card', `People: ${query}`, {
        kind: 'info-card',
        heading: `People: ${query}`,
        body: 'No people found.'
      } as IInfoCardData);
      trackCreatedBlock(currentStore, noResultBlock, deps);
      trackToolCompletion('search_people', '', true, 0, deps);
      return completeOutcome(JSON.stringify({ success: true, query, count: 0, message: 'No people found.' }));
    }).catch((err: Error) => {
      trackToolCompletion('search_people', '', false, 0, deps);
      logService.error('search', `People search failed: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Searching people for "${query || rawQuery}"...` }));
  },

  search_sites: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync, sitesService } = deps;
    const query = args.query as string;
    const maxResults = Math.max(1, parseInt((args.max_results as string) || '10', 10) || 10);
    logService.info('search', `Site search: "${query}"`);
    store.setExpression('thinking');

    if (!sitesService) {
      logService.warning('search', 'AadHttpClient not available — cannot search sites');
      return errorOutcome(JSON.stringify({ success: false, error: 'SharePoint connection not available. Please ensure you are signed in.' }));
    }

    const asyncResult = sitesService.searchSites(query).then((resp) => {
      const currentStore = useGrimoireStore.getState();
      if (resp.success && resp.data && resp.data.length > 0) {
        const seen = new Set<string>();
        const unique = resp.data.filter((site) => {
          const key = (site.siteUrl || '').toLowerCase();
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        }).slice(0, maxResults);

        if (unique.length === 1) {
          const site = unique[0];
          const siteData: ISiteInfoData = {
            kind: 'site-info',
            siteName: site.siteName,
            siteUrl: site.siteUrl,
            description: site.description,
            created: site.created,
            lastModified: site.lastModified,
            libraries: site.libraries,
            lists: site.lists
          };
          const siteBlock = createBlock('site-info', site.siteName, siteData);
          trackCreatedBlock(currentStore, siteBlock, deps);
          trackToolCompletion('search_sites', siteBlock.id, true, 1, deps);
        } else {
          const selectionData: ISelectionListData = {
            kind: 'selection-list',
            prompt: `Found ${unique.length} sites for "${query}"`,
            items: unique.map((site) => ({
              id: site.siteUrl,
              label: site.siteName,
              description: site.description || site.siteUrl,
              selected: false
            })),
            multiSelect: false
          };
          const selBlock = createBlock('selection-list', `Sites: ${query}`, selectionData, false);
          trackCreatedBlock(currentStore, selBlock, deps);
          trackToolCompletion('search_sites', selBlock.id, true, unique.length, deps);
        }

        return completeOutcome(JSON.stringify({
          success: true,
          query,
          count: unique.length,
          sites: unique.map((s) => ({ name: s.siteName, url: s.siteUrl }))
        }));
      }

      const noResultBlock = createBlock('info-card', `Sites: ${query}`, {
        kind: 'info-card',
        heading: `Sites: ${query}`,
        body: 'No sites found.'
      } as IInfoCardData);
      trackCreatedBlock(currentStore, noResultBlock, deps);
      trackToolCompletion('search_sites', '', true, 0, deps);
      return completeOutcome(JSON.stringify({ success: true, query, count: 0, message: 'No sites found.' }));
    }).catch((err: Error) => {
      trackToolCompletion('search_sites', '', false, 0, deps);
      logService.error('search', `Site search failed: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Searching sites for "${query}"...` }));
  },

  research_public_web: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync } = deps;
    const targetUrl = typeof args.target_url === 'string' ? args.target_url.trim() : undefined;
    const query = (typeof args.query === 'string' ? args.query.trim() : '')
      || (targetUrl ? `Summarize the page at ${targetUrl}` : '');

    logService.info(
      'search',
      `Public web research: "${query}"${targetUrl ? ` (${targetUrl})` : ''}`
    );
    store.setExpression('thinking');

    if (!query) {
      return errorOutcome(JSON.stringify({
        success: false,
        capabilityStatus: 'error',
        error: 'Public web search query cannot be empty.'
      }));
    }

    if (!store.publicWebSearchEnabled) {
      const currentStore = useGrimoireStore.getState();
      const block = createBlock(
        'error',
        'Public Web Search',
        buildPublicWebErrorBlock(
          'Public web search is disabled in settings.',
          'Enable "Public Web Search (Preview)" in avatar settings to research websites and URLs.'
        ) as IErrorData,
        true,
        undefined,
        { originTool: 'research_public_web' }
      );
      trackCreatedBlock(currentStore, block, deps);
      trackToolCompletion('research_public_web', block.id, false, 0, deps);
      return errorOutcome(JSON.stringify({
        success: false,
        capabilityStatus: 'error',
        error: 'Public web search is disabled in settings.'
      }));
    }

    if (!store.proxyConfig) {
      return errorOutcome(JSON.stringify({
        success: false,
        capabilityStatus: 'error',
        error: 'Proxy configuration is missing for public web search.'
      }));
    }

    if (store.publicWebSearchCapability === 'blocked' || store.publicWebSearchCapability === 'unsupported') {
      const detail = describePublicWebCapabilityIssue(
        store.publicWebSearchCapability,
        store.publicWebSearchCapabilityDetail
      );
      const currentStore = useGrimoireStore.getState();
      const block = createBlock(
        'error',
        'Public Web Search',
        buildPublicWebErrorBlock('Public web search is unavailable.', detail) as IErrorData,
        true,
        undefined,
        { originTool: 'research_public_web' }
      );
      trackCreatedBlock(currentStore, block, deps);
      trackToolCompletion('research_public_web', block.id, false, 0, deps);
      return errorOutcome(JSON.stringify({
        success: false,
        capabilityStatus: store.publicWebSearchCapability,
        error: detail
      }));
    }

    const service = new PublicWebSearchService(store.proxyConfig);
    const asyncResult = service.research(query, targetUrl).then((result) => {
      const currentStore = useGrimoireStore.getState();

      if (!result.success) {
        if (result.capabilityStatus !== 'available') {
          currentStore.setPublicWebSearchCapability?.(result.capabilityStatus, result.error?.message);
        }

        const block = createBlock(
          'error',
          'Public Web Search',
          buildPublicWebErrorBlock(
            'Public web research failed.',
            result.error?.message
          ) as IErrorData,
          true,
          undefined,
          { originTool: 'research_public_web' }
        );
        trackCreatedBlock(currentStore, block, deps);
        trackToolCompletion('research_public_web', block.id, false, 0, deps);
        return errorOutcome(JSON.stringify({
          success: false,
          capabilityStatus: result.capabilityStatus,
          error: result.error?.message || 'Public web research failed.'
        }));
      }

      currentStore.setPublicWebSearchCapability?.('available');

      const content = buildPublicWebMarkdown(result.text || '', result.references);
      const markdownData: IMarkdownData = {
        kind: 'markdown',
        content
      };
      const block = createBlock(
        'markdown',
        targetUrl ? 'Public Web Research' : `Public Web: ${query}`,
        markdownData,
        true,
        undefined,
        { originTool: 'research_public_web' }
      );
      trackCreatedBlock(currentStore, block, deps);
      trackToolCompletion('research_public_web', block.id, true, result.references.length, deps);

      return completeOutcome(JSON.stringify({
        success: true,
        query,
        targetUrl,
        referenceCount: result.references.length,
        references: result.references,
        content
      }));
    }).catch((error: Error) => {
      const currentStore = useGrimoireStore.getState();
      currentStore.setPublicWebSearchCapability?.('error', error.message);
      const block = createBlock(
        'error',
        'Public Web Search',
        buildPublicWebErrorBlock('Public web research failed.', error.message) as IErrorData,
        true,
        undefined,
        { originTool: 'research_public_web' }
      );
      trackCreatedBlock(currentStore, block, deps);
      trackToolCompletion('research_public_web', block.id, false, 0, deps);
      return errorOutcome(JSON.stringify({
        success: false,
        capabilityStatus: 'error',
        error: error.message
      }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({
      success: true,
      message: `Researching the public web for "${query}"...`
    }));
  },

  search_emails: (args, deps): ToolRuntimeHandlerResult | Promise<ToolRuntimeHandlerResult> => {
    const { store, awaitAsync } = deps;
    const query = (args.query as string) || '*';
    const folder = args.folder as string | undefined;
    const maxResults = args.max_results ? parseInt(args.max_results as string, 10) : undefined;
    logService.info('search', `Email search: "${query}"${maxResults ? ` (max ${maxResults})` : ''}`);
    store.setExpression('thinking');

    const envId = store.mcpEnvironmentId;
    const proxyConf = store.proxyConfig;
    if (!envId || !proxyConf) {
      return errorOutcome(JSON.stringify({ success: false, error: 'MCP environment not configured.' }));
    }

    const serverUrl = resolveServerUrl('mcp_MailTools', envId);
    const mcpClient = new McpClientService(proxyConf.proxyUrl, proxyConf.proxyApiKey);

    const executeEmailSearch = async (): Promise<ToolRuntimeHandlerResult> => {
      const currentStore = useGrimoireStore.getState();
      let searchMessage = query;
      if (maxResults) searchMessage += `. Show only the ${maxResults} most recent emails.`;
      searchMessage += ' List each email individually as a numbered list. For each email show **Subject:**, **From:**, **To:**, **CC:**, **Date:**, and a one-line **Preview:**. Do not group, categorize, or summarize.';
      if (folder) searchMessage += ` Search in folder: ${folder}.`;

      const execution = await executeCatalogMcpTool({
        serverId: 'mcp_MailTools',
        serverName: 'Outlook Mail',
        serverUrl,
        toolName: 'SearchMessages',
        rawArgs: { message: searchMessage },
        connections: currentStore.mcpConnections,
        getConnections: () => useGrimoireStore.getState().mcpConnections,
        mcpClient,
        sessionHelpers: {
          findExistingSession,
          connectToM365Server
        },
        getToken: currentStore.getToken,
        sourceContext: deps.sourceContext,
        taskContext: hybridInteractionEngine.getCurrentTaskContext(),
        artifacts: hybridInteractionEngine.getCurrentArtifacts(),
        currentSiteUrl: currentStore.userContext?.currentSiteUrl
      });
      logService.debug('mcp', 'MCP execution trace', JSON.stringify(execution.trace));
      if (!execution.success || !execution.mcpResult) {
        throw new Error(execution.error || 'Email search failed');
      }

      let emailBlockId = '';
      let emailItemCount = 0;
      mapMcpResultToBlock(
        'mcp_MailTools',
        execution.realToolName,
        execution.mcpResult.content,
        useGrimoireStore.getState().pushBlock,
        (block) => {
          emailBlockId = block.id;
          hybridInteractionEngine.onBlockCreated(block, deps.sourceContext);
          const tracked = hybridInteractionEngine.getBlockTracker().get(block.id);
          emailItemCount = tracked ? tracked.itemCount : 0;
        }
      );
      trackToolCompletion('search_emails', emailBlockId, true, emailItemCount, deps);

      return completeOutcome(JSON.stringify({
        success: true,
        query,
        count: emailItemCount,
        displayed: true,
        note: 'Email results displayed. search_emails does not return message IDs — do NOT pass invented IDs to GetMessage. To read full email content, use read_email_content with subject + sender from the selected result. For reply/forward, use show_compose_form.'
      }));
    };

    const asyncResult = executeEmailSearch().catch((err: Error) => {
      const errBlock = createBlock('error', 'Email Search Error', {
        kind: 'error',
        message: err.message
      } as IErrorData);
      trackCreatedBlock(useGrimoireStore.getState(), errBlock, deps);
      trackToolCompletion('search_emails', errBlock.id, false, 0, deps);
      logService.error('search', `Email search error: ${err.message}`);
      return errorOutcome(JSON.stringify({ success: false, error: err.message }));
    });

    if (awaitAsync) return asyncResult;
    return completeOutcome(JSON.stringify({ success: true, message: `Searching emails for "${query}"... Results will appear in the panel shortly. Tell the user to hold on a moment.` }));
  },
  };
}
