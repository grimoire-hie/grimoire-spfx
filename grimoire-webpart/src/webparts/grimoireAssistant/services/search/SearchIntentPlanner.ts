import type { ISearchIntentPlan } from '../../models/ISearchTypes';
import type { IProxyConfig } from '../../store/useGrimoireStore';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { buildSearchPlannerSystemPrompt } from '../../config/promptCatalog';
import { logService } from '../logging/LogService';
import { getNanoService } from '../nano/NanoService';
import {
  detectQueryLanguage,
  normalizeLanguageTag,
  shouldUseTranslationFallback
} from './SearchLanguageUtils';

interface ISearchIntentPlannerOptions {
  proxyConfig?: IProxyConfig;
  userLanguage?: string;
}

interface IPlannerResponse {
  queryLanguage?: string;
  semanticRewriteQuery?: string;
  semanticRewriteConfidence?: number;
  sharePointLexicalQuery?: string;
  sharePointLexicalConfidence?: number;
  correctedQuery?: string;
  correctionConfidence?: number;
  translationFallbackQuery?: string;
  translationFallbackLanguage?: string;
  keywordFallbackQuery?: string;
}

interface IPlannerDecisionLog {
  candidate?: string;
  accepted: boolean;
  reason: string;
  confidence?: number;
  threshold?: number;
}

function normalizeJson(raw: string): string {
  const trimmed = raw.trim();
  if (trimmed.startsWith('```')) {
    return trimmed.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  }
  return trimmed;
}

function buildConfidenceDecisionLog(
  candidate: string,
  acceptedValue: string | undefined,
  rawQuery: string,
  confidence: number,
  threshold: number
): IPlannerDecisionLog {
  if (!candidate) {
    return {
      accepted: false,
      reason: 'no_candidate',
      confidence,
      threshold
    };
  }

  if (candidate.toLowerCase() === rawQuery.toLowerCase()) {
    return {
      candidate,
      accepted: false,
      reason: 'same_as_raw_query',
      confidence,
      threshold
    };
  }

  if (!acceptedValue) {
    return {
      candidate,
      accepted: false,
      reason: confidence < threshold ? 'below_confidence_threshold' : 'rejected',
      confidence,
      threshold
    };
  }

  return {
    candidate,
    accepted: true,
    reason: 'accepted',
    confidence,
    threshold
  };
}

function buildVariantDecisionLog(
  candidate: string,
  acceptedValue: string | undefined,
  rawQuery: string,
  fallbackSource?: string
): IPlannerDecisionLog {
  if (!candidate) {
    return {
      accepted: false,
      reason: fallbackSource ? `derived_from_${fallbackSource}` : 'no_candidate'
    };
  }

  if (candidate.toLowerCase() === rawQuery.toLowerCase()) {
    return {
      candidate,
      accepted: false,
      reason: 'same_as_raw_query'
    };
  }

  if (!acceptedValue) {
    return {
      candidate,
      accepted: false,
      reason: 'rejected'
    };
  }

  return {
    candidate,
    accepted: true,
    reason: fallbackSource ? `accepted_via_${fallbackSource}` : 'accepted'
  };
}

export class SearchIntentPlanner {
  public async plan(query: string, options: ISearchIntentPlannerOptions = {}): Promise<ISearchIntentPlan> {
    const rawQuery = query.trim();
    const heuristicLanguage = detectQueryLanguage(rawQuery, options.userLanguage);
    const tuning = getRuntimeTuningConfig();
    const plannerConfidence = tuning.search.plannerConfidence;
    const fallbackPlan: ISearchIntentPlan = {
      rawQuery,
      queryLanguage: heuristicLanguage,
      usedCorrection: false,
      usedTranslation: false
    };

    const nano = getNanoService(options.proxyConfig);
    if (!nano) {
      logService.debug('search', 'SearchIntentPlanner: no Nano service, using heuristic plan');
      fallbackPlan.translationFallbackLanguage = shouldUseTranslationFallback(heuristicLanguage, options.userLanguage);
      return fallbackPlan;
    }

    try {
      const content = await nano.classify(
        buildSearchPlannerSystemPrompt(normalizeLanguageTag(options.userLanguage) || 'en'),
        rawQuery,
        tuning.nano.searchPlannerTimeoutMs,
        tuning.nano.searchPlannerMaxTokens
      );
      if (!content) {
        fallbackPlan.translationFallbackLanguage = shouldUseTranslationFallback(heuristicLanguage, options.userLanguage);
        return fallbackPlan;
      }

      const parsed = JSON.parse(normalizeJson(content)) as IPlannerResponse;
      const queryLanguage = normalizeLanguageTag(parsed.queryLanguage) || heuristicLanguage;
      const semanticRewriteCandidate = typeof parsed.semanticRewriteQuery === 'string' ? parsed.semanticRewriteQuery.trim() : '';
      const semanticRewriteConfidence = typeof parsed.semanticRewriteConfidence === 'number' ? parsed.semanticRewriteConfidence : 0;
      const semanticRewriteQuery = semanticRewriteCandidate
        && semanticRewriteCandidate.toLowerCase() !== rawQuery.toLowerCase()
        && semanticRewriteConfidence >= plannerConfidence.semanticRewrite
        ? semanticRewriteCandidate
        : undefined;
      const sharePointLexicalCandidate = typeof parsed.sharePointLexicalQuery === 'string' ? parsed.sharePointLexicalQuery.trim() : '';
      const sharePointLexicalConfidence = typeof parsed.sharePointLexicalConfidence === 'number' ? parsed.sharePointLexicalConfidence : 0;
      const sharePointLexicalQuery = sharePointLexicalCandidate
        && sharePointLexicalConfidence >= plannerConfidence.sharePointLexical
        ? sharePointLexicalCandidate
        : undefined;
      const correctedCandidate = typeof parsed.correctedQuery === 'string' ? parsed.correctedQuery.trim() : '';
      const correctionConfidence = typeof parsed.correctionConfidence === 'number' ? parsed.correctionConfidence : 0;
      const correctedQuery = correctedCandidate && correctedCandidate.toLowerCase() !== rawQuery.toLowerCase() && correctionConfidence >= plannerConfidence.correction
        ? correctedCandidate
        : undefined;

      const translationFallbackLanguage =
        normalizeLanguageTag(typeof parsed.translationFallbackLanguage === 'string' ? parsed.translationFallbackLanguage : undefined)
        || shouldUseTranslationFallback(queryLanguage, options.userLanguage);
      const translationCandidate = typeof parsed.translationFallbackQuery === 'string'
        ? parsed.translationFallbackQuery.trim()
        : '';
      const translationFallbackQuery = translationFallbackLanguage && translationCandidate
        && translationCandidate.toLowerCase() !== rawQuery.toLowerCase()
        ? translationCandidate
        : undefined;

      const keywordFallbackCandidate = typeof parsed.keywordFallbackQuery === 'string' && parsed.keywordFallbackQuery.trim()
        ? parsed.keywordFallbackQuery.trim()
        : sharePointLexicalQuery;
      const keywordFallbackQuery = keywordFallbackCandidate;

      const semanticDecision = buildConfidenceDecisionLog(
        semanticRewriteCandidate,
        semanticRewriteQuery,
        rawQuery,
        semanticRewriteConfidence,
        plannerConfidence.semanticRewrite
      );
      const sharePointLexicalDecision = buildConfidenceDecisionLog(
        sharePointLexicalCandidate,
        sharePointLexicalQuery,
        rawQuery,
        sharePointLexicalConfidence,
        plannerConfidence.sharePointLexical
      );
      const correctionDecision = buildConfidenceDecisionLog(
        correctedCandidate,
        correctedQuery,
        rawQuery,
        correctionConfidence,
        plannerConfidence.correction
      );
      const translationDecision = buildVariantDecisionLog(
        translationCandidate,
        translationFallbackQuery,
        rawQuery
      );
      const keywordDecision = buildVariantDecisionLog(
        keywordFallbackCandidate || '',
        keywordFallbackQuery,
        rawQuery,
        keywordFallbackCandidate === sharePointLexicalQuery ? 'sharePointLexicalQuery' : undefined
      );

      const plan: ISearchIntentPlan = {
        rawQuery,
        semanticRewriteQuery,
        semanticRewriteConfidence: semanticRewriteQuery ? semanticRewriteConfidence : undefined,
        sharePointLexicalQuery,
        sharePointLexicalConfidence: sharePointLexicalQuery ? sharePointLexicalConfidence : undefined,
        queryLanguage,
        correctedQuery,
        translationFallbackQuery,
        translationFallbackLanguage,
        keywordFallbackQuery,
        usedCorrection: !!correctedQuery,
        usedTranslation: !!translationFallbackQuery
      };

      const logPayload: Record<string, unknown> = {
        rawQuery: plan.rawQuery,
        queryLanguage: plan.queryLanguage,
        semanticRewriteQuery: plan.semanticRewriteQuery,
        semanticRewriteConfidence: plan.semanticRewriteConfidence,
        semanticDecision,
        sharePointLexicalQuery: plan.sharePointLexicalQuery,
        sharePointLexicalConfidence: plan.sharePointLexicalConfidence,
        sharePointLexicalDecision,
        correctedCandidate: correctedCandidate || undefined,
        correctedQuery: plan.correctedQuery,
        correctionConfidence,
        correctionDecision,
        translationCandidate: translationCandidate || undefined,
        translationFallbackQuery: plan.translationFallbackQuery,
        translationFallbackLanguage: plan.translationFallbackLanguage,
        translationDecision,
        keywordFallbackQuery: plan.keywordFallbackQuery,
        keywordDecision
      };

      logService.info(
        'search',
        `SearchIntentPlanner: lang=${plan.queryLanguage}, semantic=${plan.semanticRewriteQuery ? 'yes' : 'no'}, sharepoint=${plan.sharePointLexicalQuery ? 'yes' : 'raw'}, corrected=${plan.correctedQuery ? 'yes' : 'no'}, translation=${plan.translationFallbackQuery ? plan.translationFallbackLanguage : 'no'}`,
        JSON.stringify(logPayload, null, 2)
      );

      return plan;
    } catch (error) {
      logService.debug('search', `SearchIntentPlanner: parse error ${(error as Error).message}`);
      fallbackPlan.translationFallbackLanguage = shouldUseTranslationFallback(heuristicLanguage, options.userLanguage);
      return fallbackPlan;
    }
  }
}
