/**
 * RRFFusionService
 * Weighted Reciprocal Rank Fusion with deterministic rerank boosts.
 */

import type { ICopilotSearchResult, IRRFResult } from '../../models/ISearchTypes';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { logService } from '../logging/LogService';

interface IAggregatedEntry {
  result: ICopilotSearchResult;
  rrfScore: number;
  sources: Set<string>;
  variantKinds: Set<NonNullable<ICopilotSearchResult['variantKind']>>;
  variantLanguages: Set<string>;
  nativeScores: Map<string, number>;
}

export interface IRRFFusionOptions {
  queryText?: string;
  queryLanguage?: string;
  currentSiteUrl?: string;
  sourceWeights?: Record<string, number>;
  variantWeights?: Record<string, number>;
}

export class RRFFusionService {
  public fuse(...resultSets: ICopilotSearchResult[][]): IRRFResult[] {
    return this.fuseWithContext(resultSets);
  }

  public fuseWithContext(resultSets: ICopilotSearchResult[][], options: IRRFFusionOptions = {}): IRRFResult[] {
    const startTime = performance.now();
    const aggregated = new Map<string, IAggregatedEntry>();
    const tuning = getRuntimeTuningConfig().search.fusion;
    const sourceWeights = { ...tuning.sourceWeights, ...(options.sourceWeights || {}) };
    const variantWeights = { ...tuning.variantWeights, ...(options.variantWeights || {}) };
    const rrfK = tuning.k;
    const maxNativeScoreBySource = this.collectMaxNativeScores(resultSets);

    resultSets.forEach((results) => {
      results.forEach((result, position) => {
        const key = this.normalizeUrl(result.url);
        const sourceWeight = sourceWeights[result.source] ?? 0.7;
        const variantKey = result.variantKind || 'raw';
        const variantWeight = variantWeights[variantKey] ?? 1;
        const contribution = (sourceWeight * variantWeight) / (rrfK + position + 1);
        const existing = aggregated.get(key);

        if (existing) {
          existing.rrfScore += contribution;
          existing.sources.add(result.source);
          existing.variantKinds.add(variantKey);
          if (result.variantLanguage) existing.variantLanguages.add(result.variantLanguage);
          if (typeof result.sourceNativeScore === 'number') {
            const current = existing.nativeScores.get(result.source) || 0;
            if (result.sourceNativeScore > current) {
              existing.nativeScores.set(result.source, result.sourceNativeScore);
            }
          }
          if (result.summary.length > existing.result.summary.length) {
            existing.result = { ...existing.result, summary: result.summary };
          }
          if (!existing.result.author && result.author) existing.result.author = result.author;
          if (!existing.result.lastModified && result.lastModified) existing.result.lastModified = result.lastModified;
          if (!existing.result.siteName && result.siteName) existing.result.siteName = result.siteName;
          if (!existing.result.fileType && result.fileType) existing.result.fileType = result.fileType;
          if (!existing.result.language && result.language) existing.result.language = result.language;
        } else {
          const sources = new Set<string>();
          sources.add(result.source);
          const variantKinds = new Set<NonNullable<ICopilotSearchResult['variantKind']>>();
          variantKinds.add(variantKey);
          const variantLanguages = new Set<string>();
          if (result.variantLanguage) variantLanguages.add(result.variantLanguage);
          const nativeScores = new Map<string, number>();
          if (typeof result.sourceNativeScore === 'number') {
            nativeScores.set(result.source, result.sourceNativeScore);
          }
          aggregated.set(key, {
            result: { ...result },
            rrfScore: contribution,
            sources,
            variantKinds,
            variantLanguages,
            nativeScores
          });
        }
      });
    });

    const fused: IRRFResult[] = [];
    aggregated.forEach((entry) => {
      const sourcesArray: string[] = [];
      entry.sources.forEach((s) => sourcesArray.push(s));
      const variantKindsArray = Array.from(entry.variantKinds);
      const variantLanguagesArray = Array.from(entry.variantLanguages);
      const finalScore = entry.rrfScore + this.computeRerankBoost(entry, maxNativeScoreBySource, options);

      fused.push({
        url: entry.result.url,
        title: entry.result.title,
        summary: entry.result.summary,
        rrfScore: finalScore,
        sources: sourcesArray,
        variantKinds: variantKindsArray,
        variantLanguages: variantLanguagesArray,
        fileType: entry.result.fileType,
        lastModified: entry.result.lastModified,
        author: entry.result.author,
        siteName: entry.result.siteName,
        language: entry.result.language
      });
    });

    fused.sort((a, b) => b.rrfScore - a.rrfScore);

    const durationMs = Math.round(performance.now() - startTime);
    const totalInputs = resultSets.reduce((sum, rs) => sum + rs.length, 0);
    logService.info('search', `Weighted RRF Fusion: ${totalInputs} inputs → ${fused.length} unique results (K=${rrfK})`, undefined, durationMs);

    return fused;
  }

  private collectMaxNativeScores(resultSets: ICopilotSearchResult[][]): Map<string, number> {
    const maxScores = new Map<string, number>();
    resultSets.forEach((results) => {
      results.forEach((result) => {
        if (typeof result.sourceNativeScore !== 'number' || !Number.isFinite(result.sourceNativeScore)) return;
        const current = maxScores.get(result.source) || 0;
        if (result.sourceNativeScore > current) {
          maxScores.set(result.source, result.sourceNativeScore);
        }
      });
    });
    return maxScores;
  }

  private computeRerankBoost(
    entry: IAggregatedEntry,
    maxNativeScoreBySource: Map<string, number>,
    options: IRRFFusionOptions
  ): number {
    let boost = 0;
    const boosts = getRuntimeTuningConfig().search.fusion.boosts;

    if (options.queryLanguage) {
      if (entry.result.language && entry.result.language === options.queryLanguage) {
        boost += boosts.exactQueryLanguage;
      } else if (entry.variantLanguages.has(options.queryLanguage)) {
        boost += boosts.variantQueryLanguage;
      }
    }

    if (options.queryText && this.titleContainsQuery(entry.result.title, options.queryText)) {
      boost += boosts.titleMatch;
    }

    if (options.currentSiteUrl && this.isCurrentSiteMatch(entry.result.url, options.currentSiteUrl)) {
      boost += boosts.currentSite;
    }

    boost += this.normalizeNativeScoreBoost(entry.nativeScores.get('copilot-retrieval'), maxNativeScoreBySource.get('copilot-retrieval'), boosts.retrievalNative);
    boost += this.normalizeNativeScoreBoost(entry.nativeScores.get('sharepoint-search'), maxNativeScoreBySource.get('sharepoint-search'), boosts.sharePointNative);

    return boost;
  }

  private normalizeNativeScoreBoost(value: number | undefined, max: number | undefined, cap: number): number {
    if (typeof value !== 'number' || typeof max !== 'number' || max <= 0) return 0;
    return Math.min(1, value / max) * cap;
  }

  private titleContainsQuery(title: string, queryText: string): boolean {
    const normalizedTitle = this.normalizeText(title);
    const normalizedQuery = this.normalizeText(queryText);
    if (!normalizedTitle || !normalizedQuery || normalizedQuery.length < 3) return false;
    if (normalizedTitle.includes(normalizedQuery)) return true;
    const tokens = normalizedQuery.split(' ').filter((token) => token.length > 2);
    if (tokens.length < 2) return false;
    return tokens.every((token) => normalizedTitle.includes(token));
  }

  private normalizeText(value: string): string {
    return value
      .toLowerCase()
      .replace(/[^0-9a-z\u00c0-\u024f\u0400-\u04ff\u0600-\u06ff\u0590-\u05ff\u3040-\u30ff\u4e00-\u9fff\uac00-\ud7af\u0e00-\u0e7f]+/gi, ' ')
      .trim();
  }

  private isCurrentSiteMatch(url: string, currentSiteUrl: string): boolean {
    const normalizedResultUrl = this.normalizeUrl(url);
    const normalizedCurrentSite = this.normalizeUrl(currentSiteUrl);
    return normalizedResultUrl.startsWith(normalizedCurrentSite);
  }

  private normalizeUrl(url: string): string {
    try {
      const u = new URL(url);
      return (u.origin + u.pathname).replace(/\/+$/, '').toLowerCase();
    } catch {
      return url.toLowerCase().replace(/[?#].*$/, '').replace(/\/+$/, '');
    }
  }
}
