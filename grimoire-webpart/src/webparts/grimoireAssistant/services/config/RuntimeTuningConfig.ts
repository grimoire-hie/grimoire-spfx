/**
 * RuntimeTuningConfig
 * Centralized runtime knobs for fast-model, search orchestration, and ranking behavior.
 * Host pages can override these defaults by populating window.__GRIMOIRE_RUNTIME_TUNING__.
 */

export interface INanoRuntimeConfig {
  defaultTimeoutMs: number;
  cooldownMs: number;
  reasoningEffort?: 'minimal' | 'low' | 'medium' | 'high';
  searchPlannerTimeoutMs: number;
  searchPlannerMaxTokens: number;
  compoundWorkflowPlannerTimeoutMs: number;
  compoundWorkflowPlannerMaxTokens: number;
  compoundWorkflowPlannerConfidenceThreshold: number;
  intentRoutingTimeoutMs: number;
  intentRoutingMaxTokens: number;
  intentRoutingConfidenceThreshold: number;
  contextCompressionTimeoutMs: number;
  sentimentTimeoutMs: number;
  blockRecapTimeoutMs: number;
  blockRecapMaxTokens: number;
  blockRecapRetryHeadroomTokens: number;
  blockRecapRetryMinTokens: number;
}

export interface ISearchPlannerConfidenceConfig {
  correction: number;
  semanticRewrite: number;
  sharePointLexical: number;
}

export interface ISearchFusionBoostConfig {
  exactQueryLanguage: number;
  variantQueryLanguage: number;
  titleMatch: number;
  currentSite: number;
  retrievalNative: number;
  sharePointNative: number;
}

export interface ISearchFusionConfig {
  k: number;
  sourceWeights: Record<string, number>;
  variantWeights: Record<string, number>;
  boosts: ISearchFusionBoostConfig;
}

export interface ISearchRuntimeConfig {
  primaryUniqueResultThreshold: number;
  primaryResultMultiplier: number;
  copilotSearchPageSizeCap: number;
  copilotRetrievalMaxResultsCap: number;
  sharePointMaxResultsCap: number;
  plannerConfidence: ISearchPlannerConfidenceConfig;
  fusion: ISearchFusionConfig;
}

export interface IRuntimeTuningConfig {
  nano: INanoRuntimeConfig;
  search: ISearchRuntimeConfig;
}

export interface IRuntimeTuningOverrides {
  nano?: Partial<INanoRuntimeConfig>;
  search?: Partial<Omit<ISearchRuntimeConfig, 'plannerConfidence' | 'fusion'>> & {
    plannerConfidence?: Partial<ISearchPlannerConfidenceConfig>;
    fusion?: Partial<Omit<ISearchFusionConfig, 'sourceWeights' | 'variantWeights' | 'boosts'>> & {
      sourceWeights?: Record<string, number>;
      variantWeights?: Record<string, number>;
      boosts?: Partial<ISearchFusionBoostConfig>;
    };
  };
}

declare global {
  interface Window {
    __GRIMOIRE_RUNTIME_TUNING__?: IRuntimeTuningOverrides;
  }
}

const DEFAULT_RUNTIME_TUNING_CONFIG: IRuntimeTuningConfig = {
  nano: {
    defaultTimeoutMs: 4000,
    cooldownMs: 30000,
    reasoningEffort: 'minimal',
    searchPlannerTimeoutMs: 8000,
    searchPlannerMaxTokens: 240,
    compoundWorkflowPlannerTimeoutMs: 4500,
    compoundWorkflowPlannerMaxTokens: 120,
    compoundWorkflowPlannerConfidenceThreshold: 0.78,
    intentRoutingTimeoutMs: 5000,
    intentRoutingMaxTokens: 120,
    intentRoutingConfidenceThreshold: 0.65,
    contextCompressionTimeoutMs: 4000,
    sentimentTimeoutMs: 2500,
    blockRecapTimeoutMs: 8000,
    blockRecapMaxTokens: 640,
    blockRecapRetryHeadroomTokens: 160,
    blockRecapRetryMinTokens: 800
  },
  search: {
    primaryUniqueResultThreshold: 5,
    primaryResultMultiplier: 2,
    copilotSearchPageSizeCap: 50,
    copilotRetrievalMaxResultsCap: 25,
    sharePointMaxResultsCap: 25,
    plannerConfidence: {
      correction: 0.8,
      semanticRewrite: 0.8,
      sharePointLexical: 0.55
    },
    fusion: {
      k: 60,
      sourceWeights: {
        'copilot-retrieval': 1.0,
        'sharepoint-search': 0.6,
        'copilot-search': 0.75
      },
      variantWeights: {
        raw: 1.0,
        'semantic-rewrite': 0.98,
        corrected: 0.95,
        translation: 0.8,
        'keyword-fallback': 0.7
      },
      boosts: {
        exactQueryLanguage: 0.1,
        variantQueryLanguage: 0.06,
        titleMatch: 0.08,
        currentSite: 0.05,
        retrievalNative: 0.1,
        sharePointNative: 0.08
      }
    }
  }
};

export function getRuntimeTuningConfig(): IRuntimeTuningConfig {
  if (typeof window === 'undefined' || !window.__GRIMOIRE_RUNTIME_TUNING__) {
    return DEFAULT_RUNTIME_TUNING_CONFIG;
  }

  const overrides = window.__GRIMOIRE_RUNTIME_TUNING__;

  return {
    nano: {
      ...DEFAULT_RUNTIME_TUNING_CONFIG.nano,
      ...(overrides.nano || {})
    },
    search: {
      ...DEFAULT_RUNTIME_TUNING_CONFIG.search,
      ...(overrides.search || {}),
      plannerConfidence: {
        ...DEFAULT_RUNTIME_TUNING_CONFIG.search.plannerConfidence,
        ...(overrides.search?.plannerConfidence || {})
      },
      fusion: {
        ...DEFAULT_RUNTIME_TUNING_CONFIG.search.fusion,
        ...(overrides.search?.fusion || {}),
        sourceWeights: {
          ...DEFAULT_RUNTIME_TUNING_CONFIG.search.fusion.sourceWeights,
          ...(overrides.search?.fusion?.sourceWeights || {})
        },
        variantWeights: {
          ...DEFAULT_RUNTIME_TUNING_CONFIG.search.fusion.variantWeights,
          ...(overrides.search?.fusion?.variantWeights || {})
        },
        boosts: {
          ...DEFAULT_RUNTIME_TUNING_CONFIG.search.fusion.boosts,
          ...(overrides.search?.fusion?.boosts || {})
        }
      }
    }
  };
}
