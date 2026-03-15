/**
 * PromptRuntimeConfig
 * Runtime prompt/tooling overrides loaded from the host page (if provided).
 * Defaults stay hardcoded and safe when no override is configured.
 */

import type { PersonalityMode } from '../avatar/PersonalityEngine';

export interface IPromptRuntimeConfig {
  systemPromptPrefix?: string;
  conversationGuidelinesOverride?: string[];
  visualContextProtocolOverride?: string[];
  formCompositionProtocolOverride?: string[];
  toolDescriptionOverrides?: Record<string, string>;
  personalityPromptOverrides?: Partial<Record<PersonalityMode, string>>;
  searchQueryLanguageRule?: string;
}

declare global {
  interface Window {
    __GRIMOIRE_PROMPT_CONFIG__?: IPromptRuntimeConfig;
  }
}

/**
 * Shared language rule used in both tool definitions and system prompt.
 */
export const DEFAULT_SEARCH_QUERY_LANGUAGE_RULE =
  'Use the SAME language as the user query. Do NOT force translation to English unless you retry after sparse results.';

export function getPromptRuntimeConfig(): IPromptRuntimeConfig {
  if (typeof window === 'undefined' || !window.__GRIMOIRE_PROMPT_CONFIG__) {
    return {};
  }
  return window.__GRIMOIRE_PROMPT_CONFIG__;
}

export function getSearchQueryLanguageRule(): string {
  const cfg = getPromptRuntimeConfig();
  return cfg.searchQueryLanguageRule || DEFAULT_SEARCH_QUERY_LANGUAGE_RULE;
}

export function getToolDescriptionOverride(toolName: string, fallback: string): string {
  const cfg = getPromptRuntimeConfig();
  if (cfg.toolDescriptionOverrides && cfg.toolDescriptionOverrides[toolName]) {
    return cfg.toolDescriptionOverrides[toolName];
  }
  return fallback;
}
