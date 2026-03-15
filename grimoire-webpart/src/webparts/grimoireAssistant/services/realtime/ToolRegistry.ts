/**
 * ToolRegistry
 * Shared GPT Realtime function catalog.
 * Organized in 10 categories, sourced from the shared tool catalog.
 */

import { getToolCatalog, type IToolDefinition } from '../../config/toolCatalog';
import { getSearchQueryLanguageRule, getToolDescriptionOverride } from './PromptRuntimeConfig';

export type { IToolDefinition } from '../../config/toolCatalog';

export interface IToolRegistryOptions {
  avatarEnabled?: boolean;
}

/**
 * Get all tool definitions for GPT Realtime.
 */
export function getTools(options?: IToolRegistryOptions): IToolDefinition[] {
  return getToolCatalog(getSearchQueryLanguageRule())
    .filter((tool) => options?.avatarEnabled !== false || tool.name !== 'set_expression')
    .map((tool) => ({
      type: tool.type,
      name: tool.name,
      description: getToolDescriptionOverride(tool.name, tool.description),
      parameters: tool.parameters
    }));
}
