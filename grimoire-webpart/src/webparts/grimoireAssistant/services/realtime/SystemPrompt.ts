/**
 * SystemPrompt
 * M365 assistant persona for GPT Realtime voice sessions.
 * Includes personality modifiers, tool usage guidance, and expression control.
 */

import { PersonalityMode, PERSONALITIES } from '../avatar/PersonalityEngine';
import type { VisageMode } from '../avatar/FaceTemplateData';
import { getAvatarPersonaConfig } from '../avatar/AvatarPersonaCatalog';
import { M365_MCP_CATALOG, getCatalogSummaryForPrompt, getTotalToolCount } from '../../models/McpServerCatalog';
import type { IUserContext } from '../context/ContextService';
import {
  PERSONALITY_PROMPT_MODIFIERS,
  SYSTEM_PROMPT_ROLE_SECTION,
  buildSystemPromptAvailableToolsSection,
  getDefaultConversationGuidelineLines,
  getDefaultFormCompositionProtocolLines,
  getDefaultVisualContextProtocolLines,
  getSystemPromptEntitySpecificGuidanceLines,
  getSystemPromptErrorRecoveryLines,
  getSystemPromptExpressionGuideLines,
  getSystemPromptPersonalContextLines,
  getSystemPromptRoutingGuidanceLines,
  getSystemPromptSearchPlanningLines
} from '../../config/promptCatalog';
import { getPromptRuntimeConfig, getSearchQueryLanguageRule } from './PromptRuntimeConfig';
import { buildEnterpriseIntentRoutingGuidance } from './IntentRoutingPolicy';

/** Prompt configuration flags for conditional capabilities rendering. */
export interface IPromptConfig {
  mcpEnvironmentId?: string;
  hasGraphAccess?: boolean;
  avatarEnabled?: boolean;
  conversationLanguage?: string;
}

function buildSection(title: string, lines: string[]): string {
  return [title, '', ...lines].join('\n');
}

function buildM365CatalogSection(config?: IPromptConfig): string {
  return [
    '## M365 MCP Server Catalog',
    '',
    `You have access to ${M365_MCP_CATALOG.length} Agent 365 MCP servers with ${getTotalToolCount()} tools via \`use_m365_capability\`:`,
    '',
    getCatalogSummaryForPrompt(config)
  ].join('\n');
}

function buildRoutingGuidanceSection(): string {
  return [
    '### Routing Guidance',
    buildEnterpriseIntentRoutingGuidance(),
    ...getSystemPromptRoutingGuidanceLines()
  ].join('\n');
}

function buildUserSection(ctx: IUserContext): string {
  const lines: string[] = ['## Your User', ''];
  lines.push(`- Name: ${ctx.displayName}`);
  if (ctx.email) lines.push(`- Email: ${ctx.email}`);
  if (ctx.jobTitle) lines.push(`- Job Title: ${ctx.jobTitle}`);
  if (ctx.department) lines.push(`- Department: ${ctx.department}`);
  if (ctx.manager) lines.push(`- Manager: ${ctx.manager}`);
  if (ctx.resolvedLanguage && ctx.resolvedLanguage !== 'en') {
    lines.push(`- Preferred language (profile metadata): ${ctx.resolvedLanguage}`);
  }
  if (ctx.currentWebTitle) {
    const siteLabel = ctx.currentWebUrl ? `${ctx.currentWebTitle} (${ctx.currentWebUrl})` : ctx.currentWebTitle;
    lines.push(`- Current site: ${siteLabel}`);
  }
  if (ctx.currentListTitle) {
    lines.push(`- Viewing list: ${ctx.currentListTitle}`);
  }
  return lines.join('\n');
}

function buildAvatarPersonaSection(visage: VisageMode): string {
  const persona = getAvatarPersonaConfig(visage);
  return buildSection('## Avatar Persona', [
    `Selected avatar identity: ${persona.title}.`,
    persona.promptModifier,
    'This avatar layer is additive to the selected personality and only affects tone, self-framing, and light flavor.',
    `When a first natural self-introduction is appropriate, or when directly asked who or what you are, use this identity line: "${persona.introIdentityLine}"`,
    'Do not repeat the identity line in routine answers.',
    'Do not let avatar flavor change tool choice, routing, security posture, confidence thresholds, or factual behavior.',
    'Do not claim literal authority, sentience, or powers beyond your actual assistant capabilities.'
  ]);
}

function buildConversationLanguageSection(language?: string): string {
  if (!language) return '';
  return buildSection('## Conversation Language', [
    `Current conversation language: ${language}.`,
    'Use this language for all responses until the user explicitly switches to another language.',
    'Do not snap back to English after tool calls, follow-up questions, or action-panel updates.'
  ]);
}

/**
 * Build the system prompt for a Grimoire voice session.
 */
export function getSystemPrompt(
  personality: PersonalityMode,
  visage: VisageMode = 'classic',
  userContext?: IUserContext,
  config?: IPromptConfig
): string {
  const runtimeCfg = getPromptRuntimeConfig();
  const personalityModifier = runtimeCfg.personalityPromptOverrides?.[personality]
    || PERSONALITY_PROMPT_MODIFIERS[personality]
    || PERSONALITIES[personality].promptModifier;
  const searchLanguageRule = getSearchQueryLanguageRule();
  const expressionGuideLines = getSystemPromptExpressionGuideLines(config);

  const sections = [
    runtimeCfg.systemPromptPrefix ? `${runtimeCfg.systemPromptPrefix}\n${personalityModifier}` : personalityModifier,
    config?.avatarEnabled === false ? '' : buildAvatarPersonaSection(visage),
    buildConversationLanguageSection(config?.conversationLanguage),
    userContext ? buildUserSection(userContext) : '',
    SYSTEM_PROMPT_ROLE_SECTION,
    buildSystemPromptAvailableToolsSection(config),
    buildM365CatalogSection(config),
    buildRoutingGuidanceSection(),
    buildSection('### Entity-Specific Guidance', getSystemPromptEntitySpecificGuidanceLines()),
    buildSection(
      '## Conversation Guidelines',
      runtimeCfg.conversationGuidelinesOverride || getDefaultConversationGuidelineLines(searchLanguageRule, config)
    ),
    expressionGuideLines.length > 0 ? buildSection('## Expression Guide', expressionGuideLines) : '',
    buildSection(
      '## Visual Context Protocol',
      runtimeCfg.visualContextProtocolOverride || getDefaultVisualContextProtocolLines()
    ),
    buildSection(
      '## Form Composition Protocol',
      runtimeCfg.formCompositionProtocolOverride || getDefaultFormCompositionProtocolLines()
    ),
    buildSection('## Search Planning', getSystemPromptSearchPlanningLines()),
    buildSection('## Error Recovery', getSystemPromptErrorRecoveryLines()),
    buildSection('## Personal Context', getSystemPromptPersonalContextLines())
  ].filter((section) => !!section);

  return sections.join('\n\n');
}
