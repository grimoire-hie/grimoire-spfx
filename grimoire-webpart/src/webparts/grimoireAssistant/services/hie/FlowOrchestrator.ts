/**
 * FlowOrchestrator — Rule-based multi-step flow pattern detection.
 *
 * Detects when the user is in a multi-step interaction pattern and provides
 * richer context to the LLM. Does NOT force flows — the LLM stays in control.
 *
 * Patterns:
 *   search-then-drill      — search → click result → fetch details
 *   browse-then-open       — browse library → click folder/file → navigate
 *   confirm-before-action  — destructive tool → confirmation dialog → execute/abort
 *   select-then-act        — selection list → user selects → proceed
 */

import {
  IFlowDefinition,
  IFlowInstance,
  IBlockInteraction,
  IHIEConfig,
  DESTRUCTIVE_TOOL_NAMES
} from './HIETypes';

// ─── Flow Definitions ───────────────────────────────────────────

const FLOW_DEFINITIONS: IFlowDefinition[] = [
  {
    id: 'search-then-drill',
    name: 'Search and Drill Down',
    steps: [
      { name: 'search', description: 'Search results displayed' },
      { name: 'select', description: 'User clicked a result' },
      { name: 'detail', description: 'Detailed view shown' }
    ],
    triggerTools: ['search_sharepoint', 'search_people', 'search_sites']
  },
  {
    id: 'browse-then-open',
    name: 'Browse and Navigate',
    steps: [
      { name: 'browse', description: 'Document library displayed' },
      { name: 'navigate', description: 'User clicked folder or file' },
      { name: 'view', description: 'Content displayed' }
    ],
    triggerTools: ['browse_document_library']
  },
  {
    id: 'confirm-before-action',
    name: 'Confirm Before Action',
    steps: [
      { name: 'prompt', description: 'Confirmation dialog shown' },
      { name: 'respond', description: 'User confirmed or cancelled' }
    ],
    triggerTools: ['ask_confirmation']
  },
  {
    id: 'select-then-act',
    name: 'Select and Proceed',
    steps: [
      { name: 'present', description: 'Selection list displayed' },
      { name: 'choose', description: 'User made selection' }
    ],
    triggerTools: ['show_selection_list']
  },
  {
    id: 'compose-then-submit',
    name: 'Compose and Submit',
    steps: [
      { name: 'compose', description: 'Form displayed for user input' },
      { name: 'submit', description: 'User submitted or cancelled the form' }
    ],
    triggerTools: ['show_compose_form']
  }
];

export class FlowOrchestrator {
  private config: IHIEConfig;
  private activeFlows: Map<string, IFlowInstance> = new Map();
  private flowIdCounter: number = 0;

  constructor(config: IHIEConfig) {
    this.config = config;
  }

  /**
   * Check if a tool invocation should start a new flow.
   * Returns a flow context message if a flow is detected, or undefined.
   */
  public onToolInvoked(toolName: string, blockId: string): string | undefined {
    if (!this.config.flowOrchestrationEnabled) return undefined;

    // Check for destructive tool — inject guidance
    if (this.isDestructiveTool(toolName)) {
      return `Destructive operation detected (${toolName}). A confirmation dialog should be shown before proceeding.`;
    }

    // Check if this tool triggers a flow
    for (let i = 0; i < FLOW_DEFINITIONS.length; i++) {
      const def = FLOW_DEFINITIONS[i];
      if (def.triggerTools.indexOf(toolName) !== -1) {
        return this.startFlow(def, blockId);
      }
    }

    return undefined;
  }

  /**
   * Process a block interaction within an active flow.
   * Returns a flow context message if the interaction advances a flow.
   */
  public onInteraction(interaction: IBlockInteraction): string | undefined {
    if (!this.config.flowOrchestrationEnabled) return undefined;

    let result: string | undefined;

    // Check if this interaction belongs to any active flow
    this.activeFlows.forEach((flow, flowKey) => {
      if (flow.blockIds.indexOf(interaction.blockId) === -1) return;

      // Advance the flow
      if (flow.currentStep < flow.definition.steps.length - 1) {
        flow.currentStep++;
        const step = flow.definition.steps[flow.currentStep];
        result = `User is in step ${flow.currentStep + 1} of ${flow.definition.name}: ${step.description}`;
      } else {
        // Flow complete
        this.activeFlows.delete(flowKey);
        result = `${flow.definition.name} flow completed`;
      }
    });

    return result;
  }

  /**
   * Associate a new block with an existing flow (e.g. detail view after clicking a search result).
   */
  public addBlockToFlow(blockId: string, toolName: string): void {
    if (!this.config.flowOrchestrationEnabled) return;

    // Find an active flow that this tool could advance
    this.activeFlows.forEach((flow) => {
      if (flow.currentStep > 0 && flow.currentStep < flow.definition.steps.length - 1) {
        flow.blockIds.push(blockId);
        flow.currentStep++;
      }
    });
  }

  /**
   * Get a summary of all active flows for system prompt enrichment.
   */
  public getActiveFlowSummary(): string {
    if (this.activeFlows.size === 0) return '';

    const parts: string[] = [];
    this.activeFlows.forEach((flow) => {
      const step = flow.definition.steps[flow.currentStep];
      parts.push(`${flow.definition.name}: step ${flow.currentStep + 1}/${flow.definition.steps.length} (${step.description})`);
    });

    return parts.join('; ');
  }

  /**
   * Check if a tool name is known to be destructive.
   */
  public isDestructiveTool(toolName: string): boolean {
    return DESTRUCTIVE_TOOL_NAMES.has(toolName);
  }

  public reset(): void {
    this.activeFlows.clear();
    this.flowIdCounter = 0;
  }

  // ─── Private ──────────────────────────────────────────────────

  private startFlow(definition: IFlowDefinition, blockId: string): string {
    this.flowIdCounter++;
    const key = `${definition.id}-${this.flowIdCounter}`;

    const flow: IFlowInstance = {
      definition,
      currentStep: 0,
      context: {},
      blockIds: [blockId],
      startedAt: Date.now()
    };

    this.activeFlows.set(key, flow);

    const step = definition.steps[0];
    return `${definition.name} started: ${step.description}`;
  }
}
