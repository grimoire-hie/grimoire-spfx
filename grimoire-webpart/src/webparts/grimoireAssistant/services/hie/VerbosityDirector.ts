/**
 * VerbosityDirector — Computes how verbose the LLM should be based on
 * the current visual state (number/type of data blocks displayed).
 *
 * When rich data is already visible in the action panel the LLM should
 * speak less; when an error is displayed it should explain more.
 */

import { BlockTracker } from './BlockTracker';
import { IHIEConfig, VerbosityLevel } from './HIETypes';

/** Block types that contain dense data the user can read visually */
const DATA_BLOCK_TYPES: ReadonlySet<string> = new Set([
  'search-results',
  'document-library',
  'list-items',
  'file-preview',
  'permissions-view',
  'activity-feed'
]);

export class VerbosityDirector {
  private config: IHIEConfig;
  private tracker: BlockTracker;
  private lastInteractionTime: number = 0;

  constructor(config: IHIEConfig, tracker: BlockTracker) {
    this.config = config;
    this.tracker = tracker;
  }

  /** Record that the user interacted with the UI */
  public recordInteraction(): void {
    this.lastInteractionTime = Date.now();
  }

  /**
   * Compute verbosity hint based on current block state.
   * Returns undefined when no hint is needed (normal verbosity).
   */
  public getVerbosityHint(): { level: VerbosityLevel; hint: string } | undefined {
    if (!this.config.contextInjectionEnabled) return undefined;

    // Recent interaction (<5s) → user is clicking around the UI
    if (Date.now() - this.lastInteractionTime < 5000) {
      return { level: 'minimal', hint: 'User is interacting with the UI.' };
    }

    let dataBlockCount = 0;
    let totalItems = 0;
    let hasError = false;

    this.tracker.getAllReferences().forEach((block) => {
      const tracked = this.tracker.get(block.blockId);
      if (!tracked || tracked.state === 'dismissed') return;
      if (DATA_BLOCK_TYPES.has(tracked.type)) {
        dataBlockCount++;
        totalItems += tracked.itemCount;
      }
      if (tracked.type === 'error') {
        hasError = true;
      }
    });

    // Also check for error blocks that have no references
    // (errors don't produce references, check via active summary hack)
    if (!hasError) {
      const summary = this.tracker.getActiveSummary();
      if (summary.indexOf('Error:') !== -1) {
        hasError = true;
      }
    }

    if (hasError) {
      return { level: 'detailed', hint: 'Error displayed. Explain the issue verbally.' };
    }

    if (dataBlockCount >= 3 || totalItems > 20) {
      return { level: 'minimal', hint: 'Keep response to 1 sentence. Rich data is displayed visually.' };
    }

    if (dataBlockCount >= 1) {
      return { level: 'brief', hint: 'Be concise (1-2 sentences). Data is visible in action panel.' };
    }

    // No blocks → normal (return undefined, no injection needed)
    return undefined;
  }

  public reset(): void {
    this.lastInteractionTime = 0;
  }
}
