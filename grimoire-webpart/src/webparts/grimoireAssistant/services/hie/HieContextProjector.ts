import type { BlockTracker } from './BlockTracker';
import { HieInteractionFormatter } from './HieInteractionFormatter';
import type {
  BlockInteractionAction,
  IBlockInteraction,
  IHieArtifactRecord,
  IHieDerivedState,
  IHieEvent,
  IHieTaskContext
} from './HIETypes';
import type { IHiePromptMessage } from './HiePromptProtocol';

function clip(value: string, maxChars: number): string {
  const trimmed = value.trim();
  if (trimmed.length <= maxChars) return trimmed;
  return `${trimmed.slice(0, Math.max(0, maxChars - 1)).trimEnd()}...`;
}

function describeSelectedItems(taskContext: IHieTaskContext): string | undefined {
  if (!taskContext.selectedItems || taskContext.selectedItems.length === 0) {
    return undefined;
  }

  const parts = taskContext.selectedItems
    .slice(0, 4)
    .map((item) => {
      const indexLabel = item.index !== undefined ? `${item.index}) ` : '';
      return `${indexLabel}${item.title || item.url || item.kind || item.itemType || 'item'}`;
    });

  return parts.length > 0 ? `Selected items: ${parts.join(', ')}` : undefined;
}

function describeTaskContext(taskContext?: IHieTaskContext): string | undefined {
  if (!taskContext) return undefined;

  switch (taskContext.kind) {
    case 'search': {
      const title = taskContext.sourceBlockTitle || 'current search results';
      return `The latest task is a search centered on "${title}".`;
    }
    case 'focus': {
      const title = taskContext.sourceBlockTitle || 'current results';
      return [
        `The user focused items from "${title}".`,
        describeSelectedItems(taskContext)
      ].filter(Boolean).join(' ');
    }
    case 'summarize': {
      const title = taskContext.sourceBlockTitle || 'current results';
      const derivedTitle = taskContext.derivedBlockTitle;
      return [
        derivedTitle
          ? `The latest derived summary artifact is "${derivedTitle}", created from "${title}".`
          : `The user asked for a derived summary from "${title}".`,
        describeSelectedItems(taskContext)
      ].filter(Boolean).join(' ');
    }
    case 'chat-about': {
      const title = taskContext.sourceBlockTitle || 'current results';
      return [
        `The conversation is centered on "${title}".`,
        describeSelectedItems(taskContext)
      ].filter(Boolean).join(' ');
    }
    case 'select': {
      const title = taskContext.sourceBlockTitle || 'current options';
      if (!taskContext.selectedItems || taskContext.selectedItems.length === 0) {
        return taskContext.sourceBlockTitle
          ? `There is no current action-panel selection in "${title}".`
          : 'There is no current action-panel selection.';
      }
      return [
        `The user selected options from "${title}".`,
        describeSelectedItems(taskContext)
      ].filter(Boolean).join(' ');
    }
    case 'recap': {
      const recapTitle = taskContext.derivedBlockTitle || taskContext.sourceBlockTitle || 'current results';
      return `The latest summary artifact is "${recapTitle}". Treat that recap as the current high-level summary unless the user narrows away from it.`;
    }
    case 'form': {
      const preset = taskContext.formPreset || 'compose';
      if (taskContext.formStatus === 'failed') {
        return `The ${preset} form submission failed.`;
      }
      if (taskContext.formStatus === 'submitted') {
        return `The ${preset} form was submitted.`;
      }
      if (taskContext.formStatus === 'cancelled') {
        return `The ${preset} form was cancelled.`;
      }
      if (taskContext.formStatus === 'dismissed') {
        return `The ${preset} form was closed and is no longer visible in the action panel.`;
      }
      return `The user is working in a ${preset} form in the action panel.`;
    }
    default:
      return taskContext.sourceBlockTitle
        ? `The latest task is centered on "${taskContext.sourceBlockTitle}".`
        : undefined;
  }
}

function getLatestOpenForm(artifacts: Record<string, IHieArtifactRecord>): IHieArtifactRecord | undefined {
  return Object.values(artifacts)
    .filter((artifact) => (artifact.artifactKind === 'form' || artifact.artifactKind === 'share') && artifact.status === 'opened')
    .sort((left, right) => right.updatedAt - left.updatedAt)[0];
}

function deterministicClamp(parts: string[], maxChars: number): string {
  const output: string[] = [];
  let remaining = maxChars;

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    if (!part) continue;
    const segment = output.length === 0 ? part : ` | ${part}`;
    if (segment.length <= remaining) {
      output.push(output.length === 0 ? part : part);
      remaining -= segment.length;
      continue;
    }

    if (remaining > 12) {
      const clipped = clip(part, output.length === 0 ? remaining : remaining - 3);
      if (clipped) {
        output.push(clipped);
      }
    }
    break;
  }

  return output.join(' | ');
}

export class HieContextProjector {
  public buildCurrentStateReminder(
    visualSummary: string,
    derivedState: IHieDerivedState,
    maxChars: number
  ): string {
    const parts: string[] = [];
    if (visualSummary.trim()) {
      parts.push(visualSummary.trim());
    }

    const taskPart = describeTaskContext(derivedState.taskContext);
    if (taskPart) {
      parts.push(taskPart);
    }

    const openForm = getLatestOpenForm(derivedState.artifacts);
    if (openForm?.title) {
      parts.push(`Current form: ${openForm.title}.`);
    }

    return deterministicClamp(parts, maxChars);
  }

  public projectEvent(
    event: IHieEvent,
    derivedState: IHieDerivedState,
    tracker?: BlockTracker
  ): IHiePromptMessage | undefined {
    if (event.eventName.startsWith('block.interaction.')) {
      return this.projectBlockInteractionEvent(event, tracker);
    }

    switch (event.eventName) {
      case 'task.focused': {
        const taskContext = describeTaskContext(derivedState.taskContext);
        return taskContext ? { kind: 'task', body: taskContext } : undefined;
      }
      case 'task.selection.updated': {
        const taskContext = describeTaskContext(derivedState.taskContext);
        return taskContext ? { kind: 'task', body: taskContext } : undefined;
      }
      case 'task.recap.requested': {
        const sourceTitle = typeof event.payload.sourceBlockTitle === 'string' ? event.payload.sourceBlockTitle : 'current results';
        return { kind: 'task', body: `Generating a recap for "${sourceTitle}" in the action panel.` };
      }
      case 'artifact.recap.ready': {
        const recapTitle = typeof event.payload.derivedBlockTitle === 'string'
          ? event.payload.derivedBlockTitle
          : (typeof event.payload.sourceBlockTitle === 'string' ? event.payload.sourceBlockTitle : 'current results');
        return { kind: 'task', body: `Recap ready. The latest summary artifact is "${recapTitle}".` };
      }
      case 'artifact.recap.failed': {
        const sourceTitle = typeof event.payload.sourceBlockTitle === 'string' ? event.payload.sourceBlockTitle : 'current results';
        return { kind: 'task', body: `Recap failed for "${sourceTitle}". Keep using the source block as the current context.` };
      }
      case 'form.opened': {
        const preset = typeof event.payload.preset === 'string' ? event.payload.preset : 'compose';
        const sourceTitle = typeof event.payload.sourceBlockTitle === 'string' ? event.payload.sourceBlockTitle : undefined;
        const suffix = sourceTitle ? ` It was opened from "${sourceTitle}".` : '';
        return { kind: 'task', body: `A ${preset} form is open in the action panel.${suffix}` };
      }
      case 'form.submitted':
      case 'form.dismissed':
      case 'form.cancelled': {
        const preset = typeof event.payload.preset === 'string' ? event.payload.preset : 'compose';
        const verb = event.eventName === 'form.submitted'
          ? 'submitted'
          : (event.eventName === 'form.dismissed' ? 'closed' : 'cancelled');
        return { kind: 'task', body: `The ${preset} form was ${verb}.` };
      }
      case 'block.removed': {
        const blockTitle = typeof event.payload.blockTitle === 'string'
          ? event.payload.blockTitle
          : (typeof event.payload.blockType === 'string' ? event.payload.blockType : 'panel');
        return { kind: 'task', body: `The "${blockTitle}" panel was dismissed and is no longer visible.` };
      }
      default:
        return undefined;
    }
  }

  private projectBlockInteractionEvent(event: IHieEvent, tracker?: BlockTracker): IHiePromptMessage | undefined {
    if (!tracker || !event.blockId || !event.blockType) {
      return undefined;
    }

    const actionName = event.eventName.slice('block.interaction.'.length);
    if (!actionName) {
      return undefined;
    }

    const interaction: IBlockInteraction = {
      blockId: event.blockId,
      blockType: event.blockType,
      action: actionName as BlockInteractionAction,
      payload: event.payload,
      timestamp: event.timestamp,
      source: event.source,
      surface: event.surface,
      eventName: event.eventName,
      exposurePolicy: event.exposurePolicy,
      correlationId: event.correlationId
    };

    return new HieInteractionFormatter(tracker).format(interaction);
  }
}
