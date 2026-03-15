/**
 * ActionPanel
 * Center column — renders a scrolling stack of UI blocks.
 * Blocks are pushed by tool handlers and rendered from the block registry.
 * Uses SP theme colors from the store for adaptive theming.
 */

import * as React from 'react';
import { DefaultButton, IconButton, Text } from '@fluentui/react';
import { shallow } from 'zustand/shallow';
import * as strings from 'GrimoireAssistantWebPartStrings';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { createBlock } from '../../models/IBlock';
import type { FormPresetId, IBlock, IFormData, IInfoCardData } from '../../models/IBlock';
import { getBlockComponent } from './blocks/index';
import { FormBlock } from './blocks/FormBlock';
import { executeFormSubmission } from '../../services/forms/FormSubmissionService';
import { getFormPreset } from '../../services/forms/FormPresets';
import { hybridInteractionEngine } from '../../services/hie/HybridInteractionEngine';
import { emitBlockInteraction } from './interactionSchemas';
import {
  getSelectionCandidates,
  isActionableBlockType,
  type ISelectionCandidate
} from './selectionHelpers';
import { logService } from '../../services/logging/LogService';
import {
  evaluateHeaderAction,
  getEligibleCandidatesForAction,
  buildFocusGuardrailSummary
} from './actionPolicies';
import { BlockRecapService, canRecapBlock, getRecapOriginTool } from '../../services/recap/BlockRecapService';
import { getActionPanelCloseCopy, resolveActionPanelCloseAction } from '../layout/closeBehavior';
import { executeShareSubmission, isInternalSharePreset } from '../../services/sharing/ShareSubmissionService';
import { SessionShareFormatter, hasShareableSessionContent } from '../../services/sharing/SessionShareFormatter';
import { deriveMcpTargetContextFromUnknown } from '../../services/mcp/McpTargetContext';

// ─── Form submission adapter (stable references) ─────────────────
// These are module-level functions so they're never recreated, avoiding
// unnecessary FormBlock re-renders.

async function handleFormSubmit(
  formData: IFormData,
  fieldValues: Record<string, string>,
  emailTags: Record<string, string[]>
): Promise<{ success: boolean; message: string }> {
  const state = useGrimoireStore.getState();
  const sourceContext = hybridInteractionEngine.captureCurrentSourceContext();
  if (isInternalSharePreset(formData.preset)) {
    return executeShareSubmission(formData, fieldValues, emailTags, {
      aadHttpClient: state.aadHttpClient,
      proxyConfig: state.proxyConfig,
      getToken: state.getToken,
      mcpEnvironmentId: state.mcpEnvironmentId,
      userContext: state.userContext,
      mcpConnections: state.mcpConnections,
      pushBlock: state.pushBlock,
      updateBlock: state.updateBlock,
      removeBlock: state.removeBlock,
      clearBlocks: state.clearBlocks,
      setExpression: state.setExpression,
      setActivityStatus: state.setActivityStatus
    }, sourceContext);
  }

  return executeFormSubmission(formData, fieldValues, emailTags, {
    aadHttpClient: state.aadHttpClient,
    proxyConfig: state.proxyConfig,
    getToken: state.getToken,
    mcpEnvironmentId: state.mcpEnvironmentId,
    userContext: state.userContext,
    mcpConnections: state.mcpConnections,
    pushBlock: state.pushBlock,
    updateBlock: state.updateBlock,
    removeBlock: state.removeBlock,
    clearBlocks: state.clearBlocks,
    setExpression: state.setExpression,
    setActivityStatus: state.setActivityStatus
  }, sourceContext);
}

function handleFormUpdateBlock(blockId: string, updates: { data: IFormData }): void {
  useGrimoireStore.getState().updateBlock(blockId, updates);
}

/**
 * Renders the block body using the registry, with fallback for unregistered types.
 */
const BlockBody: React.FC<{ block: IBlock }> = ({ block }) => {
  // Form blocks receive submission callbacks from ActionPanel
  if (block.type === 'form') {
    return React.createElement(FormBlock, {
      data: block.data as IFormData,
      blockId: block.id,
      onSubmit: handleFormSubmit,
      onUpdateBlock: handleFormUpdateBlock
    });
  }
  const Component = getBlockComponent(block.type);
  if (Component) {
    return React.createElement(Component, { data: block.data, blockId: block.id, renderHints: block.renderHints });
  }
  // Fallback for unregistered block types
  return (
    <Text variant="small" styles={{ root: { opacity: 0.5 } }}>
      [{block.type}] block
    </Text>
  );
};

/**
 * Renders a single block card with header and typed body.
 * When `isFresh` the card pulses with a subtle shimmer until the user mouses over it.
 */
const BlockCard: React.FC<{
  block: IBlock;
  isFresh: boolean;
  onAcknowledge: () => void;
  onDismiss: () => void;
  cardBg: string;
  cardBorder: string;
  textColor: string;
  subtextColor: string;
}> = ({ block, isFresh, onAcknowledge, onDismiss, cardBg, cardBorder, textColor, subtextColor }) => {
  const timeStr = block.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  // Skip shimmer on large blocks — the animation + heavy DOM causes UI freeze/flash
  const dataContent = (block.data as { content?: string }).content;
  const isLargeBlock = typeof dataContent === 'string' && dataContent.length > 3000;
  const showShimmer = isFresh && !isLargeBlock;

  // Auto-acknowledge fresh blocks after 4s (prevents infinite shimmer in voice-only usage)
  // Immediately acknowledge large blocks to avoid any shimmer attempt
  React.useEffect(() => {
    if (!isFresh) return;
    const timer = setTimeout(onAcknowledge, isLargeBlock ? 0 : 4000);
    return () => clearTimeout(timer);
  }, [isFresh]);

  return (
    <div
      onMouseEnter={isFresh ? onAcknowledge : undefined}
      style={{
        backgroundColor: cardBg,
        borderRadius: 8,
        padding: 16,
        border: `1px solid ${cardBorder}`,
        position: 'relative',
        animation: showShimmer ? 'grimShimmer 2s ease-in-out infinite' : undefined,
        transition: 'border-color 0.3s ease, box-shadow 0.3s ease'
      }}
    >
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        marginBottom: 8
      }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: textColor }}>{block.title}</span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 11, color: subtextColor }}>{timeStr}</span>
          {block.dismissible && (
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              ariaLabel="Dismiss"
              onClick={onDismiss}
              styles={{
                root: { width: 20, height: 20, color: subtextColor },
                rootHovered: { color: textColor, backgroundColor: 'transparent' }
              }}
            />
          )}
        </div>
      </div>
      <BlockBody block={block} />
    </div>
  );
};

// ─── Shared style injection helper ──────────────────────────────
function injectStyleOnce(id: string, css: string): void {
  if (typeof document === 'undefined') return;
  if (document.getElementById(id)) return;
  const style = document.createElement('style');
  style.id = id;
  style.textContent = css;
  document.head.appendChild(style);
}

// ─── Shimmer animation ─────────────────────────────────────────
const SHIMMER_STYLE_ID = 'grimoire-shimmer-style';
const SHIMMER_CSS = `@keyframes grimShimmer {
  0%, 100% { border-color: rgba(120, 180, 255, 0.15); box-shadow: none; }
  50% { border-color: rgba(120, 180, 255, 0.5); box-shadow: 0 0 12px rgba(120, 180, 255, 0.15); }
}`;

// ─── Scrollbar styling ─────────────────────────────────────────
const SCROLL_STYLE_ID = 'grimoire-scroll-style';
const SCROLL_CSS = [
  '.grimoire-scroll::-webkit-scrollbar { width: 6px; }',
  '.grimoire-scroll::-webkit-scrollbar-track { background: transparent; }',
  '.grimoire-scroll::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.15); border-radius: 3px; }',
  '.grimoire-scroll::-webkit-scrollbar-thumb:hover { background: rgba(0,0,0,0.25); }',
  '.grimoire-scroll-dark::-webkit-scrollbar { width: 6px; }',
  '.grimoire-scroll-dark::-webkit-scrollbar-track { background: transparent; }',
  '.grimoire-scroll-dark::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.15); border-radius: 3px; }',
  '.grimoire-scroll-dark::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.25); }'
].join('\n');

function actionButtonStyle(enabled: boolean): React.CSSProperties {
  return {
    padding: '4px 9px',
    borderRadius: 4,
    border: '1px solid rgba(0, 100, 180, 0.2)',
    background: enabled ? 'rgba(0, 100, 180, 0.1)' : 'rgba(0, 0, 0, 0.03)',
    color: enabled ? '#0064b4' : 'rgba(0, 0, 0, 0.35)',
    fontSize: 11,
    fontWeight: 600,
    cursor: enabled ? 'pointer' : 'default'
  };
}

function buildPrefilledFormData(
  presetId: FormPresetId,
  description: string,
  defaults: Record<string, string>,
  staticArgs?: Record<string, unknown>
): IFormData {
  const preset = getFormPreset(presetId);
  return {
    kind: 'form',
    preset: presetId,
    description,
    fields: preset.fields.map((field) => (
      defaults[field.key] !== undefined
        ? { ...field, defaultValue: defaults[field.key] }
        : field
    )),
    submissionTarget: {
      ...preset.submissionTarget,
      staticArgs: {
        ...preset.submissionTarget.staticArgs,
        ...(staticArgs || {})
      }
    },
    status: 'editing'
  };
}

function buildSelectionSnapshotPayload(
  activeActionBlock: IBlock | undefined,
  selectedCandidates: ISelectionCandidate[]
): Record<string, unknown> {
  const selectedItems = selectedCandidates.map((candidate) => ({
    index: candidate.index,
    title: candidate.title,
    url: candidate.url,
    kind: candidate.kind,
    itemType: candidate.itemType,
    targetContext: deriveMcpTargetContextFromUnknown(candidate.payload, 'hie-selection')
  }));

  return {
    sourceBlockId: activeActionBlock?.id,
    sourceBlockType: activeActionBlock?.type,
    sourceBlockTitle: activeActionBlock?.title,
    selectedCount: selectedItems.length,
    selectedItems,
    selectionCleared: selectedItems.length === 0
  };
}

export interface IActionPanelProps {
  isSettingsOpen: boolean;
  onCloseApp?: () => void;
  onCloseSettings?: () => void;
}

export const ActionPanel: React.FC<IActionPanelProps> = ({ isSettingsOpen, onCloseApp, onCloseSettings }) => {
  const {
    blocks,
    transcript,
    pushBlock,
    removeBlock,
    freshBlockIds,
    acknowledgeBlock,
    activeActionBlockId,
    selectedActionIndices,
    setActionSelection,
    setActiveActionBlock,
    setFocusedContext,
    triggerAvatarActionCue,
    insertBlockAfter,
    updateBlock,
    proxyConfig
  } = useGrimoireStore((s) => ({
    blocks: s.blocks,
    transcript: s.transcript,
    pushBlock: s.pushBlock,
    removeBlock: s.removeBlock,
    freshBlockIds: s.freshBlockIds,
    acknowledgeBlock: s.acknowledgeBlock,
    activeActionBlockId: s.activeActionBlockId,
    selectedActionIndices: s.selectedActionIndices,
    setActionSelection: s.setActionSelection,
    setActiveActionBlock: s.setActiveActionBlock,
    setFocusedContext: s.setFocusedContext,
    triggerAvatarActionCue: s.triggerAvatarActionCue,
    insertBlockAfter: s.insertBlockAfter,
    updateBlock: s.updateBlock,
    proxyConfig: s.proxyConfig
  }), shallow);
  const spTheme = useGrimoireStore((s) => s.spThemeColors);
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const [actionStatus, setActionStatus] = React.useState<string>('');
  const [recappingBlockId, setRecappingBlockId] = React.useState<string | undefined>();
  const closeButtonCopy = React.useMemo(
    () => getActionPanelCloseCopy(isSettingsOpen),
    [isSettingsOpen]
  );
  const hasShareableContent = React.useMemo(
    () => hasShareableSessionContent(blocks, transcript),
    [blocks, transcript]
  );

  const activeActionBlock = React.useMemo(
    () => blocks.find((b) => b.id === activeActionBlockId),
    [blocks, activeActionBlockId]
  );

  const allCandidates = React.useMemo(
    () => activeActionBlock ? getSelectionCandidates(activeActionBlock) : [],
    [activeActionBlock]
  );

  const selectedCandidates = React.useMemo(() => {
    if (!activeActionBlock || selectedActionIndices.length === 0 || allCandidates.length === 0) return [];
    const wanted = new Set<number>(selectedActionIndices);
    return allCandidates.filter((c) => wanted.has(c.index));
  }, [activeActionBlock, allCandidates, selectedActionIndices]);

  React.useEffect(() => {
    if (!activeActionBlock) return;
    if (selectedActionIndices.length > 0) return;
    if (allCandidates.length !== 1) return;
    setActionSelection(activeActionBlock.id, [allCandidates[0].index]);
  }, [activeActionBlock, allCandidates, selectedActionIndices, setActionSelection]);

  const focusPolicy = React.useMemo(
    () => evaluateHeaderAction('focus', activeActionBlock, selectedCandidates),
    [activeActionBlock, selectedCandidates]
  );
  const summarizePolicy = React.useMemo(
    () => evaluateHeaderAction('summarize', activeActionBlock, selectedCandidates),
    [activeActionBlock, selectedCandidates]
  );
  const chatPolicy = React.useMemo(
    () => evaluateHeaderAction('chat', activeActionBlock, selectedCandidates),
    [activeActionBlock, selectedCandidates]
  );

  const canFocus = focusPolicy.enabled;
  const canSummarize = summarizePolicy.enabled;
  const canChat = chatPolicy.enabled;
  const isGeneratingRecap = !!activeActionBlock && recappingBlockId === activeActionBlock.id;
  const canRecap = !!activeActionBlock && canRecapBlock(activeActionBlock) && !isGeneratingRecap;
  const recapUnavailableReason = !activeActionBlock
    ? 'No current result block.'
    : 'This block does not have enough visible data to recap.';
  const lastSelectionEventRef = React.useRef<{ signature: string; hadSelection: boolean }>({
    signature: '',
    hadSelection: false
  });

  const setTemporaryStatus = React.useCallback((message: string) => {
    setActionStatus(message);
  }, []);

  React.useEffect(() => {
    if (!actionStatus) return;
    const timer = setTimeout(() => setActionStatus(''), 5000);
    return () => clearTimeout(timer);
  }, [actionStatus]);

  React.useEffect(() => {
    const signature = activeActionBlock
      ? `${activeActionBlock.id}:${selectedActionIndices.join(',')}`
      : 'no-active-block';
    const previous = lastSelectionEventRef.current;
    if (previous.signature === signature) {
      return;
    }

    const hasSelection = selectedCandidates.length > 0;
    const shouldEmit = hasSelection || previous.hadSelection;
    lastSelectionEventRef.current = {
      signature,
      hadSelection: hasSelection
    };

    if (!shouldEmit) {
      return;
    }

    hybridInteractionEngine.emitEvent({
      eventName: 'task.selection.updated',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: `selection-${activeActionBlock?.id || 'none'}-${Date.now()}`,
      payload: buildSelectionSnapshotPayload(activeActionBlock, selectedCandidates),
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: activeActionBlock?.id,
      blockType: activeActionBlock?.type
    });
  }, [activeActionBlock, selectedActionIndices, selectedCandidates]);

  const handleFocus = React.useCallback(() => {
    if (!activeActionBlock || !focusPolicy.enabled) return;
    if (!isActionableBlockType(activeActionBlock.type)) return;

    const focusCandidates = getEligibleCandidatesForAction('focus', selectedCandidates);
    if (focusCandidates.length === 0) return;

    triggerAvatarActionCue('focus');
    logService.debug('system', `Avatar cue trigger: focus (${focusCandidates.length} item(s))`);

    const focusedItems = focusCandidates.map((c) => ({
      index: c.index,
      title: c.title,
      kind: c.kind,
      url: c.url,
      itemType: c.itemType,
      payload: c.payload,
      targetContext: deriveMcpTargetContextFromUnknown(c.payload, 'hie-selection')
    }));
    setFocusedContext({
      blockId: activeActionBlock.id,
      blockType: activeActionBlock.type,
      blockTitle: activeActionBlock.title,
      itemCount: focusedItems.length,
      items: focusedItems,
      updatedAt: Date.now()
    });

    const guardrails = buildFocusGuardrailSummary(focusCandidates);
    const focusContext = {
      scope: 'focus-context',
      block: {
        id: activeActionBlock.id,
        type: activeActionBlock.type,
        title: activeActionBlock.title
      },
      selectedCount: focusedItems.length,
      items: focusedItems.slice(0, 10).map((item) => {
        const payload = item.payload || {};
        return {
          index: item.index,
          title: item.title,
          kind: item.kind || item.itemType || 'item',
          url: item.url,
          itemId: payload.itemId,
          subject: payload.Subject || payload.subject,
          from: payload.From || payload.from,
          date: payload.Date || payload.date,
          start: payload.Start || payload.start,
          end: payload.End || payload.end
        };
      }),
      guardrails
    };

    const preview = focusedItems.slice(0, 5).map((item) => `${item.index}) ${item.title}`).join(', ');
    const suffix = focusedItems.length > 5 ? ` ...and ${focusedItems.length - 5} more` : '';
    hybridInteractionEngine.emitEvent({
      eventName: 'task.focused',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: `focus-${activeActionBlock.id}-${Date.now()}`,
      payload: {
        sourceBlockId: activeActionBlock.id,
        sourceBlockType: activeActionBlock.type,
        sourceBlockTitle: activeActionBlock.title,
        selectedCount: focusedItems.length,
        selectedItems: focusedItems.slice(0, 10).map((item) => ({
          index: item.index,
          title: item.title,
          url: item.url,
          kind: item.kind,
          itemType: item.itemType,
          targetContext: item.targetContext
        })),
        guardrails: focusContext.guardrails,
        preview: preview ? `${preview}${suffix}` : undefined
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: activeActionBlock.id,
      blockType: activeActionBlock.type
    });

    setTemporaryStatus(`Focused on ${focusedItems.length} item${focusedItems.length !== 1 ? 's' : ''}.`);
    if (guardrails.notes.length > 0) {
      setTemporaryStatus(`Focused on ${focusedItems.length} item(s). ${guardrails.notes[0]}`);
    }
  }, [activeActionBlock, focusPolicy.enabled, selectedCandidates, setFocusedContext, setTemporaryStatus, triggerAvatarActionCue]);

  const handleSummarize = React.useCallback(() => {
    if (!activeActionBlock || !summarizePolicy.enabled) return;

    const eligible = getEligibleCandidatesForAction('summarize', selectedCandidates);
    const cap = summarizePolicy.maxItems || 5;
    const capped = eligible.slice(0, cap);
    if (capped.length > 0) {
      triggerAvatarActionCue('summarize');
      logService.debug('system', `Avatar cue trigger: summarize (${capped.length} item(s))`);
    }
    for (let i = 0; i < capped.length; i++) {
      emitBlockInteraction({
        blockId: activeActionBlock.id,
        blockType: activeActionBlock.type,
        action: 'summarize',
        payload: capped[i].payload,
        timestamp: Date.now()
      });
    }
    const skippedUnsupported = selectedCandidates.length - eligible.length;
    if (eligible.length > capped.length) {
      setTemporaryStatus(`Summarize runs up to ${cap} items at once. Processed first ${capped.length}.`);
      return;
    }
    if (skippedUnsupported > 0) {
      setTemporaryStatus(`Summarizing ${capped.length} item(s). Skipped ${skippedUnsupported} unsupported selection(s).`);
    } else {
      setTemporaryStatus(`Summarizing ${capped.length} item${capped.length !== 1 ? 's' : ''}.`);
    }
  }, [activeActionBlock, selectedCandidates, summarizePolicy.enabled, summarizePolicy.maxItems, setTemporaryStatus, triggerAvatarActionCue]);

  const handleChat = React.useCallback(() => {
    if (!activeActionBlock || !chatPolicy.enabled) return;
    const chatCandidates = getEligibleCandidatesForAction('chat', selectedCandidates);
    const cap = chatPolicy.maxItems || 5;
    const capped = chatCandidates.slice(0, cap);
    if (capped.length === 0) return;
    triggerAvatarActionCue('chat');
    logService.debug('system', `Avatar cue trigger: chat (${capped.length} item(s))`);

    const focusedItems = capped.map((c) => ({
      index: c.index,
      title: c.title,
      kind: c.kind,
      url: c.url,
      itemType: c.itemType,
      payload: c.payload
    }));

    setFocusedContext({
      blockId: activeActionBlock.id,
      blockType: activeActionBlock.type,
      blockTitle: activeActionBlock.title,
      itemCount: focusedItems.length,
      items: focusedItems,
      updatedAt: Date.now()
    });

    const primary = capped[0];
    emitBlockInteraction({
      blockId: activeActionBlock.id,
      blockType: activeActionBlock.type,
      action: 'chat-about',
      payload: {
        ...primary.payload,
        title: primary.title,
        url: primary.url,
        selectedItems: capped.map((candidate) => ({
          index: candidate.index,
          title: candidate.title,
          kind: candidate.kind,
          url: candidate.url,
          itemType: candidate.itemType
        })),
        selectedCount: chatCandidates.length
      },
      timestamp: Date.now()
    });
    if (chatCandidates.length > capped.length) {
      setTemporaryStatus(`Chat runs up to ${cap} items at once. Using first ${capped.length}.`);
      return;
    }
    setTemporaryStatus(`Chat context set to ${capped.length} item${capped.length !== 1 ? 's' : ''}.`);
  }, [activeActionBlock, chatPolicy.enabled, chatPolicy.maxItems, selectedCandidates, setFocusedContext, setTemporaryStatus, triggerAvatarActionCue]);

  const handleRecap = React.useCallback(() => {
    if (!activeActionBlock || !canRecapBlock(activeActionBlock) || isGeneratingRecap) return;

    const recapTitle = `Recap: ${activeActionBlock.title}`;
    const recapOriginTool = getRecapOriginTool(activeActionBlock.id);
    const loadingData: IInfoCardData = {
      kind: 'info-card',
      heading: recapTitle,
      body: 'Generating recap…',
      icon: 'AlignLeft'
    };

    let recapBlock = blocks.find((block) => block.type === 'info-card' && block.originTool === recapOriginTool);
    if (recapBlock) {
      const nextBlock = { ...recapBlock, title: recapTitle, data: loadingData, timestamp: new Date() };
      updateBlock(recapBlock.id, { title: recapTitle, data: loadingData, timestamp: nextBlock.timestamp });
      hybridInteractionEngine.onBlockUpdated(recapBlock.id, nextBlock);
      recapBlock = nextBlock;
    } else {
      recapBlock = createBlock('info-card', recapTitle, loadingData, true, undefined, { originTool: recapOriginTool });
      insertBlockAfter(activeActionBlock.id, recapBlock);
      hybridInteractionEngine.onBlockCreated(recapBlock);
    }

    setActiveActionBlock(activeActionBlock.id);
    setRecappingBlockId(activeActionBlock.id);
    setTemporaryStatus('Generating recap.');
    hybridInteractionEngine.emitEvent({
      eventName: 'task.recap.requested',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: `recap-${activeActionBlock.id}-${Date.now()}`,
      payload: {
        sourceBlockId: activeActionBlock.id,
        sourceBlockType: activeActionBlock.type,
        sourceBlockTitle: activeActionBlock.title,
        derivedBlockId: recapBlock.id,
        derivedBlockTitle: recapTitle
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: activeActionBlock.id,
      blockType: activeActionBlock.type
    });

    const sourceBlock = activeActionBlock;
    const recapBlockId = recapBlock.id;
    const recapService = new BlockRecapService();
    recapService.generate(sourceBlock, proxyConfig)
      .then((text) => {
        const readyData: IInfoCardData = {
          kind: 'info-card',
          heading: recapTitle,
          body: text,
          icon: 'AlignLeft'
        };
        const nextBlock = {
          ...recapBlock,
          title: recapTitle,
          data: readyData,
          timestamp: new Date()
        };
        updateBlock(recapBlockId, { title: recapTitle, data: readyData, timestamp: nextBlock.timestamp });
        hybridInteractionEngine.onBlockUpdated(recapBlockId, nextBlock);
        hybridInteractionEngine.emitEvent({
          eventName: 'artifact.recap.ready',
          source: 'action-panel',
          surface: 'action-panel',
          correlationId: `recap-ready-${sourceBlock.id}-${Date.now()}`,
          payload: {
            sourceBlockId: sourceBlock.id,
            sourceBlockType: sourceBlock.type,
            sourceBlockTitle: sourceBlock.title,
            derivedBlockId: recapBlockId,
            derivedBlockTitle: recapTitle
          },
          exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
          blockId: recapBlockId,
          blockType: nextBlock.type
        });
        setActiveActionBlock(sourceBlock.id);
        setTemporaryStatus('Recap ready.');
      })
      .catch((error: Error) => {
        const errorData: IInfoCardData = {
          kind: 'info-card',
          heading: recapTitle,
          body: error.message || 'The recap could not be generated.',
          icon: 'StatusErrorFull'
        };
        const nextBlock = {
          ...recapBlock,
          title: recapTitle,
          data: errorData,
          timestamp: new Date()
        };
        updateBlock(recapBlockId, { title: recapTitle, data: errorData, timestamp: nextBlock.timestamp });
        hybridInteractionEngine.onBlockUpdated(recapBlockId, nextBlock);
        hybridInteractionEngine.emitEvent({
          eventName: 'artifact.recap.failed',
          source: 'action-panel',
          surface: 'action-panel',
          correlationId: `recap-failed-${sourceBlock.id}-${Date.now()}`,
          payload: {
            sourceBlockId: sourceBlock.id,
            sourceBlockType: sourceBlock.type,
            sourceBlockTitle: sourceBlock.title,
            derivedBlockId: recapBlockId,
            derivedBlockTitle: recapTitle,
            errorMessage: error.message || 'The recap could not be generated.'
          },
          exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
          blockId: recapBlockId,
          blockType: nextBlock.type
        });
        setActiveActionBlock(sourceBlock.id);
        setTemporaryStatus('Recap failed.');
      })
      .finally(() => {
        setRecappingBlockId((current) => (current === sourceBlock.id ? undefined : current));
      });
  }, [activeActionBlock, blocks, insertBlockAfter, isGeneratingRecap, proxyConfig, setActiveActionBlock, setTemporaryStatus, updateBlock]);

  const handleClosePanel = React.useCallback(() => {
    switch (resolveActionPanelCloseAction(isSettingsOpen)) {
      case 'close-settings':
        onCloseSettings?.();
        return;
      default:
        onCloseApp?.();
    }
  }, [isSettingsOpen, onCloseApp, onCloseSettings]);

  const handleOpenShareForm = React.useCallback((target: 'email' | 'teams-chat' | 'teams-channel') => {
    if (!hasShareableContent) {
      setTemporaryStatus('Nothing shareable is visible yet.');
      return;
    }

    const sourceContext = hybridInteractionEngine.captureCurrentSourceContext();

    const shareContent = new SessionShareFormatter().format({
      blocks,
      transcript,
      activeBlockId: activeActionBlockId,
      selectedActionIndices
    });

    let formTitle = 'Share';
    let formData: IFormData;

    switch (target) {
      case 'email':
        formTitle = 'Share via Email';
        formData = buildPrefilledFormData(
          'email-compose',
          'Review the email before sending. Recipients stay editable.',
          {
            subject: shareContent.subject,
            body: shareContent.detailedPlainText
          },
          shareContent.attachmentUris.length > 0 ? { attachmentUris: shareContent.attachmentUris } : undefined
        );
        break;
      case 'teams-chat':
        formTitle = 'Share to New Teams Chat';
        formData = buildPrefilledFormData(
          'share-teams-chat',
          'Create a new Teams chat, then post this session summary.',
          {
            topic: shareContent.subject,
            content: shareContent.detailedPlainText
          }
        );
        break;
      default:
        formTitle = 'Share to Teams Channel';
        formData = buildPrefilledFormData(
          'share-teams-channel',
          'Resolve a team and channel by name, then post this session summary.',
          {
            content: shareContent.detailedPlainText
          }
        );
        break;
    }

    if (sourceContext?.targetContext) {
      formData = {
        ...formData,
        submissionTarget: {
          ...formData.submissionTarget,
          targetContext: sourceContext.targetContext
        }
      };
    }

    const formOriginTool = activeActionBlockId ? `share-form:${activeActionBlockId}` : 'share-form';
    const formBlock = createBlock('form', formTitle, formData, true, undefined, { originTool: formOriginTool });
    pushBlock(formBlock);
    hybridInteractionEngine.onBlockCreated(formBlock, sourceContext);
    setTemporaryStatus(`${formTitle} ready.`);
  }, [activeActionBlockId, blocks, hasShareableContent, pushBlock, selectedActionIndices, setTemporaryStatus, transcript]);

  const shareMenuProps = React.useMemo(() => ({
    items: [
      {
        key: 'share-email',
        text: 'Email',
        iconProps: { iconName: 'Mail' },
        onClick: () => handleOpenShareForm('email')
      },
      {
        key: 'share-teams-chat',
        text: 'New Teams Chat',
        iconProps: { iconName: 'Chat' },
        onClick: () => handleOpenShareForm('teams-chat')
      },
      {
        key: 'share-teams-channel',
        text: 'Teams Channel',
        iconProps: { iconName: 'TeamsLogo' },
        onClick: () => handleOpenShareForm('teams-channel')
      }
    ]
  }), [handleOpenShareForm]);

  React.useEffect(() => {
    injectStyleOnce(SCROLL_STYLE_ID, SCROLL_CSS);
    injectStyleOnce(SHIMMER_STYLE_ID, SHIMMER_CSS);
  }, []);

  React.useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [blocks.length]);

  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      minWidth: 0,
      backgroundColor: spTheme.bodyBackground,
      borderLeft: `1px solid ${spTheme.cardBorder}`,
      borderRight: `1px solid ${spTheme.cardBorder}`,
      overflow: 'hidden'
    }}>
      <div style={{
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'flex-start',
        justifyContent: 'space-between',
        padding: '12px 16px',
        borderBottom: `1px solid ${spTheme.cardBorder}`,
        gap: 10,
        flexShrink: 0
      }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 0, flex: '1 1 220px' }}>
          <Text
            variant="medium"
            styles={{ root: { color: spTheme.bodyText, fontWeight: 600 } }}
          >
            Hybrid Interaction Engine
          </Text>
          <span style={{ fontSize: 11, color: spTheme.bodySubtext }}>
            {activeActionBlock
              ? `Current results: ${activeActionBlock.title}`
              : 'No actionable result block yet'}
            {activeActionBlock && ` • Selected: ${selectedCandidates.length}`}
          </span>
          {actionStatus && (
            <span style={{ fontSize: 11, color: '#0064b4' }}>{actionStatus}</span>
          )}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
          {blocks.length > 0 && (
            <>
              <button
                type="button"
                style={actionButtonStyle(canFocus)}
                onClick={canFocus ? handleFocus : undefined}
                title={canFocus ? 'Set local focus on selected items' : (focusPolicy.reason || 'Focus is unavailable')}
              >
                {strings.FocusButton}
              </button>
              <button
                type="button"
                style={actionButtonStyle(canSummarize)}
                onClick={canSummarize ? handleSummarize : undefined}
                title={canSummarize
                  ? `Summarize selected items (max ${summarizePolicy.maxItems || 5})`
                  : (summarizePolicy.reason || 'Summarize is unavailable')}
              >
                {strings.SummarizeButton}
              </button>
              <button
                type="button"
                style={actionButtonStyle(canChat)}
                onClick={canChat ? handleChat : undefined}
                title={canChat
                  ? `Start chat about selected items (max ${chatPolicy.maxItems || 5})`
                  : (chatPolicy.reason || 'Chat is unavailable')}
              >
                {strings.ChatButton}
              </button>
              <button
                type="button"
                style={actionButtonStyle(canRecap)}
                onClick={canRecap ? handleRecap : undefined}
                title={canRecap ? 'Generate a short recap of the current results or data block' : recapUnavailableReason}
              >
                {isGeneratingRecap ? strings.RecapLoadingButton : strings.RecapButton}
              </button>
            </>
          )}
          {hasShareableContent && (
            <DefaultButton
              text={strings.ShareButton}
              iconProps={{ iconName: 'Share' }}
              menuProps={shareMenuProps}
              styles={{
                root: {
                  minWidth: 78,
                  height: 28,
                  padding: '0 10px',
                  borderRadius: 999,
                  background: spTheme.bodyBackground,
                  color: spTheme.bodySubtext,
                  border: `1px solid ${spTheme.cardBorder}`
                },
                rootHovered: {
                  background: 'rgba(15, 108, 189, 0.08)',
                  color: spTheme.bodyText,
                  border: `1px solid ${spTheme.cardBorder}`
                },
                label: {
                  fontSize: 11,
                  fontWeight: 600
                },
                menuIcon: {
                  color: spTheme.bodySubtext
                }
              }}
            />
          )}
          {(onCloseApp || onCloseSettings) && (
            <IconButton
              iconProps={{ iconName: 'ChromeClose' }}
              ariaLabel={closeButtonCopy.ariaLabel}
              title={closeButtonCopy.title}
              onClick={handleClosePanel}
              styles={{
                root: { color: spTheme.bodySubtext, width: 28, height: 28 },
                rootHovered: { color: spTheme.bodyText, backgroundColor: 'transparent' }
              }}
            />
          )}
        </div>
      </div>

      {blocks.length === 0 ? (
        <div style={{
          flex: 1,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: spTheme.bodySubtext,
          fontStyle: 'italic',
          fontSize: 14,
          fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif'
        }}>
          {strings.EmptyPanelHint}
        </div>
      ) : (
        <div ref={scrollRef} className="grimoire-scroll" style={{
          flex: 1,
          overflowY: 'auto',
          padding: 16,
          display: 'flex',
          flexDirection: 'column',
          gap: 12
        }}>
          {blocks.map((block) => (
            <BlockCard
              key={block.id}
              block={block}
              isFresh={freshBlockIds.includes(block.id)}
              onAcknowledge={() => {
                acknowledgeBlock(block.id);
                hybridInteractionEngine.onBlockAcknowledged(block.id);
              }}
              onDismiss={() => { hybridInteractionEngine.onBlockRemoved(block.id); removeBlock(block.id); }}
              cardBg={spTheme.cardBackground}
              cardBorder={spTheme.cardBorder}
              textColor={spTheme.bodyText}
              subtextColor={spTheme.bodySubtext}
            />
          ))}
        </div>
      )}
    </div>
  );
};
