import type {
  IContextMessage,
  IHieArtifactRecord,
  IHieEvent,
  IHieTaskContext,
  IHieTurnLineage,
  VerbosityLevel
} from '../../services/hie/HIETypes';
import type { IVisualStateSnapshot } from '../../services/hie/HAEContracts';
import type { IExpressionTriggerSnapshot } from '../../services/hie/DataExpressionDirector';
import {
  describeMcpTargetContext,
  mergeMcpTargetContexts,
  type IMcpTargetContext
} from '../../services/mcp/McpTargetContext';
import { resolveCurrentArtifactContext, resolveLatestArtifactContext } from '../../services/hie/HieArtifactLinkage';

function clip(value: string, maxChars: number): string {
  const trimmed = value.trim();
  if (trimmed.length <= maxChars) {
    return trimmed;
  }
  return `${trimmed.slice(0, Math.max(0, maxChars - 1)).trimEnd()}...`;
}

// ─── Shared Interfaces ─────────────────────────────────────────

export interface IHieThreadNode {
  key: string;
  role: 'Root' | 'Parent' | 'Current';
  value: string;
}

export interface IHieLinkageNode {
  key: string;
  label: string;
  value: string;
  meta?: string;
}

export interface IHieTargetContextSummary {
  value: string;
  description?: string;
  meta?: string;
}

// ─── State Compass ─────────────────────────────────────────────

export type CompassCardKey = 'thread' | 'task' | 'content' | 'artifacts';

export interface IStateCompassCard {
  key: CompassCardKey;
  label: string;
  value: string;
  isEmpty: boolean;
  /** Merged description + reason (shown when expanded) */
  detail?: string;
  /** Turn IDs and metadata (shown when expanded) */
  meta?: string;
}

// ─── Situation Pills ───────────────────────────────────────────

export interface ISituationPill {
  key: string;
  label: string;
  accent: string;
}

// ─── Shared Utilities ──────────────────────────────────────────

export function shortTurnId(turnId?: string): string {
  if (!turnId) {
    return 'n/a';
  }
  if (turnId.length <= 18) {
    return turnId;
  }
  return `${turnId.slice(0, 10)}...${turnId.slice(-4)}`;
}

function buildMeta(parts: Array<string | undefined>): string | undefined {
  const filtered = parts.filter((part): part is string => !!part && part.trim().length > 0);
  return filtered.length > 0 ? filtered.join(' · ') : undefined;
}

// ─── Target Context ────────────────────────────────────────────

function getTargetSourceLabel(source?: IMcpTargetContext['source']): string | undefined {
  switch (source) {
    case 'explicit-user':
      return 'explicit user target';
    case 'hie-selection':
      return 'HIE selection';
    case 'current-page':
      return 'current page fallback';
    case 'recovered':
      return 'recovered target';
    default:
      return undefined;
  }
}

function buildTargetContextDescription(targetContext?: Readonly<IMcpTargetContext>): string | undefined {
  if (!targetContext) {
    return undefined;
  }

  const parts = [
    targetContext.siteName || targetContext.siteUrl ? `Site: ${targetContext.siteName || targetContext.siteUrl}` : undefined,
    targetContext.documentLibraryName || targetContext.documentLibraryUrl
      ? `Library: ${targetContext.documentLibraryName || targetContext.documentLibraryUrl}`
      : undefined,
    targetContext.listName || targetContext.listUrl ? `List: ${targetContext.listName || targetContext.listUrl}` : undefined,
    targetContext.fileOrFolderName || targetContext.fileOrFolderUrl
      ? `Item: ${targetContext.fileOrFolderName || targetContext.fileOrFolderUrl}`
      : undefined,
    targetContext.personDisplayName || targetContext.personEmail
      ? `Person: ${targetContext.personDisplayName || targetContext.personEmail}`
      : undefined,
    targetContext.teamName || targetContext.channelName
      ? `Destination: ${[targetContext.teamName, targetContext.channelName].filter(Boolean).join(' / ')}`
      : undefined
  ].filter((part): part is string => !!part);

  return parts.length > 0 ? parts.join(' · ') : undefined;
}

export function buildHieTargetContextSummary(
  task?: Readonly<IHieTaskContext>,
  artifacts: Readonly<Record<string, IHieArtifactRecord>> = {}
): IHieTargetContextSummary | undefined {
  const context = task
    ? resolveCurrentArtifactContext(task, artifacts)
    : resolveLatestArtifactContext(artifacts);
  const latestContext = resolveLatestArtifactContext(artifacts);
  const targetContext = mergeMcpTargetContexts(
    task?.targetContext,
    context.currentArtifact?.targetContext,
    context.primaryArtifact?.targetContext,
    latestContext.currentArtifact?.targetContext,
    latestContext.primaryArtifact?.targetContext
  );
  const value = describeMcpTargetContext(targetContext);
  if (!value) {
    return undefined;
  }

  return {
    value,
    description: buildTargetContextDescription(targetContext),
    meta: buildMeta([
      getTargetSourceLabel(targetContext?.source),
      targetContext?.siteId ? 'site id ready' : undefined,
      targetContext?.documentLibraryId ? 'library id ready' : undefined,
      targetContext?.listId ? 'list id ready' : undefined,
      targetContext?.fileOrFolderId ? 'item id ready' : undefined,
      targetContext?.teamId ? 'team id ready' : undefined,
      targetContext?.channelId ? 'channel id ready' : undefined
    ])
  };
}

// ─── Task Formatting ───────────────────────────────────────────

export function formatHieTaskMeta(task?: Readonly<IHieTaskContext>): string | undefined {
  if (!task) {
    return undefined;
  }

  return buildMeta([
    task.eventName,
    task.correlationId ? `corr ${shortTurnId(task.correlationId)}` : undefined,
    task.turnId ? `turn ${shortTurnId(task.turnId)}` : undefined
  ]);
}

export function formatHieTaskSummary(task?: Readonly<IHieTaskContext>): string {
  if (!task) {
    return 'No clear task yet';
  }

  switch (task.kind) {
    case 'search':
      return task.sourceBlockTitle || 'Current search results';
    case 'recap':
      return `Recap: ${task.derivedBlockTitle || task.sourceBlockTitle || task.derivedBlockId || 'current artifact'}`;
    case 'form':
      return `Form: ${task.formPreset || task.derivedBlockTitle || 'compose'} (${task.formStatus || 'opened'})`;
    case 'focus':
      return `Focus: ${task.sourceBlockTitle || task.sourceBlockId || 'current selection'}`;
    case 'summarize':
      return `Summary: ${task.derivedBlockTitle || task.sourceBlockTitle || task.sourceBlockId || 'current content'}`;
    case 'chat-about':
      return `Chat: ${task.sourceBlockTitle || task.sourceBlockId || 'current content'}`;
    case 'select':
      return `Selection: ${task.sourceBlockTitle || task.sourceBlockId || 'current options'}`;
    default:
      return `${task.kind}: ${task.sourceBlockTitle || task.derivedBlockTitle || task.sourceBlockId || 'current task'}`;
  }
}

// ─── Event & Artifact Description ──────────────────────────────

export function describeHieEvent(entry: IHieEvent): string {
  if (entry.eventName.startsWith('thread.')) {
    const resolvedMode = typeof entry.payload.resolvedMode === 'string' ? entry.payload.resolvedMode : undefined;
    const reason = typeof entry.payload.reason === 'string' ? entry.payload.reason : undefined;
    const suffixParts = [resolvedMode, reason].filter((part): part is string => !!part);
    return `${entry.eventName} · ${entry.exposurePolicy.mode}${suffixParts.length > 0 ? ` · ${suffixParts.join(' · ')}` : ''}`;
  }

  const payloadTitle = typeof entry.payload.derivedBlockTitle === 'string'
    ? entry.payload.derivedBlockTitle
    : (typeof entry.payload.sourceBlockTitle === 'string'
      ? entry.payload.sourceBlockTitle
      : (typeof entry.payload.blockTitle === 'string' ? entry.payload.blockTitle : undefined));
  const suffix = payloadTitle ? ` · ${payloadTitle}` : '';
  return `${entry.eventName} · ${entry.exposurePolicy.mode}${suffix}`;
}

export function describeHieArtifact(artifact: IHieArtifactRecord): string {
  const title = artifact.title || artifact.preset || artifact.blockId || artifact.artifactId;
  return `${title} (${artifact.status})`;
}

export function getArtifactLinkageLabel(artifact: IHieArtifactRecord): string {
  switch (artifact.artifactKind) {
    case 'summary':
      return 'Summary';
    case 'preview':
      return 'Preview';
    case 'lookup':
      return 'Lookup';
    case 'error':
      return 'Error';
    case 'recap':
      return 'Recap';
    case 'share':
      return 'Share';
    case 'form':
      return 'Form';
    case 'block':
      return artifact.blockType ? artifact.blockType.replace(/-/g, ' ') : 'Artifact';
    default:
      return 'Artifact';
  }
}

// ─── Private Compass Helpers ───────────────────────────────────

function describeThreadReason(event?: Readonly<IHieEvent>): string | undefined {
  if (!event) {
    return undefined;
  }

  const resolvedMode = typeof event.payload.resolvedMode === 'string' ? event.payload.resolvedMode : undefined;
  const reason = typeof event.payload.reason === 'string' ? event.payload.reason : undefined;

  switch (reason) {
    case 'contextual-follow-up':
    case 'active-task-context':
      return 'the user is continuing the current path';
    case 'visible-title-follow-up':
      return 'the user referred to visible content already on screen';
    case 'explicit-reset-phrase':
      return 'the user explicitly started a new topic';
    case 'explicit-new-root':
      return 'a new root thread was started explicitly';
    case 'no-active-context':
      return 'no earlier path was active';
    default:
      if (resolvedMode === 'inherit') {
        return 'the current path was inherited from the previous turn';
      }
      if (resolvedMode === 'new-root') {
        return 'a fresh root thread was started';
      }
      return undefined;
  }
}

function describeTaskOrigin(task: Readonly<IHieTaskContext>): string {
  switch (task.eventName) {
    case 'block.created':
    case 'block.updated':
      return 'visible content appearing in the UI';
    case 'task.focused':
      return 'the user focusing visible content';
    case 'form.opened':
      return 'opening a form from the current path';
    case 'artifact.recap.ready':
    case 'task.recap.requested':
      return 'generating a recap from the current path';
    default:
      if (task.eventName.startsWith('block.interaction.')) {
        return `the user ${task.eventName.slice('block.interaction.'.length).replace(/-/g, ' ')}`;
      }
      return task.eventName.replace(/\./g, ' ');
  }
}

function formatThreadSummary(lineage?: Readonly<IHieTurnLineage>): string {
  if (!lineage) {
    return 'No thread yet';
  }

  if (!lineage.parentTurnId || lineage.turnId === lineage.rootTurnId) {
    return 'Root thread';
  }

  return 'Follow-up thread';
}

function formatThreadDetail(
  lineage?: Readonly<IHieTurnLineage>,
  latestThreadEvent?: Readonly<IHieEvent>
): string | undefined {
  if (!lineage) {
    return undefined;
  }

  const threadReason = describeThreadReason(latestThreadEvent);
  const isRoot = !lineage.parentTurnId || lineage.turnId === lineage.rootTurnId;
  const reasonText = isRoot
    ? (threadReason
      ? `Root thread because ${threadReason}.`
      : 'Root thread for the active path.')
    : (threadReason
      ? `Continues the current thread because ${threadReason}.`
      : 'Keeps the current follow-up path together.');

  return reasonText;
}

function formatTaskDetail(task?: Readonly<IHieTaskContext>): string | undefined {
  if (!task) {
    return undefined;
  }

  const origin = describeTaskOrigin(task);
  const subject = task.sourceBlockTitle || task.derivedBlockTitle || task.sourceBlockId || 'the current path';
  return `From ${origin} around "${subject}".`;
}

function formatVisibleContentValue(
  snapshot?: Readonly<IVisualStateSnapshot>,
  task?: Readonly<IHieTaskContext>
): string {
  if (!snapshot || snapshot.blocks.length === 0) {
    return 'Nothing visible';
  }

  const focusedBlock = task?.sourceBlockId
    ? snapshot.blocks.find((block) => block.blockId === task.sourceBlockId)
    : undefined;
  if (focusedBlock) {
    return clip(focusedBlock.title, 30);
  }

  if (snapshot.blocks.length === 1) {
    return clip(snapshot.blocks[0].title, 30);
  }

  return `${snapshot.blocks.length} blocks`;
}

function formatDerivedOutputValue(
  task?: Readonly<IHieTaskContext>,
  artifacts: Readonly<Record<string, IHieArtifactRecord>> = {}
): string {
  const context = task
    ? resolveCurrentArtifactContext(task, artifacts)
    : resolveLatestArtifactContext(artifacts);
  const currentArtifact = context.currentArtifact || context.primaryArtifact;
  if (!currentArtifact) {
    return 'No output yet';
  }

  const title = currentArtifact.title || currentArtifact.preset || currentArtifact.blockId || currentArtifact.artifactId;
  return clip(title, 30);
}

function formatDerivedOutputDetail(
  task?: Readonly<IHieTaskContext>,
  artifacts: Readonly<Record<string, IHieArtifactRecord>> = {}
): string | undefined {
  const context = task
    ? resolveCurrentArtifactContext(task, artifacts)
    : resolveLatestArtifactContext(artifacts);
  const currentArtifact = context.currentArtifact || context.primaryArtifact;
  if (!currentArtifact) {
    return undefined;
  }

  const chain = context.artifactChain.length > 0
    ? context.artifactChain.map((artifact) => getArtifactLinkageLabel(artifact)).join(' -> ')
    : undefined;
  const sourceTitle = task?.sourceBlockTitle;
  const artifactLabel = getArtifactLinkageLabel(currentArtifact).toLowerCase();
  const title = currentArtifact.title || currentArtifact.preset || currentArtifact.blockId || currentArtifact.artifactId;

  const parts: string[] = [];
  if (sourceTitle) {
    parts.push(`${artifactLabel} "${title}" from "${sourceTitle}".`);
  } else {
    parts.push(`${artifactLabel} "${title}".`);
  }
  if (chain) {
    parts.push(`Chain: ${chain}`);
  }
  return parts.join(' ');
}

function formatDerivedOutputMeta(
  task?: Readonly<IHieTaskContext>,
  artifacts: Readonly<Record<string, IHieArtifactRecord>> = {}
): string | undefined {
  const context = task
    ? resolveCurrentArtifactContext(task, artifacts)
    : resolveLatestArtifactContext(artifacts);
  const artifact = context.currentArtifact || context.primaryArtifact;
  if (!artifact) {
    return undefined;
  }

  return buildMeta([
    artifact.sourceEventName,
    artifact.correlationId ? `corr ${shortTurnId(artifact.correlationId)}` : undefined,
    artifact.sourceTurnId ? `turn ${shortTurnId(artifact.sourceTurnId)}` : undefined
  ]);
}

// ─── State Compass Builder ─────────────────────────────────────

const EMPTY_COMPASS_VALUES: Record<CompassCardKey, string> = {
  thread: 'No thread yet',
  task: 'No clear task yet',
  content: 'Nothing visible',
  artifacts: 'No output yet'
};

export function buildStateCompassCards(
  lineage: Readonly<IHieTurnLineage> | undefined,
  latestThreadEvent: Readonly<IHieEvent> | undefined,
  task: Readonly<IHieTaskContext> | undefined,
  artifacts: Readonly<Record<string, IHieArtifactRecord>> = {},
  snapshot?: Readonly<IVisualStateSnapshot>
): IStateCompassCard[] {
  const threadValue = formatThreadSummary(lineage);
  const taskValue = formatHieTaskSummary(task);
  const contentValue = formatVisibleContentValue(snapshot, task);
  const artifactsValue = formatDerivedOutputValue(task, artifacts);

  return [
    {
      key: 'thread',
      label: 'Thread',
      value: threadValue,
      isEmpty: threadValue === EMPTY_COMPASS_VALUES.thread,
      detail: formatThreadDetail(lineage, latestThreadEvent),
      meta: lineage
        ? buildMeta([
          `root ${shortTurnId(lineage.rootTurnId)}`,
          lineage.parentTurnId && lineage.parentTurnId !== lineage.rootTurnId
            ? `parent ${shortTurnId(lineage.parentTurnId)}`
            : undefined,
          `current ${shortTurnId(lineage.turnId)}`
        ])
        : undefined
    },
    {
      key: 'task',
      label: 'Task',
      value: taskValue,
      isEmpty: taskValue === EMPTY_COMPASS_VALUES.task,
      detail: formatTaskDetail(task),
      meta: formatHieTaskMeta(task)
    },
    {
      key: 'content',
      label: 'Content',
      value: contentValue,
      isEmpty: contentValue === EMPTY_COMPASS_VALUES.content,
      detail: snapshot && snapshot.blocks.length > 0
        ? snapshot.blocks
          .slice(0, 4)
          .map((block) => `${block.blockType}: ${block.title}`)
          .join(', ')
        : undefined
    },
    {
      key: 'artifacts',
      label: 'Artifacts',
      value: artifactsValue,
      isEmpty: artifactsValue === EMPTY_COMPASS_VALUES.artifacts,
      detail: formatDerivedOutputDetail(task, artifacts),
      meta: formatDerivedOutputMeta(task, artifacts)
    }
  ];
}

export function getDefaultExpandedCompassCard(
  cards: ReadonlyArray<IStateCompassCard>
): CompassCardKey | undefined {
  const firstMeaningful = cards.find((card) => !card.isEmpty);
  return firstMeaningful?.key;
}

// ─── Situation Pills Builder ───────────────────────────────────

const PILL_ACCENTS = {
  flow: '#7c3aed',
  verbosity: '#a16207',
  expression: '#0f766e',
  target: '#0f6cbd',
  idle: '#64748b'
};

export function buildSituationPills(options: {
  flowState?: { flowName: string; stepIndex: number; totalSteps: number };
  verbosity: VerbosityLevel;
  expressionTrigger?: Readonly<IExpressionTriggerSnapshot>;
  targetSummary?: IHieTargetContextSummary;
}): ISituationPill[] {
  const pills: ISituationPill[] = [];

  if (options.flowState) {
    pills.push({
      key: 'flow',
      label: `${options.flowState.flowName} ${options.flowState.stepIndex}/${options.flowState.totalSteps}`,
      accent: PILL_ACCENTS.flow
    });
  }

  if (options.verbosity !== 'normal') {
    pills.push({
      key: 'verbosity',
      label: `Verbosity: ${options.verbosity}`,
      accent: PILL_ACCENTS.verbosity
    });
  }

  if (options.expressionTrigger && (Date.now() - options.expressionTrigger.firedAt) < 10000) {
    pills.push({
      key: 'expression',
      label: `${options.expressionTrigger.expression} · ${options.expressionTrigger.triggerId}`,
      accent: PILL_ACCENTS.expression
    });
  }

  if (options.targetSummary) {
    pills.push({
      key: 'target',
      label: `Target: ${clip(options.targetSummary.value, 25)}`,
      accent: PILL_ACCENTS.target
    });
  }

  if (pills.length === 0) {
    pills.push({
      key: 'idle',
      label: 'Idle',
      accent: PILL_ACCENTS.idle
    });
  }

  return pills;
}

// ─── Linkage Nodes ─────────────────────────────────────────────

export function buildHieLinkageNodes(
  task?: Readonly<IHieTaskContext>,
  artifacts: Readonly<Record<string, IHieArtifactRecord>> = {}
): IHieLinkageNode[] {
  const nodes: IHieLinkageNode[] = [];
  const context = task
    ? resolveCurrentArtifactContext(task, artifacts)
    : resolveLatestArtifactContext(artifacts);

  if (task?.sourceBlockTitle || task?.sourceBlockId) {
    nodes.push({
      key: `source-${task.sourceBlockId || task.sourceBlockTitle}`,
      label: 'Source',
      value: task.sourceBlockTitle || task.sourceBlockId || 'Current block',
      meta: formatHieTaskMeta(task)
    });
  }

  context.artifactChain.forEach((artifact) => {
    nodes.push({
      key: `artifact-${artifact.artifactId}`,
      label: getArtifactLinkageLabel(artifact),
      value: artifact.title || artifact.preset || artifact.blockId || artifact.artifactId,
      meta: buildMeta([
        artifact.sourceEventName,
        artifact.correlationId ? `corr ${shortTurnId(artifact.correlationId)}` : undefined,
        artifact.sourceTurnId ? `turn ${shortTurnId(artifact.sourceTurnId)}` : undefined
      ])
    });
  });

  if (nodes.length === 0 && task) {
    nodes.push({
      key: `task-${task.updatedAt}`,
      label: 'Task',
      value: formatHieTaskSummary(task)
    });
  }

  return nodes;
}

// ─── Thread Nodes ──────────────────────────────────────────────

export function buildThreadNodes(lineage?: Readonly<IHieTurnLineage>): IHieThreadNode[] {
  if (!lineage) {
    return [];
  }

  const nodes: IHieThreadNode[] = [];
  const seen = new Set<string>();

  const pushNode = (role: IHieThreadNode['role'], turnId?: string): void => {
    if (!turnId || seen.has(turnId)) {
      return;
    }
    seen.add(turnId);
    nodes.push({
      key: `${role.toLowerCase()}-${turnId}`,
      role,
      value: shortTurnId(turnId)
    });
  };

  pushNode('Root', lineage.rootTurnId);
  if (lineage.parentTurnId && lineage.parentTurnId !== lineage.rootTurnId) {
    pushNode('Parent', lineage.parentTurnId);
  }
  pushNode('Current', lineage.turnId);

  if (nodes.length === 1) {
    return [{ ...nodes[0], role: 'Current' }];
  }

  return nodes;
}

// ─── Context Preview ───────────────────────────────────────────

export function formatHieContextPreview(message: IContextMessage): string {
  const compact = message.text.replace(/\s+/g, ' ').trim();
  const unwrapped = compact.startsWith('[') && compact.endsWith(']')
    ? compact.slice(1, -1)
    : compact;
  return clip(unwrapped, 160);
}
