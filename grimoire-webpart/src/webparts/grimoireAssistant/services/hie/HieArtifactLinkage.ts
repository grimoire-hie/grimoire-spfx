import type { IHieArtifactRecord, IHieTaskContext } from './HIETypes';
import { mergeMcpTargetContexts, type IMcpTargetContext } from '../mcp/McpTargetContext';

export interface IHieResolvedArtifactContext {
  currentArtifact?: IHieArtifactRecord;
  primaryArtifact?: IHieArtifactRecord;
  artifactChain: IHieArtifactRecord[];
  targetContext?: IMcpTargetContext;
}

const INACTIVE_ARTIFACT_STATUSES = new Set<IHieArtifactRecord['status']>([
  'cancelled',
  'dismissed'
]);

function getLatestRelevantArtifact(
  artifacts: Readonly<Record<string, IHieArtifactRecord>>
): IHieArtifactRecord | undefined {
  const values = Object.values(artifacts);
  if (values.length === 0) {
    return undefined;
  }

  const activeArtifacts = values
    .filter((artifact) => !INACTIVE_ARTIFACT_STATUSES.has(artifact.status))
    .sort((left, right) => right.updatedAt - left.updatedAt);
  if (activeArtifacts.length > 0) {
    return activeArtifacts[0];
  }

  return values.sort((left, right) => right.updatedAt - left.updatedAt)[0];
}

export function resolveArtifactChain(
  artifacts: Readonly<Record<string, IHieArtifactRecord>>,
  startArtifactId?: string
): IHieArtifactRecord[] {
  if (!startArtifactId) {
    return [];
  }

  const chain: IHieArtifactRecord[] = [];
  const seen = new Set<string>();
  let current: IHieArtifactRecord | undefined = artifacts[startArtifactId];

  while (current && !seen.has(current.artifactId)) {
    seen.add(current.artifactId);
    chain.push(current);
    const nextArtifactId: string | undefined = current.sourceArtifactId;
    current = nextArtifactId ? artifacts[nextArtifactId] : undefined;
  }

  return chain.reverse();
}

function resolvePrimaryArtifact(
  currentArtifact: IHieArtifactRecord | undefined,
  artifactChain: IHieArtifactRecord[]
): IHieArtifactRecord | undefined {
  if (!currentArtifact) {
    return undefined;
  }

  return [...artifactChain].reverse().find((artifact) => (
    artifact.artifactKind !== 'form' && artifact.artifactKind !== 'share'
  )) || currentArtifact;
}

export function resolveCurrentArtifactContext(
  taskContext: Readonly<IHieTaskContext> | undefined,
  artifacts: Readonly<Record<string, IHieArtifactRecord>>
): IHieResolvedArtifactContext {
  const currentArtifact = taskContext?.derivedBlockId ? artifacts[taskContext.derivedBlockId] : undefined;
  const artifactChain = resolveArtifactChain(artifacts, currentArtifact?.artifactId);
  const targetContext = mergeMcpTargetContexts(
    taskContext?.targetContext,
    currentArtifact?.targetContext,
    ...artifactChain.map((artifact) => artifact.targetContext)
  );
  return {
    currentArtifact,
    primaryArtifact: resolvePrimaryArtifact(currentArtifact, artifactChain),
    artifactChain,
    targetContext
  };
}

export function resolveLatestArtifactContext(
  artifacts: Readonly<Record<string, IHieArtifactRecord>>
): IHieResolvedArtifactContext {
  const currentArtifact = getLatestRelevantArtifact(artifacts);
  const artifactChain = resolveArtifactChain(artifacts, currentArtifact?.artifactId);
  const targetContext = mergeMcpTargetContexts(
    currentArtifact?.targetContext,
    ...artifactChain.map((artifact) => artifact.targetContext)
  );
  return {
    currentArtifact,
    primaryArtifact: resolvePrimaryArtifact(currentArtifact, artifactChain),
    artifactChain,
    targetContext
  };
}
