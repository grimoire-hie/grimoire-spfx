import type { BlockType } from '../../models/IBlock';
import type { HieArtifactKind } from './HIETypes';

const LOOKUP_BLOCK_TYPES: ReadonlySet<BlockType> = new Set<BlockType>([
  'site-info',
  'user-card',
  'list-items',
  'permissions-view',
  'activity-feed',
  'chart',
  'document-library',
  'markdown'
]);

function startsWithTitlePrefix(title: string | undefined, prefix: string): boolean {
  return !!title && title.trim().toLowerCase().startsWith(prefix.toLowerCase());
}

export function isHieArtifactKind(value: string | undefined): value is HieArtifactKind {
  return value === 'block'
    || value === 'summary'
    || value === 'preview'
    || value === 'lookup'
    || value === 'error'
    || value === 'recap'
    || value === 'form'
    || value === 'share'
    || value === 'generic';
}

export function resolveArtifactKindFromBlockContext(
  blockType: BlockType,
  options?: {
    title?: string;
    sourceTaskKind?: string;
    originTool?: string;
  }
): HieArtifactKind {
  if (options?.originTool?.startsWith('block-recap:')) {
    return 'recap';
  }

  if (blockType === 'error') {
    return 'error';
  }

  if (blockType === 'file-preview') {
    return 'preview';
  }

  if (options?.sourceTaskKind === 'summarize' || startsWithTitlePrefix(options?.title, 'Summary:')) {
    return 'summary';
  }

  if (LOOKUP_BLOCK_TYPES.has(blockType) || blockType === 'info-card') {
    return 'lookup';
  }

  return 'block';
}
