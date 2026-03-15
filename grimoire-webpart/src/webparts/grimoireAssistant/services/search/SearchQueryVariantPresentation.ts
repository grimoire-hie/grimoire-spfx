import type { ISearchQueryVariantInfo } from '../../models/IBlock';

export function getUserFacingSearchQueryVariants(
  variants?: ReadonlyArray<ISearchQueryVariantInfo>
): ISearchQueryVariantInfo[] {
  if (!variants || variants.length === 0) {
    return [];
  }

  return variants.filter((variant) => variant.kind !== 'semantic-rewrite');
}

export function formatSearchQueryVariantLabel(variant: Readonly<ISearchQueryVariantInfo>): string {
  switch (variant.kind) {
    case 'corrected':
      return `Corrected query: ${variant.query}`;
    case 'translation':
      return variant.language
        ? `Translated (${variant.language.toUpperCase()}): ${variant.query}`
        : `Translated: ${variant.query}`;
    case 'keyword-fallback':
      return `Extra keywords: ${variant.query}`;
    default:
      return variant.query;
  }
}

export function describeSearchQueryBreadth(
  variants?: ReadonlyArray<ISearchQueryVariantInfo>
): string | undefined {
  const userFacingVariants = getUserFacingSearchQueryVariants(variants);
  if (userFacingVariants.length === 0) {
    return undefined;
  }

  const correctedCount = userFacingVariants.filter((variant) => variant.kind === 'corrected').length;
  const translationCount = userFacingVariants.filter((variant) => variant.kind === 'translation').length;
  const keywordCount = userFacingVariants.filter((variant) => variant.kind === 'keyword-fallback').length;
  const parts: string[] = [];

  if (correctedCount > 0) {
    parts.push(correctedCount === 1 ? 'a corrected query' : `${correctedCount} corrected queries`);
  }
  if (translationCount > 0) {
    parts.push(translationCount === 1 ? 'a translated query' : `${translationCount} translated queries`);
  }
  if (keywordCount > 0) {
    parts.push(keywordCount === 1 ? 'extra keywords' : `${keywordCount} keyword expansions`);
  }

  if (parts.length === 0) {
    return undefined;
  }

  if (parts.length === 1) {
    return `Search breadth used ${parts[0]}.`;
  }

  return `Search breadth used ${parts.slice(0, -1).join(', ')} and ${parts[parts.length - 1]}.`;
}

export function formatSearchQueryBreadthLine(
  variants?: ReadonlyArray<ISearchQueryVariantInfo>
): string | undefined {
  const userFacingVariants = getUserFacingSearchQueryVariants(variants);
  if (userFacingVariants.length === 0) {
    return undefined;
  }

  return `Expanded search with: ${userFacingVariants.map((variant) => variant.query).join(' | ')}`;
}
