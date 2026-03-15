import {
  describeSearchQueryBreadth,
  formatSearchQueryBreadthLine,
  formatSearchQueryVariantLabel,
  getUserFacingSearchQueryVariants
} from './SearchQueryVariantPresentation';

describe('SearchQueryVariantPresentation', () => {
  it('hides semantic rewrites from user-facing query breadth details', () => {
    expect(getUserFacingSearchQueryVariants([
      { kind: 'semantic-rewrite', query: 'sharepoint framework' },
      { kind: 'keyword-fallback', query: 'spfx, sharepoint framework' }
    ])).toEqual([
      { kind: 'keyword-fallback', query: 'spfx, sharepoint framework' }
    ]);
  });

  it('formats query breadth labels in user-facing language', () => {
    expect(formatSearchQueryVariantLabel({
      kind: 'translation',
      query: 'cadre sharepoint',
      language: 'fr'
    })).toBe('Translated (FR): cadre sharepoint');
  });

  it('summarizes visible search breadth without leaking internal planner terms', () => {
    expect(describeSearchQueryBreadth([
      { kind: 'semantic-rewrite', query: 'sharepoint framework' },
      { kind: 'corrected', query: 'spfx' },
      { kind: 'keyword-fallback', query: 'spfx, sharepoint framework' }
    ])).toBe('Search breadth used a corrected query and extra keywords.');

    expect(formatSearchQueryBreadthLine([
      { kind: 'semantic-rewrite', query: 'sharepoint framework' },
      { kind: 'keyword-fallback', query: 'spfx, sharepoint framework' }
    ])).toBe('Expanded search with: spfx, sharepoint framework');
  });
});
