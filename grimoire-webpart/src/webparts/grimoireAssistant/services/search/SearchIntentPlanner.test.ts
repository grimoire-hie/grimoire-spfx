jest.mock('../logging/LogService', () => ({
  logService: {
    info: jest.fn(),
    warning: jest.fn(),
    error: jest.fn(),
    debug: jest.fn()
  }
}));

jest.mock('../nano/NanoService', () => ({
  getNanoService: jest.fn()
}));

import { getNanoService } from '../nano/NanoService';
import { logService } from '../logging/LogService';
import { SearchIntentPlanner } from './SearchIntentPlanner';

describe('SearchIntentPlanner', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('falls back to heuristic planning when Nano is unavailable', async () => {
    (getNanoService as jest.Mock).mockReturnValue(undefined);

    const planner = new SearchIntentPlanner();
    const plan = await planner.plan('rapport budget', { userLanguage: 'fr' });

    expect(plan.rawQuery).toBe('rapport budget');
    expect(plan.queryLanguage).toBe('fr');
    expect(plan.sharePointLexicalQuery).toBeUndefined();
    expect(plan.keywordFallbackQuery).toBeUndefined();
    expect(plan.translationFallbackLanguage).toBe('en');
  });

  it('keeps only high-confidence corrections and captures translation fallback', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        queryLanguage: 'en',
        semanticRewriteQuery: 'finance budget report',
        semanticRewriteConfidence: 0.9,
        sharePointLexicalQuery: 'budget finance',
        sharePointLexicalConfidence: 0.95,
        correctedQuery: 'budget report',
        correctionConfidence: 0.93,
        translationFallbackQuery: 'rapport budget',
        translationFallbackLanguage: 'fr',
        keywordFallbackQuery: 'budget report finance'
      }))
    });

    const planner = new SearchIntentPlanner();
    const plan = await planner.plan('budegt report', {
      proxyConfig: {
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      userLanguage: 'fr'
    });

    expect(plan.semanticRewriteQuery).toBe('finance budget report');
    expect(plan.semanticRewriteConfidence).toBe(0.9);
    expect(plan.sharePointLexicalQuery).toBe('budget finance');
    expect(plan.sharePointLexicalConfidence).toBe(0.95);
    expect(plan.correctedQuery).toBe('budget report');
    expect(plan.translationFallbackQuery).toBe('rapport budget');
    expect(plan.translationFallbackLanguage).toBe('fr');
    expect(plan.keywordFallbackQuery).toBe('budget report finance');
    expect(plan.usedCorrection).toBe(true);
    expect(plan.usedTranslation).toBe(true);
  });

  it('keeps a planner lexical query even when it matches the raw query', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        queryLanguage: 'en',
        sharePointLexicalQuery: 'spfx',
        sharePointLexicalConfidence: 0.7
      }))
    });

    const planner = new SearchIntentPlanner();
    const plan = await planner.plan('spfx', {
      proxyConfig: {
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      userLanguage: 'en'
    });

    expect(plan.sharePointLexicalQuery).toBe('spfx');
    expect(plan.sharePointLexicalConfidence).toBe(0.7);
  });

  it('logs rejected correction candidates with confidence and reason', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        queryLanguage: 'en',
        correctedQuery: 'animals',
        correctionConfidence: 0.78,
        keywordFallbackQuery: 'animals'
      }))
    });

    const planner = new SearchIntentPlanner();
    const plan = await planner.plan('aminals', {
      proxyConfig: {
        proxyUrl: 'https://example.com/api',
        proxyApiKey: 'test',
        backend: 'reasoning',
        deployment: 'grimoire-reasoning',
        apiVersion: '2024-10-21'
      },
      userLanguage: 'en'
    });

    expect(plan.correctedQuery).toBeUndefined();
    expect(plan.keywordFallbackQuery).toBe('animals');

    const plannerLogPayload = (logService.info as jest.Mock).mock.calls.find(
      (call) => call[0] === 'search' && typeof call[2] === 'string'
    )?.[2];
    expect(plannerLogPayload).toBeTruthy();

    const parsedPayload = JSON.parse(plannerLogPayload as string) as {
      correctedCandidate?: string | null;
      correctionConfidence?: number;
      correctionDecision?: { accepted?: boolean; reason?: string; threshold?: number };
      keywordDecision?: { accepted?: boolean; reason?: string };
    };

    expect(parsedPayload.correctedCandidate).toBe('animals');
    expect(parsedPayload.correctionConfidence).toBe(0.78);
    expect(parsedPayload.correctionDecision).toMatchObject({
      accepted: false,
      reason: 'below_confidence_threshold',
      threshold: 0.8
    });
    expect(parsedPayload.keywordDecision).toMatchObject({
      accepted: true,
      reason: 'accepted'
    });
  });

});
