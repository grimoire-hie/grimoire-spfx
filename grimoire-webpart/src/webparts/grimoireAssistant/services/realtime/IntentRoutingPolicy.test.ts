jest.mock('../nano/NanoService', () => ({
  getNanoService: jest.fn()
}));

import {
  buildEnterpriseIntentRoutingGuidance,
  classifyFirstTurnRouting,
  classifyAssistantFirstTurnOutcome,
  ENTERPRISE_INTENT_REGRESSION_FIXTURES,
  getForcedFirstToolName,
  getFirstTurnRoutingObservation,
  getObservedFirstToolName,
  isClarificationResponse
} from './IntentRoutingPolicy';
import { getNanoService } from '../nano/NanoService';
import type { IProxyConfig } from '../../store/useGrimoireStore';

const proxyConfig: IProxyConfig = {
  proxyUrl: 'https://example.com/api',
  proxyApiKey: 'test-key',
  backend: 'reasoning',
  deployment: 'grimoire-reasoning',
  apiVersion: '2024-10-21'
};

describe('IntentRoutingPolicy guidance', () => {
  it('encodes enterprise-first precedence without the old blanket ambiguity fallback', () => {
    const guidance = buildEnterpriseIntentRoutingGuidance();

    expect(guidance).toContain('enterprise-first and action-oriented');
    expect(guidance).toContain('call `list_m365_servers`');
    expect(guidance).toContain('pass `focus`');
    expect(guidance).toContain('use `research_public_web`');
    expect(guidance).toContain('use `search_emails`');
    expect(guidance).toContain('use `search_people`');
    expect(guidance).toContain('use `search_sites`');
    expect(guidance).toContain('use `search_sharepoint`');
    expect(guidance).toContain('default to `search_sharepoint`');
    expect(guidance).toContain('"marketing"');
    expect(guidance).toContain('Clarify only when no plausible non-destructive capability family fits');
    expect(guidance).not.toContain("When the user's intent is ambiguous, ask them to clarify before choosing a tool");
  });
});

describe('IntentRoutingPolicy regression fixtures', () => {
  it.each(ENTERPRISE_INTENT_REGRESSION_FIXTURES)(
    'maps "$utterance" to $expectedRoute',
    ({ utterance, expectedRoute, expectedToolName, expectedToolArgs, expectsClarification }) => {
      const observation = getFirstTurnRoutingObservation(utterance);

      expect(observation).toBeDefined();
      expect(observation).toMatchObject({
        expectedRoute,
        expectedToolName,
        expectedToolArgs
      });
      expect(expectsClarification).toBe(false);
    }
  );

  it('returns undefined for generic search without domain hints', () => {
    const observation = getFirstTurnRoutingObservation('i am searching for informations about animals');
    expect(observation).toBeUndefined();
  });

  it('forces the capability overview tool for German capability queries', () => {
    const observation = getFirstTurnRoutingObservation('was kannst du alles?');

    expect(observation).toMatchObject({
      expectedRoute: 'list_m365_servers',
      expectedToolName: 'list_m365_servers'
    });
    expect(getForcedFirstToolName(observation)).toBe('list_m365_servers');
  });

  it('adds focused tool args for workload-specific capability questions', () => {
    const observation = getFirstTurnRoutingObservation('what can you do for SharePoint?');

    expect(observation).toMatchObject({
      expectedRoute: 'list_m365_servers',
      expectedToolName: 'list_m365_servers',
      expectedToolArgs: { focus: 'sharepoint' },
      capabilityFocus: 'sharepoint'
    });
  });

  it('does not force SharePoint search for explicit document-library browse requests', () => {
    const observation = getFirstTurnRoutingObservation('show me all the files in the document library Documents in the site copilot-test-cooking');

    expect(observation).toBeUndefined();
  });

  it('does not force SharePoint search for explicit personal OneDrive browse requests', () => {
    const observation = getFirstTurnRoutingObservation('show me my onedrive files');

    expect(observation).toBeUndefined();
  });

  it('does not force SharePoint search for explicit personal OneDrive filename searches', () => {
    const observation = getFirstTurnRoutingObservation('show me my onedrive files starting with nova');

    expect(observation).toBeUndefined();
  });
});

describe('IntentRoutingPolicy fast classification', () => {
  afterEach(() => {
    jest.restoreAllMocks();
    jest.resetAllMocks();
  });

  it.each(ENTERPRISE_INTENT_REGRESSION_FIXTURES)(
    'maps "$utterance" through fast classification',
    async ({ utterance, expectedRoute, expectedToolName }) => {
      (getNanoService as jest.Mock).mockReturnValue({
        classify: jest.fn().mockResolvedValue(JSON.stringify({
          route: expectedRoute,
          confidence: 0.93
        }))
      });

      const observation = await classifyFirstTurnRouting(utterance, proxyConfig);

      expect(observation).toMatchObject({
        expectedRoute,
        expectedToolName,
        source: 'fast',
        confidence: 0.93
      });
    }
  );

  it('normalizes fenced JSON from the fast classifier', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue('```json\n{"route":"list_m365_servers","confidence":0.91}\n```')
    });

    const observation = await classifyFirstTurnRouting('cosa puoi fare?', proxyConfig);

    expect(observation).toMatchObject({
      expectedRoute: 'list_m365_servers',
      expectedToolName: 'list_m365_servers',
      source: 'fast',
      confidence: 0.91
    });
  });

  it('falls back to heuristics when the fast classifier returns malformed JSON', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue('not json')
    });

    const observation = await classifyFirstTurnRouting('find emails about animals', proxyConfig);

    expect(observation).toMatchObject({
      expectedRoute: 'search_emails',
      expectedToolName: 'search_emails',
      source: 'heuristic',
      confidence: 1
    });
  });

  it('falls back to heuristics when the fast classifier is low confidence', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        route: 'search_people',
        confidence: 0.4
      }))
    });

    const observation = await classifyFirstTurnRouting('find emails about animals', proxyConfig);

    expect(observation).toMatchObject({
      expectedRoute: 'search_emails',
      expectedToolName: 'search_emails',
      source: 'heuristic',
      confidence: 1
    });
  });

  it('falls back to heuristics when the fast backend is unavailable', async () => {
    (getNanoService as jest.Mock).mockReturnValue(undefined);

    const observation = await classifyFirstTurnRouting('i am searching for documents about animals', proxyConfig);

    expect(observation).toMatchObject({
      expectedRoute: 'search_sharepoint',
      expectedToolName: 'search_sharepoint',
      source: 'heuristic',
      confidence: 1,
      isGenericEnterpriseSearch: false
    });
  });

  it('trusts Nano classification even for generic enterprise search phrasing', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        route: 'search_people',
        confidence: 0.72
      }))
    });

    const observation = await classifyFirstTurnRouting('search for nova marketing', proxyConfig);

    expect(observation).toMatchObject({
      expectedRoute: 'search_people',
      expectedToolName: 'search_people',
      source: 'fast',
      confidence: 0.72
    });
  });

  it('does not route contextual send-by-email follow-ups into email search', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        route: 'search_emails',
        confidence: 0.98
      }))
    });

    const observation = await classifyFirstTurnRouting('great results, i want to send the results by mail', proxyConfig);

    expect(observation).toBeUndefined();
    expect(getNanoService).not.toHaveBeenCalled();
  });

  it('keeps explicit document-library browse requests out of first-turn forced search routing', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        route: 'none',
        confidence: 0.95
      }))
    });

    const observation = await classifyFirstTurnRouting(
      'mostrami tutti i file nella raccolta documenti Documents del sito copilot-test-cooking',
      proxyConfig
    );

    expect(observation).toMatchObject({
      expectedRoute: 'none',
      source: 'fast',
      confidence: 0.95
    });
    expect(observation?.expectedToolName).toBeUndefined();
  });

  it('skips fast routing for explicit personal OneDrive browse requests', async () => {
    (getNanoService as jest.Mock).mockReturnValue({
      classify: jest.fn().mockResolvedValue(JSON.stringify({
        route: 'search_sharepoint',
        confidence: 0.95
      }))
    });

    const observation = await classifyFirstTurnRouting('show me my onedrive files', proxyConfig);

    expect(observation).toBeUndefined();
    expect(getNanoService).not.toHaveBeenCalled();
  });
});

describe('IntentRoutingPolicy assistant outcome classification', () => {
  it('detects the cross-domain clarification pattern that caused the regression', () => {
    const response = 'Hi Test User - I can help with that. Do you want me to search your Microsoft 365 (SharePoint/OneDrive) content, look the web up for general animal information, or search something else (e.g., images, PDFs, species lists, conservation reports)?';

    expect(isClarificationResponse(response)).toBe(true);
    expect(classifyAssistantFirstTurnOutcome(response)).toBe('clarification');
  });

  it('keeps normal answers out of the clarification bucket', () => {
    const response = 'I found five documents about animals in SharePoint.';

    expect(isClarificationResponse(response)).toBe(false);
    expect(classifyAssistantFirstTurnOutcome(response)).toBe('answer_only');
  });
});

describe('IntentRoutingPolicy observed tool name', () => {
  it('ignores status-only tool calls when reporting first-turn routing', () => {
    expect(getObservedFirstToolName(['set_expression', 'show_progress', 'search_sharepoint'])).toBe('search_sharepoint');
  });

  it('returns undefined when only non-decision tools were called', () => {
    expect(getObservedFirstToolName(['set_expression', 'show_progress'])).toBeUndefined();
  });
});
