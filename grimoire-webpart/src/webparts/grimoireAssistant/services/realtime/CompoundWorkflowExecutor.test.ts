import {
  executeCompoundWorkflowPlan,
  parseCompoundWorkflowPlannerResponse,
  planCompoundWorkflow,
  shouldConsiderCompoundWorkflow,
  type ICompoundWorkflowPlan
} from './CompoundWorkflowExecutor';
import { COMPOUND_WORKFLOW_PLANNER_SYSTEM_PROMPT } from '../../config/promptCatalog';
import { createBlock } from '../../models/IBlock';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { BlockRecapService } from '../recap/BlockRecapService';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import * as McpExecutionAdapter from '../mcp/McpExecutionAdapter';

function buildPlan(
  familyId: ICompoundWorkflowPlan['familyId'],
  query: string,
  selectionHint: ICompoundWorkflowPlan['slots']['selectionHint'] = 'none'
): ICompoundWorkflowPlan {
  const stepMap: Record<ICompoundWorkflowPlan['familyId'], Array<{ id: string; kind: ICompoundWorkflowPlan['steps'][number]['kind']; label: string }>> = {
    search_sharepoint_recap: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' }
    ],
    search_sharepoint_recap_email: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ],
    search_sharepoint_recap_teams_chat: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' },
      { id: 'share-teams-chat', kind: 'share-teams-chat', label: 'Open Teams chat share' }
    ],
    search_sharepoint_recap_teams_channel: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'recap-results', kind: 'recap-results', label: 'Summarize visible results' },
      { id: 'share-teams-channel', kind: 'share-teams-channel', label: 'Open Teams channel share' }
    ],
    search_sharepoint_summarize_document: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'summarize-document', kind: 'summarize-document', label: 'Summarize the selected document' }
    ],
    search_sharepoint_summarize_document_email: [
      { id: 'search', kind: 'search', label: 'Search SharePoint' },
      { id: 'summarize-document', kind: 'summarize-document', label: 'Summarize the selected document' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ],
    search_people_email: [
      { id: 'find-person', kind: 'find-person', label: 'Find the person' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ],
    search_emails_summarize_email: [
      { id: 'search', kind: 'search', label: 'Search emails' },
      { id: 'summarize-email', kind: 'summarize-email', label: 'Summarize the selected email' }
    ],
    visible_recap_reply_all_mail_discussion: [
      { id: 'resolve-mail-thread', kind: 'resolve-mail-thread', label: 'Resolve the mail discussion thread' },
      { id: 'compose-email', kind: 'compose-email', label: 'Open email draft' }
    ]
  };

  return {
    shouldPlan: true,
    familyId,
    confidence: 0.96,
    slots: {
      query,
      selectionHint
    },
    steps: stepMap[familyId].map((step) => ({ ...step })),
    label: familyId,
    userText: query
  };
}

function buildFormBlock(title: string): ReturnType<typeof createBlock> {
  return createBlock('form', title, {
    kind: 'form',
    preset: 'email-compose',
    fields: [],
    submissionTarget: {
      toolName: 'SendEmailWithAttachments',
      serverId: 'mcp_MailTools',
      staticArgs: {}
    },
    status: 'editing'
  });
}

function parseJsonArg(arg: unknown): Record<string, unknown> {
  return JSON.parse(String(arg || '{}')) as Record<string, unknown>;
}

function buildGetMessageExecution(payload: Record<string, unknown>): Record<string, unknown> {
  return {
    success: true,
    serverId: 'mcp_MailTools',
    serverName: 'Outlook Mail',
    serverUrl: 'https://example.invalid/mail',
    sessionId: 'session-mail',
    realToolName: 'GetMessage',
    requiredFields: ['id'],
    schemaProps: {},
    normalizedArgs: { id: 'mail-exact', bodyPreviewOnly: true },
    resolvedArgs: { id: 'mail-exact', bodyPreviewOnly: true },
    targetSource: 'unknown',
    recoverySteps: [],
    mcpResult: {
      success: true,
      content: [{
        type: 'text',
        text: JSON.stringify({ payload })
      }]
    },
    trace: {
      toolName: 'GetMessage',
      rawArgs: { id: 'mail-exact', bodyPreviewOnly: true },
      recoverySteps: [],
      targetSource: 'unknown'
    }
  };
}

afterEach(() => {
  jest.restoreAllMocks();
  hybridInteractionEngine.reset();
  delete window.__GRIMOIRE_RUNTIME_TUNING__;
  useGrimoireStore.setState({
    blocks: [],
    transcript: [],
    activeActionBlockId: undefined,
    selectedActionIndices: [],
    proxyConfig: undefined
  });
});

describe('CompoundWorkflowExecutor planner', () => {
  it('only considers likely summarize-and-share turns and excludes ordinary single-step prompts', () => {
    expect(shouldConsiderCompoundWorkflow('search for spfx, summarize the results and send by email')).toBe(true);
    expect(shouldConsiderCompoundWorkflow('search for spfx, summarize and send by email')).toBe(true);
    expect(shouldConsiderCompoundWorkflow('search for spfx')).toBe(false);
    expect(shouldConsiderCompoundWorkflow('hello there')).toBe(false);
    expect(shouldConsiderCompoundWorkflow('   ')).toBe(false);
    expect(shouldConsiderCompoundWorkflow('search github for spfx and summarize it')).toBe(false);
  });

  it('parses a valid legacy planner response into a workflow plan', () => {
    const plan = parseCompoundWorkflowPlannerResponse(
      JSON.stringify({
        shouldPlan: true,
        familyId: 'search_sharepoint_recap_email',
        confidence: 0.94,
        slots: {
          query: 'spfx',
          selectionHint: 'none'
        }
      }),
      'search for spfx, summarize the results and send by email'
    );

    expect(plan).toMatchObject({
      familyId: 'search_sharepoint_recap_email',
      slots: {
        query: 'spfx',
        selectionHint: 'none'
      }
    });
    expect(plan?.steps).toHaveLength(3);
  });

  it('parses a slot-based planner response into a workflow plan', () => {
    const plan = parseCompoundWorkflowPlannerResponse(
      JSON.stringify({
        p: 1,
        d: 'sp',
        q: 'spfx',
        t: 'r',
        a: 'e',
        s: 'n',
        c: 0.94
      }),
      'search for spfx, summarize the results and send by email'
    );

    expect(plan).toMatchObject({
      familyId: 'search_sharepoint_recap_email',
      slots: {
        query: 'spfx',
        selectionHint: 'none'
      }
    });
    expect(plan?.steps).toHaveLength(3);
  });

  it('derives workflow families deterministically from planner slots', () => {
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'sp', q: 'spfx', t: 'r', a: 'e', s: 'n', c: 0.95 }),
      'search for spfx, summarize the results and send by email'
    )?.familyId).toBe('search_sharepoint_recap_email');
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'sp', q: 'spfx', t: 'u', a: 'e', s: 'n', c: 0.95 }),
      'search for spfx, summarize and send by email'
    )?.familyId).toBe('search_sharepoint_recap_email');
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'sp', q: 'spfx', t: 'r', a: 'tc', s: 'n', c: 0.95 }),
      'search for spfx, summarize the results and share to Teams'
    )?.familyId).toBe('search_sharepoint_recap_teams_chat');
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'sp', q: 'spfx', t: 'i', a: 'e', s: 't', c: 0.95 }),
      'search for spfx, summarize the top document and send by email'
    )?.familyId).toBe('search_sharepoint_summarize_document_email');
  });

  it('rejects unsupported slot combinations', () => {
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'pe', q: 'john doe', t: 'u', a: 'e', s: 'n', c: 0.95 }),
      'find john doe and draft an email'
    )).toBeUndefined();
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'em', q: 'budget', t: 'i', a: 'n', s: 't', c: 0.95 }),
      'find emails about budget and summarize the top one'
    )).toBeUndefined();
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'em', q: 'budget', t: 'r', a: 'n', s: 'n', c: 0.95 }),
      'find emails about budget and summarize the results'
    )).toBeUndefined();
    expect(parseCompoundWorkflowPlannerResponse(
      JSON.stringify({ p: 1, d: 'xx', q: 'spfx', t: 'r', a: 'e', s: 'n', c: 0.95 }),
      'search for spfx, summarize the results and send by email'
    )).toBeUndefined();
  });

  it('falls back when the planner response is malformed or low-confidence', async () => {
    const malformed = await planCompoundWorkflow(
      'search for spfx, summarize and send by email',
      undefined,
      {
        classify: jest.fn().mockResolvedValue('not-json')
      }
    );
    expect(malformed).toBeUndefined();

    const lowConfidence = await planCompoundWorkflow(
      'search for spfx, summarize and send by email',
      undefined,
      {
        classify: jest.fn().mockResolvedValue(JSON.stringify({
          p: 1,
          d: 'sp',
          q: 'spfx',
          t: 'r',
          a: 'e',
          s: 'n',
          c: 0.41
        }))
      }
    );
    expect(lowConfidence).toBeUndefined();
  });

  it('uses the centralized prompt and runtime tuning for compound planning', async () => {
    window.__GRIMOIRE_RUNTIME_TUNING__ = {
      nano: {
        compoundWorkflowPlannerTimeoutMs: 4321,
        compoundWorkflowPlannerMaxTokens: 77,
        compoundWorkflowPlannerConfidenceThreshold: 0.88
      }
    };
    const classify = jest.fn().mockResolvedValue(JSON.stringify({
      p: 1,
      d: 'sp',
      q: 'spfx',
      t: 'r',
      a: 'e',
      s: 'n',
      c: 0.92
    }));

    const plan = await planCompoundWorkflow(
      'search for spfx, summarize and send by email',
      undefined,
      { classify }
    );

    expect(classify).toHaveBeenCalledWith(
      COMPOUND_WORKFLOW_PLANNER_SYSTEM_PROMPT,
      'search for spfx, summarize and send by email',
      4321,
      77
    );
    expect(plan?.familyId).toBe('search_sharepoint_recap_email');
    expect(getRuntimeTuningConfig().nano.compoundWorkflowPlannerConfidenceThreshold).toBe(0.88);
  });

  it('allows multilingual explicit and bare summarize phrases through the planner path', async () => {
    const classify = jest.fn()
      .mockResolvedValueOnce(JSON.stringify({
        p: 1,
        d: 'sp',
        q: 'spfx',
        t: 'r',
        a: 'e',
        s: 'n',
        c: 0.93
      }))
      .mockResolvedValueOnce(JSON.stringify({
        p: 1,
        d: 'sp',
        q: 'spfx',
        t: 'r',
        a: 'e',
        s: 'n',
        c: 0.93
      }))
      .mockResolvedValueOnce(JSON.stringify({
        p: 1,
        d: 'sp',
        q: 'spfx',
        t: 'r',
        a: 'e',
        s: 'n',
        c: 0.93
      }))
      .mockResolvedValueOnce(JSON.stringify({
        p: 1,
        d: 'sp',
        q: 'spfx',
        t: 'r',
        a: 'tc',
        s: 'n',
        c: 0.93
      }))
      .mockResolvedValueOnce(JSON.stringify({
        p: 1,
        d: 'sp',
        q: 'spfx',
        t: 'r',
        a: 'e',
        s: 'n',
        c: 0.93
      }));

    await expect(planCompoundWorkflow('cerca spfx, riassumi i risultati e invia per email', undefined, { classify }))
      .resolves.toMatchObject({ familyId: 'search_sharepoint_recap_email' });
    await expect(planCompoundWorkflow('cerca spfx, riassumi e invia per email', undefined, { classify }))
      .resolves.toMatchObject({ familyId: 'search_sharepoint_recap_email' });
    await expect(planCompoundWorkflow('busca spfx, resume los resultados y envialos por correo', undefined, { classify }))
      .resolves.toMatchObject({ familyId: 'search_sharepoint_recap_email' });
    await expect(planCompoundWorkflow('cerca spfx, riassumi i risultati e condividi su Teams', undefined, { classify }))
      .resolves.toMatchObject({ familyId: 'search_sharepoint_recap_teams_chat' });
    await expect(planCompoundWorkflow('recherche spfx, résume les résultats et envoie-les par e-mail', undefined, { classify }))
      .resolves.toMatchObject({ familyId: 'search_sharepoint_recap_email' });

    expect(classify).toHaveBeenCalledTimes(5);
  });
});

describe('CompoundWorkflowExecutor execution', () => {
  it('runs a search -> recap -> email workflow and leaves the compose panel visible', async () => {
    jest.spyOn(BlockRecapService.prototype, 'generate').mockResolvedValue('SPFx recap body');

    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        switch (funcName) {
          case 'search_sharepoint': {
            const block = createBlock('search-results', 'Search: SPFx', {
              kind: 'search-results',
              query: String(args.query || 'spfx'),
              totalCount: 2,
              source: 'copilot-search',
              results: [
                {
                  title: 'SPFx Overview.docx',
                  summary: 'Overview',
                  url: 'https://tenant.sharepoint.com/sites/dev/SPFx-Overview.docx',
                  sources: ['copilot-search']
                },
                {
                  title: 'SPFx Deep Dive.pdf',
                  summary: 'Deep dive',
                  url: 'https://tenant.sharepoint.com/sites/dev/SPFx-Deep-Dive.pdf',
                  sources: ['copilot-search']
                }
              ]
            });
            store.pushBlock(block);
            return JSON.stringify({ success: true, displayedResults: 2 });
          }
          case 'show_compose_form': {
            store.pushBlock(buildFormBlock(String(args.title || 'Compose')));
            return JSON.stringify({ success: true });
          }
          default:
            return JSON.stringify({ success: true });
        }
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('search_sharepoint_recap_email', 'spfx'),
      callbacks
    );

    expect(assistantText).toContain('opened an email draft');
    expect(callbacks.onFunctionCall).toHaveBeenNthCalledWith(
      1,
      'compound-search_sharepoint_recap_email-search',
      'search_sharepoint',
      { query: 'spfx' }
    );
    expect(callbacks.onFunctionCall).toHaveBeenNthCalledWith(
      2,
      'compound-search_sharepoint_recap_email-compose-email',
      'show_compose_form',
      expect.objectContaining({
        preset: 'email-compose',
        title: 'Share by Email'
      })
    );
    const composeArgs = callbacks.onFunctionCall.mock.calls[1][2] as Record<string, unknown>;
    expect(parseJsonArg(composeArgs.prefill_json)).toMatchObject({
      subject: 'Recap: Search: SPFx',
      body: 'SPFx recap body'
    });

    const blocks = useGrimoireStore.getState().blocks;
    expect(blocks.map((block) => block.type)).toEqual([
      'progress-tracker',
      'search-results',
      'info-card',
      'form'
    ]);

    const tracker = blocks[0].data as { status: string; steps?: Array<{ status: string }> };
    expect(tracker.status).toBe('complete');
    expect(tracker.steps?.every((step) => step.status === 'complete')).toBe(true);
  });

  it('passes the document summary text into the email draft and preserves attachment args', async () => {
    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        switch (funcName) {
          case 'search_sharepoint': {
            store.pushBlock(createBlock('search-results', 'Search: SPFx', {
              kind: 'search-results',
              query: String(args.query || 'spfx'),
              totalCount: 1,
              source: 'copilot-search',
              results: [
                {
                  title: 'SPFx Overview.docx',
                  summary: 'Overview',
                  url: 'https://tenant.sharepoint.com/sites/dev/SPFx-Overview.docx',
                  sources: ['copilot-search']
                }
              ]
            }));
            return JSON.stringify({ success: true, displayedResults: 1 });
          }
          case 'read_file_content':
            return JSON.stringify({ success: true, content: 'SPFx file summary body' });
          case 'show_info_card': {
            store.pushBlock(createBlock('info-card', String(args.heading || 'Summary'), {
              kind: 'info-card',
              heading: String(args.heading || 'Summary'),
              body: String(args.body || ''),
              icon: String(args.icon || 'AlignLeft')
            }));
            return JSON.stringify({ success: true });
          }
          case 'show_compose_form': {
            store.pushBlock(buildFormBlock(String(args.title || 'Compose')));
            return JSON.stringify({ success: true });
          }
          default:
            return JSON.stringify({ success: true });
        }
      })
    };

    await executeCompoundWorkflowPlan(
      buildPlan('search_sharepoint_summarize_document_email', 'spfx'),
      callbacks
    );

    const composeArgs = callbacks.onFunctionCall.mock.calls[3][2] as Record<string, unknown>;
    expect(parseJsonArg(composeArgs.prefill_json)).toMatchObject({
      subject: 'Summary: SPFx Overview.docx',
      body: 'SPFx file summary body'
    });
    expect(parseJsonArg(composeArgs.static_args_json)).toMatchObject({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/SPFx-Overview.docx'],
      fileOrFolderUrl: 'https://tenant.sharepoint.com/sites/dev/SPFx-Overview.docx',
      fileOrFolderName: 'SPFx Overview.docx'
    });
  });

  it('uses recap text as the Teams share content', async () => {
    jest.spyOn(BlockRecapService.prototype, 'generate').mockResolvedValue('SPFx recap body');

    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        if (funcName === 'search_sharepoint') {
          store.pushBlock(createBlock('search-results', 'Search: SPFx', {
            kind: 'search-results',
            query: String(args.query || 'spfx'),
            totalCount: 1,
            source: 'copilot-search',
            results: [
              {
                title: 'SPFx Overview.docx',
                summary: 'Overview',
                url: 'https://tenant.sharepoint.com/sites/dev/SPFx-Overview.docx',
                sources: ['copilot-search']
              }
            ]
          }));
          return JSON.stringify({ success: true, displayedResults: 1 });
        }
        if (funcName === 'show_compose_form') {
          store.pushBlock(createBlock('form', String(args.title || 'Compose'), {
            kind: 'form',
            preset: 'share-teams-chat',
            fields: [],
            submissionTarget: {
              toolName: 'PostToTeamsChat',
              serverId: 'mcp_TeamsTools',
              staticArgs: {}
            },
            status: 'editing'
          }));
          return JSON.stringify({ success: true });
        }
        return JSON.stringify({ success: true });
      })
    };

    await executeCompoundWorkflowPlan(
      buildPlan('search_sharepoint_recap_teams_chat', 'spfx'),
      callbacks
    );

    const shareArgs = callbacks.onFunctionCall.mock.calls[1][2] as Record<string, unknown>;
    expect(parseJsonArg(shareArgs.prefill_json)).toMatchObject({
      topic: 'Recap: Search: SPFx',
      content: 'SPFx recap body'
    });
  });

  it('fails the share step when no recap or summary artifact exists', async () => {
    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        switch (funcName) {
          case 'search_sharepoint': {
            store.pushBlock(createBlock('search-results', 'Search: SPFx', {
              kind: 'search-results',
              query: String(args.query || 'spfx'),
              totalCount: 1,
              source: 'copilot-search',
              results: [
                {
                  title: 'SPFx Overview.docx',
                  summary: 'Overview',
                  url: 'https://tenant.sharepoint.com/sites/dev/SPFx-Overview.docx',
                  sources: ['copilot-search']
                }
              ]
            }));
            return JSON.stringify({ success: true, displayedResults: 1 });
          }
          case 'read_file_content':
            return JSON.stringify({ success: true, content: 'SPFx file summary body' });
          case 'show_info_card':
            return JSON.stringify({ success: true });
          case 'show_compose_form':
            return JSON.stringify({ success: true });
          default:
            return JSON.stringify({ success: true });
        }
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('search_sharepoint_summarize_document_email', 'spfx'),
      callbacks
    );

    expect(assistantText).toContain('visible recap or summary');
    expect(callbacks.onFunctionCall).toHaveBeenCalledTimes(3);
  });

  it('stops gracefully on ambiguous single-document requests and preserves the visible search results', async () => {
    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        if (funcName === 'search_sharepoint') {
          useGrimoireStore.getState().pushBlock(createBlock('search-results', 'Search: Budget', {
            kind: 'search-results',
            query: String(args.query || 'budget'),
            totalCount: 2,
            source: 'copilot-search',
            results: [
              {
                title: 'Budget-Q1.xlsx',
                summary: 'Quarter 1 budget',
                url: 'https://tenant.sharepoint.com/sites/finance/Budget-Q1.xlsx',
                sources: ['copilot-search']
              },
              {
                title: 'Budget-Q2.xlsx',
                summary: 'Quarter 2 budget',
                url: 'https://tenant.sharepoint.com/sites/finance/Budget-Q2.xlsx',
                sources: ['copilot-search']
              }
            ]
          }));
          return JSON.stringify({ success: true, displayedResults: 2 });
        }
        if (funcName === 'read_file_content') {
          return JSON.stringify({ success: true, content: 'This should not run.' });
        }
        return JSON.stringify({ success: true });
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('search_sharepoint_summarize_document', 'budget'),
      callbacks
    );

    expect(assistantText).toContain('Pick one');
    expect(callbacks.onFunctionCall).toHaveBeenCalledTimes(1);

    const blocks = useGrimoireStore.getState().blocks;
    expect(blocks.map((block) => block.type)).toEqual(['progress-tracker', 'search-results']);
    const tracker = blocks[0].data as { status: string; detail?: string };
    expect(tracker.status).toBe('error');
    expect(tracker.detail).toContain('Multiple documents matched');
  });

  it('auto-picks a strong exact mail subject match and opens an email draft with recipients from the selected email', async () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Project Nova launch recap summary.'
    });
    jest.spyOn(McpExecutionAdapter, 'executeCatalogMcpTool').mockResolvedValue(buildGetMessageExecution({
      from: {
        emailAddress: {
          address: 'alice.smith@contoso.onmicrosoft.com'
        }
      },
      toRecipients: [
        { emailAddress: { address: 'test.user@contoso.onmicrosoft.com' } },
        { emailAddress: { address: 'bob.jones@contoso.onmicrosoft.com' } }
      ],
      ccRecipients: [
        { emailAddress: { address: 'carol.wilson@contoso.onmicrosoft.com' } },
        { emailAddress: { address: 'test.user@contoso.onmicrosoft.com' } }
      ]
    }) as never);
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:00:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id,
      userContext: {
        displayName: 'Test User',
        email: 'test.user@contoso.onmicrosoft.com',
        loginName: 'test.user@contoso.onmicrosoft.com',
        resolvedLanguage: 'en',
        currentWebTitle: 'Project Nova',
        currentWebUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        currentSiteTitle: 'Project Nova',
        currentSiteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova'
      },
      proxyConfig: {
        proxyUrl: 'https://proxy.invalid',
        proxyApiKey: 'test-key',
        backend: 'backend',
        deployment: 'deployment',
        apiVersion: '2024-10-21'
      },
      mcpEnvironmentId: 'env-test'
    });

    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        if (funcName === 'search_emails') {
          expect(String(args.query || '')).toContain('subject is exactly or very close');
          store.pushBlock(createBlock('markdown', 'MCP: SearchMessages', {
            kind: 'markdown',
            content: [
              '1. **Subject:** Nova Launch Blockers',
              '   **From:** Anna Mueller',
              '   **Date:** Today',
              '   **Preview:** Exact subject match.',
              '',
              '2. **Subject:** Nova Site',
              '   **From:** Test User',
              '   **Date:** Today',
              '   **Preview:** Mentions the blockers document in the body.'
            ].join('\n'),
            itemIds: {
              1: 'mail-exact',
              2: 'mail-generic'
            }
          }));
          return JSON.stringify({ success: true, count: 2 });
        }
        if (funcName === 'show_compose_form') {
          store.pushBlock(buildFormBlock(String(args.title || 'Compose')));
          return JSON.stringify({ success: true });
        }
        throw new Error(`Unexpected tool call: ${funcName}`);
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('visible_recap_reply_all_mail_discussion', 'Nova Launch Blockers'),
      callbacks
    );

    expect(assistantText).toContain('email draft');
    expect(callbacks.onFunctionCall).toHaveBeenCalledTimes(2);
    const composeArgs = callbacks.onFunctionCall.mock.calls[1][2] as Record<string, unknown>;
    expect(composeArgs).toMatchObject({
      preset: 'email-compose',
      title: 'Email Mail Participants'
    });
    expect(parseJsonArg(composeArgs.prefill_json)).toMatchObject({
      to: 'alice.smith@contoso.onmicrosoft.com, bob.jones@contoso.onmicrosoft.com',
      cc: 'carol.wilson@contoso.onmicrosoft.com',
      subject: 'Nova Launch Recap',
      body: 'Project Nova launch recap summary.'
    });
    expect(parseJsonArg(composeArgs.static_args_json)).toMatchObject({
      skipSessionHydration: true
    });
  });

  it('opens a chooser instead of auto-picking when multiple plausible mail subjects match', async () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Project Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:02:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id
    });

    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, _args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        if (funcName === 'search_emails') {
          store.pushBlock(createBlock('markdown', 'MCP: SearchMessages', {
            kind: 'markdown',
            content: [
              '1. **Subject:** Re: Nova Launch Blockers',
              '   **From:** Anna Mueller',
              '   **Date:** Today',
              '',
              '2. **Subject:** FW: Nova Launch Blockers',
              '   **From:** Bruno Meier',
              '   **Date:** Today'
            ].join('\n'),
            itemIds: {
              1: 'mail-re',
              2: 'mail-fw'
            }
          }));
          return JSON.stringify({ success: true, count: 2 });
        }
        if (funcName === 'show_selection_list') {
          return JSON.stringify({ success: true });
        }
        throw new Error(`Unexpected tool call: ${funcName}`);
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('visible_recap_reply_all_mail_discussion', 'Nova Launch Blockers'),
      callbacks
    );

    expect(assistantText).toContain('Choose one in the panel');
    expect(callbacks.onFunctionCall).toHaveBeenNthCalledWith(
      2,
      'compound-visible_recap_reply_all_mail_discussion-choose-mail-thread',
      'show_selection_list',
      expect.objectContaining({
        prompt: 'Choose the email to use for recipients.',
        multi_select: 'false'
      })
    );
  });

  it('falls back after generic subject misses and then resolves a real subject match', async () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Project Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:02:30.000Z') }
      ],
      activeActionBlockId: recapBlock.id,
      userContext: {
        displayName: 'Test User',
        email: 'test.user@contoso.onmicrosoft.com',
        loginName: 'test.user@contoso.onmicrosoft.com',
        resolvedLanguage: 'en',
        currentWebTitle: 'Project Nova',
        currentWebUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        currentSiteTitle: 'Project Nova',
        currentSiteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova'
      },
      proxyConfig: {
        proxyUrl: 'https://proxy.invalid',
        proxyApiKey: 'test-key',
        backend: 'backend',
        deployment: 'deployment',
        apiVersion: '2024-10-21'
      },
      mcpEnvironmentId: 'env-test'
    });

    let searchCallCount = 0;
    jest.spyOn(McpExecutionAdapter, 'executeCatalogMcpTool').mockResolvedValue(buildGetMessageExecution({
      from: {
        emailAddress: {
          address: 'alice.smith@contoso.onmicrosoft.com'
        }
      },
      toRecipients: [
        { emailAddress: { address: 'test.user@contoso.onmicrosoft.com' } },
        { emailAddress: { address: 'bob.jones@contoso.onmicrosoft.com' } }
      ],
      ccRecipients: []
    }) as never);
    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        const store = useGrimoireStore.getState();
        if (funcName === 'search_emails') {
          searchCallCount++;
          if (searchCallCount === 1) {
            store.pushBlock(createBlock('markdown', 'MCP: SearchMessages', {
              kind: 'markdown',
              content: [
                '1. **Subject:** Nova Site',
                '   **From:** Test User',
                '   **Date:** Today',
                '',
                '2. **Subject:** WG: Nova Docs',
                '   **From:** Test User',
                '   **Date:** Today'
              ].join('\n'),
              itemIds: {
                1: 'mail-generic-1',
                2: 'mail-generic-2'
              }
            }));
          } else {
            expect(String(args.query || '')).toContain('Find emails about');
            store.pushBlock(createBlock('markdown', 'MCP: SearchMessages', {
              kind: 'markdown',
              content: [
                '1. **Subject:** Re: Nova Launch Blockers',
                '   **From:** Anna Mueller',
                '   **Date:** Today'
              ].join('\n'),
              itemIds: {
                1: 'mail-fallback'
              }
            }));
          }
          return JSON.stringify({ success: true, count: 2 });
        }
        if (funcName === 'show_compose_form') {
          return JSON.stringify({ success: true });
        }
        throw new Error(`Unexpected tool call: ${funcName}`);
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('visible_recap_reply_all_mail_discussion', 'Nova Launch Blockers'),
      callbacks
    );

    expect(assistantText).toContain('email draft');
    expect(searchCallCount).toBe(2);
    expect(callbacks.onFunctionCall).toHaveBeenNthCalledWith(
      3,
      'compound-visible_recap_reply_all_mail_discussion-compose-email',
      'show_compose_form',
      expect.objectContaining({
        preset: 'email-compose'
      })
    );
  });

  it('stops clearly when no visible recap is available for the mail-recipient workflow', async () => {
    useGrimoireStore.setState({
      blocks: [],
      transcript: [],
      activeActionBlockId: undefined
    });

    const callbacks = {
      onFunctionCall: jest.fn(async () => JSON.stringify({ success: true }))
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('visible_recap_reply_all_mail_discussion', 'Nova Launch Blockers'),
      callbacks
    );

    expect(assistantText).toContain('Create the recap first');
    expect(callbacks.onFunctionCall).not.toHaveBeenCalled();
  });

  it('uses the current HIE mail selection and skips email search for mail-recipient workflows', async () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Project Nova launch recap summary.'
    });
    jest.spyOn(McpExecutionAdapter, 'executeCatalogMcpTool').mockResolvedValue(buildGetMessageExecution({
      from: {
        emailAddress: {
          address: 'alice.smith@contoso.onmicrosoft.com'
        }
      },
      toRecipients: [
        { emailAddress: { address: 'test.user@contoso.onmicrosoft.com' } },
        { emailAddress: { address: 'bob.jones@contoso.onmicrosoft.com' } }
      ],
      ccRecipients: [
        { emailAddress: { address: 'carol.wilson@contoso.onmicrosoft.com' } }
      ]
    }) as never);
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:03:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id,
      userContext: {
        displayName: 'Test User',
        email: 'test.user@contoso.onmicrosoft.com',
        loginName: 'test.user@contoso.onmicrosoft.com',
        resolvedLanguage: 'en',
        currentWebTitle: 'Project Nova',
        currentWebUrl: 'https://contoso.sharepoint.com/sites/ProjectNova',
        currentSiteTitle: 'Project Nova',
        currentSiteUrl: 'https://contoso.sharepoint.com/sites/ProjectNova'
      },
      proxyConfig: {
        proxyUrl: 'https://proxy.invalid',
        proxyApiKey: 'test-key',
        backend: 'backend',
        deployment: 'deployment',
        apiVersion: '2024-10-21'
      },
      mcpEnvironmentId: 'env-test'
    });

    hybridInteractionEngine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });
    hybridInteractionEngine.emitEvent({
      eventName: 'task.selection.updated',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'selection-mail-thread-current',
      payload: {
        sourceBlockId: 'block-thread',
        sourceBlockType: 'markdown',
        sourceBlockTitle: 'Nova launch blockers',
        selectedCount: 1,
        selectedItems: [{
          index: 1,
          title: 'Nova Launch Blockers',
          kind: 'email',
          itemType: 'email',
          targetContext: {
            mailItemId: 'mail-current',
            source: 'hie-selection'
          }
        }]
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-thread',
      blockType: 'markdown'
    });

    const callbacks = {
      onFunctionCall: jest.fn(async (_callId: string, funcName: string, args: Record<string, unknown>) => {
        if (funcName === 'show_compose_form') {
          expect(parseJsonArg(args.prefill_json)).toMatchObject({
            to: 'alice.smith@contoso.onmicrosoft.com, bob.jones@contoso.onmicrosoft.com',
            cc: 'carol.wilson@contoso.onmicrosoft.com',
            body: 'Project Nova launch recap summary.'
          });
          expect(parseJsonArg(args.static_args_json)).toMatchObject({
            skipSessionHydration: true
          });
          return JSON.stringify({ success: true });
        }
        throw new Error(`Unexpected tool call: ${funcName}`);
      })
    };

    const assistantText = await executeCompoundWorkflowPlan(
      buildPlan('visible_recap_reply_all_mail_discussion', 'current mail discussion'),
      callbacks
    );

    expect(assistantText).toContain('email draft');
    expect(callbacks.onFunctionCall).toHaveBeenCalledTimes(1);
  });
});
