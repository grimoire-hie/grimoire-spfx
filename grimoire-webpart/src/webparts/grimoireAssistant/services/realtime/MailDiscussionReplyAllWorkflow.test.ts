import { createBlock } from '../../models/IBlock';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import {
  buildMailDiscussionReplyAllPlan,
  consumePendingMailDiscussionReplyAllPlan,
  extractMailDiscussionSearchQuery,
  MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
  resolveMailDiscussionReplyAllPlanFromToolCall,
  setPendingMailDiscussionReplyAllPlan
} from './MailDiscussionReplyAllWorkflow';

afterEach(() => {
  hybridInteractionEngine.reset();
  useGrimoireStore.setState({
    blocks: [],
    transcript: [],
    activeActionBlockId: undefined,
    selectedActionIndices: [],
    proxyConfig: undefined
  });
  setPendingMailDiscussionReplyAllPlan(undefined);
});

describe('MailDiscussionReplyAllWorkflow', () => {
  it('extracts the thread subject from recap-share prompts', () => {
    expect(extractMailDiscussionSearchQuery(
      'send the recap to all the person involved in the launch blockers email'
    )).toBe('launch blockers');
    expect(extractMailDiscussionSearchQuery(
      'reply all to the nova launch blockers mail'
    )).toBe('nova launch blockers');
  });

  it('builds the shared mail-recipient workflow plan from a recap-share prompt', () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:10:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id
    });

    const plan = buildMailDiscussionReplyAllPlan(
      'send the recap to all the person involved in the launch blockers email'
    );

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'launch blockers',
        selectionHint: 'none'
      }
    });
    expect(plan?.steps.map((step) => step.kind)).toEqual([
      'resolve-mail-thread',
      'compose-email'
    ]);
  });

  it('builds the shared mail-recipient workflow plan from a voice email-search tool call when a summary card is visible', () => {
    const summaryBlock = createBlock('info-card', 'Summary of Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary of Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    const plan = resolveMailDiscussionReplyAllPlanFromToolCall('search_emails', {
      query: 'Nova Launch Blockers'
    });

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'Nova Launch Blockers',
        selectionHint: 'none'
      }
    });
  });

  it('upgrades generic email-compose tool calls into the mail-recipient workflow when the model asks for thread participants', () => {
    const summaryBlock = createBlock('info-card', 'Summary of Nova_Launch_Recap.docx', {
      kind: 'info-card',
      heading: 'Summary of Nova_Launch_Recap.docx',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    const plan = resolveMailDiscussionReplyAllPlanFromToolCall('show_compose_form', {
      preset: 'email-compose',
      title: 'Send Nova Launch Recap',
      description: 'Send the Nova Launch Recap summary to the participants from the Launch Blockers document.',
      prefill_json: {
        subject: 'Nova Launch Recap Summary',
        body: 'Summary body without recipients.'
      }
    });

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'Launch Blockers',
        selectionHint: 'none'
      }
    });
  });

  it('intercepts compose-form calls when description uses "behind the X" topic pattern without email/mail keywords', () => {
    const summaryBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    const plan = resolveMailDiscussionReplyAllPlanFromToolCall('show_compose_form', {
      preset: 'email-compose',
      title: 'Send Nova Launch Recap',
      description: 'Send the recap summary to all involved, including Alice Smith and the team behind the launch blockers.',
      prefill_json: {
        to: 'alice.smith@contoso.onmicrosoft.com',
        subject: 'Nova Launch Recap Summary',
        body: 'Recap body text.'
      }
    });

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'launch blockers',
        selectionHint: 'none'
      }
    });
  });

  it('extracts topic from "involved in" even when description starts with "Compose an email"', () => {
    const summaryBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    const plan = resolveMailDiscussionReplyAllPlanFromToolCall('show_compose_form', {
      preset: 'email-compose',
      title: 'Send Nova Launch Recap',
      description: 'Compose an email with the Nova Launch Recap attached or summarized for everyone involved in the launch blockers.',
      prefill_json: {
        to: 'alice.smith@contoso.onmicrosoft.com',
        subject: 'Nova Launch Recap Summary',
        body: 'Hi all,\n\nPlease find the Nova Launch Recap.\n\nBest,\nTest'
      }
    });

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'launch blockers',
        selectionHint: 'none'
      }
    });
  });

  it('upgrades email-read tool calls into the mail-recipient workflow when a recap is visible', () => {
    const summaryBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    const plan = resolveMailDiscussionReplyAllPlanFromToolCall('read_email_content', {
      subject: 'Nova_Launch_Blockers',
      mode: 'summarize'
    });

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'Nova Launch Blockers',
        selectionHint: 'none'
      }
    });
  });

  it('still overrides author-resolution compose calls even if the model already inserted one recipient', () => {
    const summaryBlock = createBlock('info-card', 'Summary of Nova_Launch_Recap.docx', {
      kind: 'info-card',
      heading: 'Summary of Nova_Launch_Recap.docx',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    expect(resolveMailDiscussionReplyAllPlanFromToolCall('show_compose_form', {
      preset: 'email-compose',
      title: 'Send Nova Launch Recap to Launch Blockers Authors',
      description: 'Compose an email to the authors of the "Nova_Launch_Blockers" document and include the recap.',
      prefill_json: {
        to: 'alice.smith@contoso.onmicrosoft.com',
        subject: 'Nova Launch Recap Summary',
        body: 'Hi all,\n\nPlease find attached the Nova Launch Recap.\n\nBest,\nTest'
      }
    })).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'Nova Launch Blockers',
        selectionHint: 'none'
      }
    });
  });

  it('does not override ordinary explicit-recipient compose calls without recipient-source language', () => {
    const summaryBlock = createBlock('info-card', 'Summary of Nova_Launch_Recap.docx', {
      kind: 'info-card',
      heading: 'Summary of Nova_Launch_Recap.docx',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [summaryBlock],
      transcript: [],
      activeActionBlockId: summaryBlock.id
    });

    expect(resolveMailDiscussionReplyAllPlanFromToolCall('show_compose_form', {
      preset: 'email-compose',
      title: 'Email Alice Smith',
      description: 'Send the recap to Anna.',
      prefill_json: {
        to: 'alice.smith@contoso.onmicrosoft.com',
        subject: 'Nova Launch Recap Summary'
      }
    })).toBeUndefined();
  });

  it('matches summary artifact titles where "Summary" or "Recap" appears at the end', () => {
    const recapBlock = createBlock('info-card', 'Nova_Launch_Recap Summary', {
      kind: 'info-card',
      heading: 'Nova_Launch_Recap Summary',
      body: 'Nova launch recap summary body text.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:10:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id
    });

    const plan = buildMailDiscussionReplyAllPlan(
      'send the recap to all the person involved in the launch blockers email'
    );

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'launch blockers',
        selectionHint: 'none'
      }
    });
  });

  it('matches summary artifact titles with "Recap" embedded in the document name', () => {
    const recapBlock = createBlock('info-card', 'Nova_Launch_Recap Summary', {
      kind: 'info-card',
      heading: 'Nova_Launch_Recap Summary',
      body: 'Nova launch recap summary body text.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [],
      activeActionBlockId: recapBlock.id
    });

    const plan = resolveMailDiscussionReplyAllPlanFromToolCall('search_emails', {
      query: 'Nova Launch Blockers'
    });

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID
    });
  });

  it('does not infer the mail-recipient workflow from a raw email search when no visible summary card is selected', () => {
    const personBlock = createBlock('info-card', 'Alice Smith', {
      kind: 'info-card',
      heading: 'Alice Smith',
      body: 'Head of Marketing'
    });
    useGrimoireStore.setState({
      blocks: [personBlock],
      transcript: [],
      activeActionBlockId: personBlock.id
    });

    expect(resolveMailDiscussionReplyAllPlanFromToolCall('search_emails', {
      query: 'Nova Launch Blockers'
    })).toBeUndefined();
  });

  it('still plans mail-recipient requests against the current thread when HIE already carries mail context', () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:11:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id
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

    const plan = buildMailDiscussionReplyAllPlan('reply all to this mail discussion with the recap');

    expect(plan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'current mail discussion'
      }
    });
  });

  it('consumes pending chooser workflows only after a mail thread selection is available', () => {
    const recapBlock = createBlock('info-card', 'Summary: Nova_Launch_Recap', {
      kind: 'info-card',
      heading: 'Summary: Nova_Launch_Recap',
      body: 'Nova launch recap summary.'
    });
    useGrimoireStore.setState({
      blocks: [recapBlock],
      transcript: [
        { role: 'assistant', text: 'The recap is in the panel.', timestamp: new Date('2026-03-11T20:12:00.000Z') }
      ],
      activeActionBlockId: recapBlock.id
    });

    const plan = buildMailDiscussionReplyAllPlan(
      'send the recap to all the person involved in the launch blockers email'
    );
    expect(plan).toBeDefined();
    setPendingMailDiscussionReplyAllPlan(plan);
    expect(consumePendingMailDiscussionReplyAllPlan()).toBeUndefined();

    hybridInteractionEngine.initialize(() => { /* no-op */ }, { sendContextMessage: jest.fn() });
    hybridInteractionEngine.emitEvent({
      eventName: 'task.selection.updated',
      source: 'action-panel',
      surface: 'action-panel',
      correlationId: 'selection-mail-thread-ready',
      payload: {
        sourceBlockId: 'block-thread-picker',
        sourceBlockType: 'selection-list',
        sourceBlockTitle: 'Choose the email to use for recipients.',
        selectedCount: 1,
        selectedItems: [{
          index: 1,
          title: 'Nova Launch Blockers',
          kind: 'email',
          itemType: 'email',
          targetContext: {
            mailItemId: 'mail-selected',
            source: 'hie-selection'
          }
        }]
      },
      exposurePolicy: { mode: 'silent-context', relevance: 'contextual' },
      blockId: 'block-thread-picker',
      blockType: 'selection-list'
    });

    const resumedPlan = consumePendingMailDiscussionReplyAllPlan();
    expect(resumedPlan).toMatchObject({
      familyId: MAIL_DISCUSSION_REPLY_ALL_FAMILY_ID,
      slots: {
        query: 'launch blockers'
      }
    });
  });
});
