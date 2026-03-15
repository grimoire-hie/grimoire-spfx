import type { IFormData, ISelectionItem } from '../../models/IBlock';
import {
  buildShareSubmissionTrace,
  buildTeamsMessageBody,
  isInternalSharePreset,
  resolveNamedSelectionItem,
  resolveShareSubmissionTargetContext
} from './ShareSubmissionService';

const TEAMS: ISelectionItem[] = [
  { id: 'team-1', label: 'Marketing' },
  { id: 'team-2', label: 'Marketing Operations' },
  { id: 'team-3', label: 'Finance' }
];

describe('ShareSubmissionService helpers', () => {
  it('matches exact names before partial names', () => {
    expect(resolveNamedSelectionItem(TEAMS, 'Finance', 'team')).toEqual({
      item: TEAMS[2]
    });
  });

  it('accepts a unique partial name when there is no exact match', () => {
    expect(resolveNamedSelectionItem(TEAMS, 'Operations', 'team')).toEqual({
      item: TEAMS[1]
    });
  });

  it('returns a readable ambiguity error when multiple matches remain', () => {
    const result = resolveNamedSelectionItem(TEAMS, 'Market', 'team');

    expect(result.item).toBeUndefined();
    expect(result.error).toBe('More than one team matched "Market": Marketing, Marketing Operations.');
  });

  it('returns a readable missing error when there is no match', () => {
    const result = resolveNamedSelectionItem(TEAMS, 'Legal', 'team');

    expect(result.item).toBeUndefined();
    expect(result.error).toBe('No team matched "Legal". Available teams: Marketing, Marketing Operations, Finance.');
  });

  it('recognizes only the internal share presets', () => {
    expect(isInternalSharePreset('share-teams-chat')).toBe(true);
    expect(isInternalSharePreset('share-teams-channel')).toBe(true);
    expect(isInternalSharePreset('email-compose')).toBe(false);
  });

  it('formats Teams post bodies as HTML with preserved line breaks and links', () => {
    const payload = buildTeamsMessageBody('Hi team,\n\nReview this:\nhttps://contoso.sharepoint.com/sites/dev/SPFx.pdf');

    expect(payload.contentType).toBe('html');
    expect(payload.content).toContain('Hi team,');
    expect(payload.content).toContain('<br/><br/>');
    expect(payload.content).toContain('Review this:');
    expect(payload.content).toContain('<a href="https://contoso.sharepoint.com/sites/dev/SPFx.pdf"');
    expect(payload.content).toContain('>https://contoso.sharepoint.com/sites/dev/SPFx.pdf</a>');
  });

  it('prefers explicit form target context for internal share traces', () => {
    const formData = {
      preset: 'share-teams-channel',
      description: 'Share to Teams',
      submissionTarget: {
        toolName: 'internal-share',
        serverId: 'internal-share',
        staticArgs: {},
        targetContext: {
          siteUrl: 'https://tenant.sharepoint.com/sites/copilot-test',
          fileOrFolderUrl: 'https://tenant.sharepoint.com/sites/copilot-test/Shared%20Documents/Power%20Platform.pdf',
          fileOrFolderName: 'Power Platform.pdf',
          source: 'explicit-user'
        }
      }
    } as unknown as IFormData;

    const sourceContext = {
      targetContext: {
        siteUrl: 'https://tenant.sharepoint.com/sites/current-site',
        siteName: 'current-site',
        source: 'current-page' as const
      }
    };

    expect(resolveShareSubmissionTargetContext(formData, sourceContext)).toEqual(expect.objectContaining({
      siteUrl: 'https://tenant.sharepoint.com/sites/copilot-test',
      fileOrFolderName: 'Power Platform.pdf',
      source: 'explicit-user'
    }));
  });

  it('builds a structured share trace for internal channel shares', () => {
    const formData = {
      preset: 'share-teams-channel',
      description: 'Share to Teams',
      submissionTarget: {
        toolName: 'internal-share',
        serverId: 'internal-share',
        staticArgs: {},
        targetContext: {
          siteUrl: 'https://tenant.sharepoint.com/sites/copilot-test',
          fileOrFolderUrl: 'https://tenant.sharepoint.com/sites/copilot-test/Shared%20Documents/Power%20Platform.pdf',
          fileOrFolderName: 'Power Platform.pdf',
          source: 'explicit-user'
        }
      }
    } as unknown as IFormData;

    const trace = buildShareSubmissionTrace(
      formData,
      {
        content: 'Shared from Grimoire.',
        teamName: 'Marketing',
        channelName: 'General'
      },
      {},
      undefined,
      { success: true, message: 'Shared to Marketing / General.' }
    );

    expect(trace).toEqual(expect.objectContaining({
      serverId: 'internal-share',
      serverName: 'Internal share flow',
      toolName: 'share-teams-channel',
      targetSummary: 'Power Platform.pdf',
      targetSource: 'explicit-user',
      finalSummary: 'Shared to Marketing / General.'
    }));
    expect(trace.resolvedArgs).toEqual(expect.objectContaining({
      teamName: 'Marketing',
      channelName: 'General'
    }));
  });
});
