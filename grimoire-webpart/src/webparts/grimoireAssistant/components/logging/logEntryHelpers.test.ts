import type { ILogEntry } from '../../services/logging/LogTypes';
import { findLatestMcpExecutionTrace, getLogEntryToggleLabel, hasLogEntryDetails, parseMcpExecutionTrace } from './logEntryHelpers';

function createLogEntry(detail?: string): ILogEntry {
  return {
    id: 'log-1',
    timestamp: new Date('2026-03-07T10:00:00.000Z'),
    level: 'debug',
    category: 'system',
    message: 'Test message',
    detail
  };
}

describe('logEntryHelpers', () => {
  it('treats missing or blank detail as non-expandable', () => {
    expect(hasLogEntryDetails(createLogEntry())).toBe(false);
    expect(hasLogEntryDetails(createLogEntry('   '))).toBe(false);
  });

  it('treats populated detail as expandable', () => {
    expect(hasLogEntryDetails(createLogEntry('Payload detail'))).toBe(true);
  });

  it('returns user-facing detail toggle labels', () => {
    expect(getLogEntryToggleLabel(false)).toBe('Show details');
    expect(getLogEntryToggleLabel(true)).toBe('Hide details');
  });

  it('parses MCP execution traces from structured log entries', () => {
    const entry: ILogEntry = {
      ...createLogEntry(JSON.stringify({
        serverName: 'SharePoint Lists',
        toolName: 'listLists',
        rawArgs: { siteName: 'copilot-test-cooking' },
        normalizedArgs: { siteName: 'copilot-test-cooking' },
        resolvedArgs: { siteId: 'tenant,site,web' },
        requiredFields: ['siteId'],
        unwrapPath: ['response', 'payload'],
        targetSummary: 'copilot-test',
        targetSource: 'hie-selection',
        finalBlockTitle: 'Lists',
        inferredBlockType: 'list-items'
      })),
      category: 'mcp',
      message: 'MCP execution trace'
    };

    expect(parseMcpExecutionTrace(entry)).toEqual({
      trace: expect.objectContaining({
        serverName: 'SharePoint Lists',
        toolName: 'listLists'
      }),
      toolLabel: 'SharePoint Lists / listLists',
      serverLabel: 'SharePoint Lists',
      targetLabel: 'copilot-test',
      targetSourceLabel: 'HIE selection',
      resultLabel: 'Lists · list-items',
      requiredLabel: 'siteId',
      rawArgsLabel: '{"siteName":"copilot-test-cooking"}',
      normalizedArgsLabel: '{"siteName":"copilot-test-cooking"}',
      resolvedArgsLabel: '{"siteId":"tenant,site,web"}',
      unwrapLabel: 'response -> payload',
      recoveryLabel: undefined
    });
  });

  it('ignores non-trace log entries', () => {
    expect(parseMcpExecutionTrace(createLogEntry('Payload detail'))).toBeUndefined();
  });

  it('parses internal share traces from structured log entries', () => {
    const entry: ILogEntry = {
      ...createLogEntry(JSON.stringify({
        serverId: 'internal-share',
        serverName: 'Internal share flow',
        toolName: 'share-teams-channel',
        rawArgs: { fieldValues: { content: 'Shared from Grimoire.' } },
        resolvedArgs: {
          teamId: 'team-1',
          teamName: 'Marketing',
          channelId: 'channel-1',
          channelName: 'General'
        },
        targetSummary: 'Power Platform.pdf @ copilot-test',
        targetSource: 'explicit-user',
        finalSummary: 'Shared to Marketing / General.'
      })),
      category: 'mcp',
      message: 'Share execution trace'
    };

    expect(parseMcpExecutionTrace(entry)).toEqual(expect.objectContaining({
      toolLabel: 'Internal share flow / share-teams-channel',
      serverLabel: 'Internal share flow',
      targetLabel: 'Power Platform.pdf @ copilot-test',
      targetSourceLabel: 'Explicit user target',
      resultLabel: 'Shared to Marketing / General.',
      rawArgsLabel: '{"fieldValues":{"content":"Shared from Grimoire."}}',
      resolvedArgsLabel: '{"teamId":"team-1","teamName":"Marketing","channelId":"channel-1","channelName":"General"}'
    }));
  });

  it('finds the latest MCP trace in a mixed log stream', () => {
    const latest = findLatestMcpExecutionTrace([
      createLogEntry('Payload detail'),
      {
        ...createLogEntry(JSON.stringify({
          serverName: 'Profile',
          toolName: 'GetUserDetails',
          rawArgs: { userEmail: 'someone@example.com' },
          recoverySteps: ['normalized args', 'resolved selected person'],
          targetSource: 'hie-selection'
        })),
        category: 'mcp',
        message: 'MCP execution trace'
      }
    ]);

    expect(latest).toEqual(expect.objectContaining({
      toolLabel: 'Profile / GetUserDetails',
      recoveryLabel: 'normalized args -> resolved selected person',
      targetSourceLabel: 'HIE selection'
    }));
  });
});
