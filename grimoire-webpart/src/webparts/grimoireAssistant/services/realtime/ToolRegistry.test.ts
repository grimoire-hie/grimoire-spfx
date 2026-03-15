import { getTools } from './ToolRegistry';

describe('ToolRegistry MCP schema alignment', () => {
  it('keeps MCP tool required parameters aligned with runtime behavior', () => {
    const tools = getTools();
    const byName = new Map(tools.map((t) => [t.name, t]));

    expect(byName.get('connect_mcp_server')?.parameters.required).toEqual(['server_url']);
    expect(byName.get('call_mcp_tool')?.parameters.required).toEqual(['tool_name']);
    expect(byName.get('list_mcp_tools')?.parameters.required).toEqual([]);
  });

  it('registers content reader tools with expected schema', () => {
    const tools = getTools();
    const byName = new Map(tools.map((t) => [t.name, t]));

    expect(tools).toHaveLength(35);
    expect(byName.has('research_public_web')).toBe(true);
    expect(byName.has('read_file_content')).toBe(true);
    expect(byName.has('read_email_content')).toBe(true);
    expect(byName.has('read_teams_messages')).toBe(true);
    expect(byName.get('read_email_content')?.parameters.required).toEqual([]);
    expect(byName.get('read_file_content')?.parameters.properties.mode?.enum).toEqual(['summarize', 'full', 'answer']);
    expect(byName.get('read_file_content')?.parameters.properties.file_urls?.type).toBe('array');
    expect(byName.get('read_file_content')?.parameters.required || []).not.toContain('file_url');
    expect(byName.get('read_email_content')?.parameters.properties.mode?.enum).toEqual(['summarize', 'full', 'answer']);
    expect(byName.get('read_teams_messages')?.parameters.properties.mode?.enum).toEqual(['summarize', 'full', 'answer']);
  });

  it('omits avatar-only tools when the avatar is disabled', () => {
    const tools = getTools({ avatarEnabled: false });
    const byName = new Map(tools.map((t) => [t.name, t]));

    expect(tools).toHaveLength(34);
    expect(byName.has('set_expression')).toBe(false);
  });
});
