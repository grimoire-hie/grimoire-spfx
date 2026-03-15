/**
 * User Profile MCP Server catalog entry.
 * Generated from discovery-output.json — 5 tools.
 * Source of truth: MCP protocol discovery, not guesses.
 */

import type { IMcpCatalogEntry } from '../McpServerCatalog';
import { GATEWAY_BASE } from './constants';

export const ME_SERVER: IMcpCatalogEntry = {
  id: 'mcp_MeServer',
  name: 'User Profile',
  description: 'User profiles, org hierarchy, manager chain',
  scope: 'McpServers.Me.All',
  urlTemplate: `${GATEWAY_BASE}/mcp_MeServer`,
  tools: [
    { name: 'GetMyDetails', description: 'Retrieve profile details for the currently signed-in user ("me, my"). Use this when you need the signed-in user\'s identity or profile details (e.g., display name, email, job title), including for e...', inputSchema: {"type":"object","properties":{"select":{"type":"string","description":"Always pass in comma-separated list of properties you need"},"expand":{"type":"string","description":"Expand related entities"}}}, blockHint: 'user-card' as const },
    { name: 'GetUserDetails', description: 'Find a specified user\'s profile by name, email, or ID. Use this when you need to look up a specific person in your organization.', inputSchema: {"type":"object","properties":{"userIdentifier":{"type":"string","description":"The user's name or object ID (GUID) or userPrincipalName (email-like UPN)."},"select":{"type":"string","description":"Always pass in comma-separated list of properties you need"},"expand":{"type":"string","description":"Expand a related entity for the user"}},"required":["userIdentifier"]}, blockHint: 'user-card' as const },
    { name: 'GetMultipleUsersDetails', description: 'Search for multiple users in the directory by name, job title, office location, or other properties.', inputSchema: {"type":"object","properties":{"searchValues":{"type":"array","description":"List of search terms (e.g., ['John Smith', 'Jane Doe'] or ['Software Engineer', 'Product Manager'] or ['Building 40', 'Building 41']). Each term is...","items":{"type":"string"}},"propertyToSearchBy":{"type":"string","description":"User property to search (e.g., 'displayName', 'jobTitle', 'officeLocation', 'userPrincipalName', 'id')."},"select":{"type":"string","description":"Comma-separated list of user properties to include in response (e.g., 'displayName,mail,jobTitle,officeLocation,mobilePhone')"},"expand":{"type":"string","description":"Navigation properties to expand (e.g., 'manager' to include manager details)"},"top":{"type":"object"},"orderby":{"type":"string","description":"Property name to sort results by (e.g., 'displayName', 'jobTitle')"}},"required":["searchValues"]}, blockHint: 'selection-list' as const },
    { name: 'GetManagerDetails', description: 'Get a user\'s manager information - name, email, job title, etc.,', inputSchema: {"type":"object","properties":{"userId":{"type":"string","description":"Name of the user whose manager to retrieve. Use \"me\" for current / signed-in user."},"select":{"type":"string","description":"Always pass in comma-separated list of properties you need"}},"required":["userId"]}, blockHint: 'user-card' as const },
    { name: 'GetDirectReportsDetails', description: 'Retrieve a user\'s team, or direct reports (people who report to them in the organizational hierarchy). Use this for organizational team structure, NOT for Microsoft Teams workspace membership. Exam...', inputSchema: {"type":"object","properties":{"userId":{"type":"string","description":"Name of the user whose direct reports (organizational team members) to retrieve. Use \"me\" for current / signed-in user."},"select":{"type":"string","description":"Always pass in comma-separated list of properties you need for each direct report"}},"required":["userId"]}, blockHint: 'selection-list' as const }
  ]
};
