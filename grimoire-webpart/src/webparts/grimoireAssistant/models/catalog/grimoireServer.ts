/**
 * Grimoire Context — Virtual catalog entry for Grimoire's built-in personal context tools.
 * Not a real MCP server — tools are handled directly in handleFunctionCall.
 * Listed in the catalog so `list_m365_servers` shows it alongside M365 servers.
 */

import type { IMcpCatalogEntry } from '../McpServerCatalog';

export const GRIMOIRE_SERVER: IMcpCatalogEntry = {
  id: 'grimoire_context',
  name: 'Grimoire Context',
  description: 'Personal context, recent work, and persistent notes — built into Grimoire (no MCP connection needed)',
  scope: 'local',
  urlTemplate: '', // Not a real MCP server — no URL
  tools: [
    {
      name: 'get_my_profile',
      description: 'Get the current user\'s full profile including job title, department, manager chain, and contact info.',
      inputSchema: {
        type: 'object',
        properties: {},
        required: []
      },
      blockHint: 'user-card'
    },
    {
      name: 'get_recent_documents',
      description: 'Get documents recently opened or edited by the current user.',
      inputSchema: {
        type: 'object',
        properties: {
          max_results: { type: 'number', description: 'Maximum results (default: 10)' }
        }
      }
    },
    {
      name: 'get_trending_documents',
      description: 'Get documents trending around the current user in the organization.',
      inputSchema: {
        type: 'object',
        properties: {
          max_results: { type: 'number', description: 'Maximum results (default: 10)' }
        }
      }
    },
    {
      name: 'save_note',
      description: 'Save a note to persistent memory. Notes survive across sessions.',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'The note text' },
          tags: { type: 'string', description: 'Comma-separated tags for organization' }
        },
        required: ['text']
      }
    },
    {
      name: 'recall_notes',
      description: 'Recall saved notes by keyword or tag. Returns matching notes from persistent memory.',
      inputSchema: {
        type: 'object',
        properties: {
          keyword: { type: 'string', description: 'Keyword to search for in notes' }
        }
      }
    }
  ]
};
