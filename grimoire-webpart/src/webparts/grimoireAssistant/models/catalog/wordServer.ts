/**
 * Word Documents MCP Server catalog entry.
 * Generated from discovery-output.json — 4 tools.
 * Source of truth: MCP protocol discovery, not guesses.
 */

import type { IMcpCatalogEntry } from '../McpServerCatalog';
import { GATEWAY_BASE } from './constants';

export const WORD_SERVER: IMcpCatalogEntry = {
  id: 'mcp_WordServer',
  name: 'Word Documents',
  description: 'Create and edit Word documents, add comments',
  scope: 'McpServers.Word.All',
  urlTemplate: `${GATEWAY_BASE}/mcp_WordServer`,
  tools: [
    { name: 'CreateDocument', description: 'Create a new Word document in the root of the user\'s OneDrive. Provide a desired file name (may be empty) and HTML or plain text content to populate the document body. Make sure all the tags and ur...', inputSchema: {"type":"object","properties":{"fileName":{"type":"string","description":"The desired name for the new Word document. If empty, a default name will be generated."},"contentInHtml":{"type":"string","description":"[Required] The HTML or plain text content to populate the document body. Make sure all the tags and urls in the content are valid and resolvable."},"shareWith":{"type":"string","description":"[Optional] Email address to share the document with. If you have the email address of the user who requested the document creation, automatically s..."}},"required":["fileName"]}, blockHint: 'info-card' as const },
    { name: 'GetDocumentContent', description: 'Fetch the raw Word document (DOCX) content from OneDrive/SharePoint given its SharePoint sharing URL. Returns a JSON payload with filename, size, driveId, documentId, plain text content and all com...', inputSchema: {"type":"object","properties":{"url":{"type":"string","description":"[Required] The OneDrive/SharePoint URL of the Word document."}},"required":["url"]}, blockHint: 'file-preview' as const },
    { name: 'AddComment', description: 'Add a new comment in the Word document. Provide driveId, documentId, and the comment text.', inputSchema: {"type":"object","properties":{"driveId":{"type":"string","description":"[Required] The drive ID of the document."},"documentId":{"type":"string","description":"[Required] The document ID."},"newComment":{"type":"string","description":"[Required] The comment text to add."}},"required":["driveId","documentId","newComment"]}, blockHint: 'info-card' as const },
    { name: 'ReplyToComment', description: 'Reply to a comment in the Word document. Provide commentId, driveId, documentId, and the reply text.', inputSchema: {"type":"object","properties":{"commentId":{"type":"string","description":"[Required] The ID of the comment to reply to."},"driveId":{"type":"string","description":"[Required] The drive ID of the document."},"documentId":{"type":"string","description":"[Required] The document ID."},"newComment":{"type":"string","description":"[Required] The reply text to add."}},"required":["commentId","driveId","documentId","newComment"]}, blockHint: 'info-card' as const }
  ]
};
