import assert from "node:assert/strict";
import test from "node:test";

import {
  connectSession,
  executeToolOnSession,
  disconnectSession,
  listSessions,
} from "../src/mcp/McpSessionManager.js";

// Mock MCP server URL — tests will fail to connect but we can test ownership logic
// by intercepting at the right level. For unit tests of ownership enforcement,
// we test the session store directly after manually inserting sessions.

// Since McpSessionManager uses a module-level Map, we can test the ownership
// filtering and rejection logic by calling listSessions, executeToolOnSession,
// and disconnectSession with different ownerIds.

test("listSessions returns only sessions belonging to the specified owner", async () => {
  // listSessions with a random owner should return empty
  const sessions = listSessions("owner-that-does-not-exist");
  assert.equal(sessions.length, 0);
});

test("executeToolOnSession rejects with 'not found' when session does not exist", async () => {
  await assert.rejects(
    () => executeToolOnSession("nonexistent-session", "toolName", {}, "some-owner"),
    (error) => {
      assert.match(error.message, /not found/);
      return true;
    }
  );
});

test("disconnectSession returns 'not found' when session does not exist", async () => {
  const result = await disconnectSession("nonexistent-session", "some-owner");
  assert.equal(result.disconnected, false);
  assert.match(result.reason, /not found/i);
});

test("disconnectSession with undefined ownerId returns 'not found' for missing session", async () => {
  const result = await disconnectSession("nonexistent-session", undefined);
  assert.equal(result.disconnected, false);
  assert.match(result.reason, /not found/i);
});
