/**
 * User Context handler — Notes & Preferences CRUD.
 *
 * POST /api/user/notes        → save, list, or delete notes
 * POST /api/user/preferences  → get or set preferences
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { enforceRateLimit } from "../middleware/rateLimit.js";
import { requireAuthenticatedUser } from "../middleware/userIdentity.js";
import {
  saveNote,
  listNotes,
  deleteNote,
  getPreferences,
  setPreference,
} from "../storage/UserContextStore.js";
import { createTimingHeaders } from "../utils/diagnostics.js";

// ─── Shared preamble (auth, CORS, body parsing) ─────────────────────

/**
 * Run the shared handler preamble: preflight, CORS, authenticated user resolution,
 * rate-limit, and JSON body parsing.
 * Returns { corsHeaders, body, user } on success, or an early HTTP response on failure.
 */
async function runPreamble(request) {
  const preflight = handlePreflight(request);
  if (preflight) return { response: preflight };

  const corsHeaders = getCorsHeaders(request);

  const auth = requireAuthenticatedUser(request, corsHeaders);
  if (!auth.authenticated) return { response: auth.errorResponse };
  const rateLimitError = enforceRateLimit(auth.user.objectId, corsHeaders);
  if (rateLimitError) return { response: rateLimitError };

  let body;
  try {
    body = await request.json();
  } catch {
    return {
      response: {
        status: 400,
        headers: corsHeaders,
        jsonBody: { error: "Invalid JSON body." },
        },
    };
  }

  return { corsHeaders, body, user: auth.user };
}

/**
 * Handle /api/user/notes
 */
export async function userNotesHandler(request, context) {
  const requestStartedAt = Date.now();
  const diagnostics = { storageInitDurationMs: 0 };
  const preamble = await runPreamble(request);
  if (preamble.response) return preamble.response;
  const { corsHeaders, body, user } = preamble;
  const { action } = body;

  try {
    switch (action) {
      case "save": {
        const { text, tags } = body;
        if (!text || typeof text !== "string") {
          return {
            status: 400,
            headers: corsHeaders,
            jsonBody: { error: "Missing required field: text" },
          };
        }
        const result = await saveNote(
          user.objectId,
          text,
          Array.isArray(tags) ? tags : [],
          user.upnOrEmail,
          diagnostics
        );
        context.log(`[user-notes] Saved note for ${user.objectId}: ${result.id}`);
        return {
          headers: createTimingHeaders(corsHeaders, {
            totalDurationMs: Date.now() - requestStartedAt,
            storageInitDurationMs: diagnostics.storageInitDurationMs,
          }),
          jsonBody: result,
        };
      }

      case "list": {
        const notes = await listNotes(user.objectId, body.keyword, diagnostics);
        context.log(
          `[user-notes] Listed ${notes.length} notes for ${user.objectId}`
        );
        return {
          headers: createTimingHeaders(corsHeaders, {
            totalDurationMs: Date.now() - requestStartedAt,
            storageInitDurationMs: diagnostics.storageInitDurationMs,
          }),
          jsonBody: { notes },
        };
      }

      case "delete": {
        const { noteId } = body;
        if (!noteId) {
          return {
            status: 400,
            headers: corsHeaders,
            jsonBody: { error: "Missing required field: noteId" },
          };
        }
        await deleteNote(user.objectId, noteId, diagnostics);
        context.log(`[user-notes] Deleted note ${noteId} for ${user.objectId}`);
        return {
          headers: createTimingHeaders(corsHeaders, {
            totalDurationMs: Date.now() - requestStartedAt,
            storageInitDurationMs: diagnostics.storageInitDurationMs,
          }),
          jsonBody: { deleted: true },
        };
      }

      default:
        return {
          status: 400,
          headers: corsHeaders,
          jsonBody: {
            error: `Unknown action: ${action}. Expected: save, list, delete`,
          },
        };
    }
  } catch (error) {
    context.error(`[user-notes] Error: ${error.message}`);
    return {
      status: 500,
      headers: createTimingHeaders(corsHeaders, {
        totalDurationMs: Date.now() - requestStartedAt,
        storageInitDurationMs: diagnostics.storageInitDurationMs,
      }),
      jsonBody: { error: error.message },
    };
  }
}

/**
 * Handle /api/user/preferences
 */
export async function userPreferencesHandler(request, context) {
  const requestStartedAt = Date.now();
  const diagnostics = { storageInitDurationMs: 0 };
  const preamble = await runPreamble(request);
  if (preamble.response) return preamble.response;
  const { corsHeaders, body, user } = preamble;
  const { action } = body;

  try {
    switch (action) {
      case "get": {
        const preferences = await getPreferences(user.objectId, diagnostics);
        return {
          headers: createTimingHeaders(corsHeaders, {
            totalDurationMs: Date.now() - requestStartedAt,
            storageInitDurationMs: diagnostics.storageInitDurationMs,
          }),
          jsonBody: { preferences },
        };
      }

      case "set": {
        const { key, value } = body;
        if (!key || typeof key !== "string") {
          return {
            status: 400,
            headers: corsHeaders,
            jsonBody: { error: "Missing required field: key" },
          };
        }
        await setPreference(user.objectId, key, String(value ?? ""), user.upnOrEmail, diagnostics);
        context.log(`[user-prefs] Set ${key} for ${user.objectId}`);
        return {
          headers: createTimingHeaders(corsHeaders, {
            totalDurationMs: Date.now() - requestStartedAt,
            storageInitDurationMs: diagnostics.storageInitDurationMs,
          }),
          jsonBody: { set: true },
        };
      }

      default:
        return {
          status: 400,
          headers: corsHeaders,
          jsonBody: {
            error: `Unknown action: ${action}. Expected: get, set`,
          },
        };
    }
  } catch (error) {
    context.error(`[user-prefs] Error: ${error.message}`);
    return {
      status: 500,
      headers: createTimingHeaders(corsHeaders, {
        totalDurationMs: Date.now() - requestStartedAt,
        storageInitDurationMs: diagnostics.storageInitDurationMs,
      }),
      jsonBody: { error: error.message },
    };
  }
}
