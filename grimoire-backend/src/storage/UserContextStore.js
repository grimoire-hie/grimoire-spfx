/**
 * UserContextStore — Azure Table Storage wrapper for user notes and preferences.
 * Uses @azure/data-tables with the AzureWebJobsStorage connection string.
 * Tables are auto-created on first use (idempotent).
 *
 * Tables:
 *   GrimoireNotes       — PartitionKey: userObjectId, RowKey: noteId (UUID)
 *   GrimoirePreferences — PartitionKey: userObjectId, RowKey: key (preference name)
 */

import { TableClient } from "@azure/data-tables";
import { randomUUID } from "crypto";

const NOTES_TABLE = "GrimoireNotes";
const PREFERENCES_TABLE = "GrimoirePreferences";
const MAX_NOTES_PER_USER = 200;

export function normalizeUserPartitionKey(userObjectId) {
  const normalized = String(userObjectId || "").trim().toLowerCase();
  if (!normalized) {
    throw new Error("Authenticated user object ID is required.");
  }
  return normalized;
}

/**
 * Sanitize and build an OData PartitionKey filter.
 * Escapes single quotes to prevent OData injection.
 * @param {string} partitionKey
 * @returns {string}
 */
function oDataPartitionFilter(partitionKey) {
  const safe = partitionKey.replace(/'/g, "''");
  return `PartitionKey eq '${safe}'`;
}

/** @type {Map<string, TableClient>} */
const tableClients = new Map();
/** @type {Map<string, Promise<TableClient>>} */
const tableClientInitializers = new Map();
let tableClientFactory = (connString, tableName) => TableClient.fromConnectionString(connString, tableName);

function recordStorageInitDuration(diagnostics, durationMs) {
  if (!diagnostics || !(durationMs > 0)) {
    return;
  }

  diagnostics.storageInitDurationMs = (diagnostics.storageInitDurationMs || 0) + durationMs;
}

/**
 * Get or create a TableClient for the given table name.
 * Creates the table on first access (idempotent).
 * @param {string} tableName
 * @returns {Promise<TableClient>}
 */
async function getTableClient(tableName, diagnostics) {
  if (tableClients.has(tableName)) {
    return tableClients.get(tableName);
  }

  if (tableClientInitializers.has(tableName)) {
    return tableClientInitializers.get(tableName);
  }

  const connString = process.env.AzureWebJobsStorage;
  if (!connString) {
    throw new Error("AzureWebJobsStorage connection string not configured");
  }

  const initStartedAt = Date.now();
  const initPromise = (async () => {
    const client = tableClientFactory(connString, tableName);

    // Create table if it doesn't exist (idempotent)
    try {
      await client.createTable();
    } catch (err) {
      // 409 = table already exists — that's fine
      if (err.statusCode !== 409) {
        throw err;
      }
    }

    tableClients.set(tableName, client);
    return client;
  })();

  tableClientInitializers.set(tableName, initPromise);

  try {
    const client = await initPromise;
    recordStorageInitDuration(diagnostics, Date.now() - initStartedAt);
    return client;
  } finally {
    tableClientInitializers.delete(tableName);
  }
}

// ─── Notes ──────────────────────────────────────────────────────

/**
 * Save a note for a user.
 * @param {string} userObjectId
 * @param {string} text — Note text (max 2KB)
 * @param {string[]} tags — Tags for organization
 * @param {string} [ownerUpnOrEmail]
 * @returns {Promise<{id: string}>}
 */
export async function saveNote(userObjectId, text, tags = [], ownerUpnOrEmail = "", diagnostics) {
  const client = await getTableClient(NOTES_TABLE, diagnostics);
  const partitionKey = normalizeUserPartitionKey(userObjectId);

  // Enforce max notes per user
  const existingCount = await countEntities(client, partitionKey);
  if (existingCount >= MAX_NOTES_PER_USER) {
    throw new Error(`Maximum ${MAX_NOTES_PER_USER} notes per user. Delete old notes first.`);
  }

  // Enforce max note size
  const trimmedText = text.length > 2048 ? text.substring(0, 2048) : text;

  const noteId = randomUUID();
  await client.createEntity({
    partitionKey,
    rowKey: noteId,
    text: trimmedText,
    tags: tags.join(","),
    ownerObjectId: partitionKey,
    ownerUpnOrEmail: String(ownerUpnOrEmail || "").trim(),
    createdAt: new Date().toISOString(),
  });

  return { id: noteId };
}

/**
 * List notes for a user, optionally filtered by keyword.
 * @param {string} userObjectId
 * @param {string} [keyword]
 * @returns {Promise<Array<{id: string, text: string, tags: string[], createdAt: string}>>}
 */
export async function listNotes(userObjectId, keyword, diagnostics) {
  const client = await getTableClient(NOTES_TABLE, diagnostics);
  const partitionKey = normalizeUserPartitionKey(userObjectId);

  const notes = [];
  const query = client.listEntities({
    queryOptions: {
      filter: oDataPartitionFilter(partitionKey),
    },
  });

  for await (const entity of query) {
    const note = {
      id: entity.rowKey,
      text: entity.text || "",
      tags: entity.tags ? String(entity.tags).split(",").filter(Boolean) : [],
      createdAt: entity.createdAt || "",
    };

    // Filter by keyword if provided
    if (keyword) {
      const kw = keyword.toLowerCase();
      const matchesText = note.text.toLowerCase().includes(kw);
      const matchesTags = note.tags.some((t) => t.toLowerCase().includes(kw));
      if (!matchesText && !matchesTags) continue;
    }

    notes.push(note);
  }

  // Sort by createdAt descending (newest first)
  notes.sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || ""));

  return notes;
}

/**
 * Delete a note by ID.
 * @param {string} userObjectId
 * @param {string} noteId
 */
export async function deleteNote(userObjectId, noteId, diagnostics) {
  const client = await getTableClient(NOTES_TABLE, diagnostics);
  await client.deleteEntity(normalizeUserPartitionKey(userObjectId), noteId);
}

// ─── Preferences ────────────────────────────────────────────────

/**
 * Get all preferences for a user.
 * @param {string} userObjectId
 * @returns {Promise<Record<string, string>>}
 */
export async function getPreferences(userObjectId, diagnostics) {
  const client = await getTableClient(PREFERENCES_TABLE, diagnostics);
  const partitionKey = normalizeUserPartitionKey(userObjectId);

  const prefs = {};
  const query = client.listEntities({
    queryOptions: {
      filter: oDataPartitionFilter(partitionKey),
    },
  });

  for await (const entity of query) {
    prefs[entity.rowKey] = entity.value || "";
  }

  return prefs;
}

/**
 * Set a single preference.
 * @param {string} userObjectId
 * @param {string} key
 * @param {string} value
 * @param {string} [ownerUpnOrEmail]
 */
export async function setPreference(userObjectId, key, value, ownerUpnOrEmail = "", diagnostics) {
  const client = await getTableClient(PREFERENCES_TABLE, diagnostics);
  const pk = normalizeUserPartitionKey(userObjectId);
  await client.upsertEntity({
    partitionKey: pk,
    rowKey: key,
    value,
    ownerObjectId: pk,
    ownerUpnOrEmail: String(ownerUpnOrEmail || "").trim(),
    updatedAt: new Date().toISOString(),
  });
}

// ─── Helpers ────────────────────────────────────────────────────

/**
 * Count entities in a partition (for enforcing limits).
 * @param {TableClient} client
 * @param {string} partitionKey
 * @returns {Promise<number>}
 */
async function countEntities(client, partitionKey, limit = MAX_NOTES_PER_USER) {
  let count = 0;
  const query = client.listEntities({
    queryOptions: {
      filter: oDataPartitionFilter(partitionKey),
      select: ["PartitionKey"],
    },
  });
  for await (const _entity of query) {
    count++;
    if (count >= limit) break;
  }
  return count;
}

export function __setTableClientFactoryForTests(factory) {
  tableClientFactory = factory;
}

export function __resetTableClientCachesForTests() {
  tableClients.clear();
  tableClientInitializers.clear();
  tableClientFactory = (connString, tableName) => TableClient.fromConnectionString(connString, tableName);
}
