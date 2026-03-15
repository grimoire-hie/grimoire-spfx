import assert from "node:assert/strict";
import { webcrypto } from "node:crypto";
import test from "node:test";

if (!globalThis.crypto) {
  Object.defineProperty(globalThis, "crypto", {
    configurable: true,
    value: webcrypto,
  });
}

const {
  __resetTableClientCachesForTests,
  __setTableClientFactoryForTests,
  getPreferences,
} = await import("../src/storage/UserContextStore.js");

test("getPreferences shares a single table initialization across concurrent cold-start requests", async () => {
  process.env.AzureWebJobsStorage = "UseDevelopmentStorage=true";
  __resetTableClientCachesForTests();

  let createTableCalls = 0;

  class FakeTableClient {
    async createTable() {
      createTableCalls += 1;
      await new Promise((resolve) => setTimeout(resolve, 20));
    }

    async *listEntities() {
      return;
    }
  }

  __setTableClientFactoryForTests(() => new FakeTableClient());

  const diagnosticsA = { storageInitDurationMs: 0 };
  const diagnosticsB = { storageInitDurationMs: 0 };

  await Promise.all([
    getPreferences("abc123", diagnosticsA),
    getPreferences("abc123", diagnosticsB),
  ]);

  assert.equal(createTableCalls, 1);
  assert.equal(
    [diagnosticsA.storageInitDurationMs > 0, diagnosticsB.storageInitDurationMs > 0].filter(Boolean).length,
    1
  );

  __resetTableClientCachesForTests();
});
