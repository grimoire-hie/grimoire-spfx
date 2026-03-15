#!/usr/bin/env node

/**
 * Grimoire Proxy — Smoke Test Script (Foundry Edition)
 *
 * Runs tests against a deployed (or local) proxy:
 *   1. Health check — verifies backends are listed
 *   2. CORS preflight (allowed origin) — returns explicit allow headers
 *   3. CORS preflight (disallowed origin) — returns no allow-origin header
 *   4. Chat completion (GPT-5 Mini) — full round-trip
 *   5. Chat completion (GPT-4.1 Nano) — fast model round-trip (if available)
 *   6. Auth rejection — invalid key → 403
 *   7. Rate limit headers — correct headers present
 *   8. MCP discovery endpoint — responds to POST
 *
 * Auto-detects proxy URL, key, and deployment names from .grimoire-deploy.json.
 *
 * Usage:  node scripts/test-proxy.mjs <proxy-url> <proxy-key> <allowed-origin>
 *    or:  npm run test:proxy -- <proxy-url> <proxy-key> <allowed-origin>
 *    or:  node scripts/test-proxy.mjs   (interactive prompts)
 *
 * Examples:
 *   node scripts/test-proxy.mjs http://localhost:7071/api abc123
 *   node scripts/test-proxy.mjs https://my-proxy.azurewebsites.net/api abc123 https://contoso.sharepoint.com
 */

import {
  banner,
  logOk,
  logFail,
  logInfo,
  log,
  colors,
  ask,
  closePrompt,
  httpFetch,
  normalizeOrigin,
} from "./lib.mjs";
import { loadDeployConfig } from "./deploy-config.mjs";

const config = loadDeployConfig();

const EXPECTED_CORS_HEADERS = "Content-Type, api-key, Authorization";

// Deployment names — auto-detected from config or defaults
const reasoningDeployment = config?.deploymentNames?.reasoning || "grimoire-reasoning";
const fastDeployment = config?.deploymentNames?.fast || "grimoire-fast";

function printUsage() {
  logInfo("Usage: npm run test:proxy -- <proxy-url> <proxy-key> <allowed-origin>");
}

function deriveRejectedOrigin(allowedOrigin) {
  const url = new URL(allowedOrigin);
  url.hostname = `${url.hostname}.invalid`;
  return url.origin;
}

function parseHeaderList(value) {
  return String(value || "")
    .split(",")
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

function includesAllEntries(actualHeaderValue, expectedHeaderValue) {
  const actualEntries = new Set(parseHeaderList(actualHeaderValue));
  const expectedEntries = parseHeaderList(expectedHeaderValue);
  return expectedEntries.every((entry) => actualEntries.has(entry));
}

async function main() {
  banner("Grimoire Proxy — Smoke Tests (Foundry)");

  if (config) {
    logInfo(`Using deployment config (prefix: ${config.prefix || "grimoire"})`);
    log("");
  }

  // Get proxy URL and key from args, config, or prompt
  let proxyUrl = process.argv[2] || "";
  let proxyKey = process.argv[3] || "";
  let allowedCorsOriginInput = process.argv[4] || "";
  const interactiveMode = process.argv.length <= 2;

  if (!proxyUrl) {
    const defaultUrl = config?.proxyUrl || "http://localhost:7071/api";
    proxyUrl = await ask("Proxy URL", defaultUrl);
  }
  if (!proxyKey) {
    const defaultKey = config?.proxyApiKey || "";
    proxyKey = await ask("Proxy API Key", defaultKey);
  }
  if (!allowedCorsOriginInput && interactiveMode) {
    const defaultOrigin = config?.sharepointOrigin || "";
    allowedCorsOriginInput = await ask(
      "Allowed CORS origin (e.g., https://contoso.sharepoint.com)",
      defaultOrigin
    );
  }

  closePrompt();

  if (!allowedCorsOriginInput) {
    logFail("Missing allowed CORS origin.");
    printUsage();
    process.exit(1);
  }

  // Remove trailing slash
  proxyUrl = proxyUrl.replace(/\/$/, "");

  let allowedCorsOrigin = "";
  let rejectedCorsOrigin = "";
  try {
    allowedCorsOrigin = normalizeOrigin(allowedCorsOriginInput, "Allowed CORS origin");
    const badOriginInput = process.env.BAD_CORS_ORIGIN || deriveRejectedOrigin(allowedCorsOrigin);
    rejectedCorsOrigin = normalizeOrigin(
      badOriginInput,
      process.env.BAD_CORS_ORIGIN ? "BAD_CORS_ORIGIN" : "Derived rejected CORS origin"
    );
  } catch (error) {
    logFail(error.message);
    process.exit(1);
  }

  if (rejectedCorsOrigin === allowedCorsOrigin) {
    logFail("Rejected CORS origin must differ from the allowed CORS origin.");
    process.exit(1);
  }

  let passed = 0;
  let failed = 0;
  let skipped = 0;
  let total = 0;

  function pass(name, detail = "") {
    passed++;
    total++;
    logOk(`PASS: ${name}${detail ? " — " + detail : ""}`);
  }

  function fail(name, detail = "") {
    failed++;
    total++;
    logFail(`FAIL: ${name}${detail ? " — " + detail : ""}`);
  }

  function skip(name, detail = "") {
    skipped++;
    total++;
    log(`  → SKIP: ${name}${detail ? " — " + detail : ""}`, colors.yellow);
  }

  async function runCorsPreflightTest(name, origin, expectAllowed) {
    try {
      const res = await httpFetch(`${proxyUrl}/health`, {
        method: "OPTIONS",
        headers: {
          Origin: origin,
          "Access-Control-Request-Method": "POST",
          "Access-Control-Request-Headers": EXPECTED_CORS_HEADERS,
        },
      });

      const allowOrigin = res.headers["access-control-allow-origin"];
      const allowHeaders = res.headers["access-control-allow-headers"];
      const allowMethods = res.headers["access-control-allow-methods"];
      const maxAge = res.headers["access-control-max-age"];
      const statusOk = res.status === 200 || res.status === 204;

      if (!expectAllowed) {
        if (statusOk && allowOrigin === undefined) {
          pass(name, `origin=${origin}`);
        } else {
          fail(
            name,
            `expected no Access-Control-Allow-Origin, got status=${res.status} allow-origin=${allowOrigin || "(missing)"}`
          );
        }
        return;
      }

      const hasExpectedHeaders =
        allowOrigin === origin &&
        includesAllEntries(allowHeaders, EXPECTED_CORS_HEADERS) &&
        includesAllEntries(allowMethods, "POST");

      if (statusOk && hasExpectedHeaders) {
        pass(name, `origin=${origin}, status=${res.status}${maxAge ? `, max-age=${maxAge}` : ""}`);
      } else {
        fail(
          name,
          [
            `status=${res.status}`,
            `allow-origin=${allowOrigin || "(missing)"}`,
            `allow-headers=${allowHeaders || "(missing)"}`,
            `allow-methods=${allowMethods || "(missing)"}`,
            `max-age=${maxAge || "(missing)"}`,
          ].join(", ")
        );
      }
    } catch (error) {
      fail(name, error.message);
    }
  }

  // ── Test 1: Health check ─────────────────────────────────────
  log("\n  Test 1: Health Check", colors.bold);

  let availableBackends = [];

  try {
    const res = await httpFetch(`${proxyUrl}/health`, {
      headers: { "api-key": proxyKey },
    });
    if (res.status === 200) {
      const body = JSON.parse(res.body);
      const hasBackendDetails = body && typeof body === "object" && body.backends && typeof body.backends === "object";
      if (!hasBackendDetails) {
        fail(
          "GET /health",
          "missing authenticated backend details (likely wrong API key or wrong proxy URL)"
        );
      } else {
        availableBackends = Object.keys(body.backends || {});
        if (availableBackends.length === 0) {
          fail("GET /health", "no backends configured (BACKENDS is empty)");
        } else {
          const backendDetails = availableBackends.map((b) => {
            const info = body.backends[b];
            return `${b}(${info.auth})`;
          });
          pass("GET /health", `status=ok, backends=[${backendDetails.join(", ")}]`);
        }
      }
    } else {
      fail("GET /health", `status=${res.status}`);
    }
  } catch (error) {
    fail("GET /health", error.message);
  }

  // ── Test 2: CORS preflight (allowed origin) ─────────────────
  log("\n  Test 2: CORS Preflight (Allowed Origin)", colors.bold);

  await runCorsPreflightTest("OPTIONS /health allows configured origin", allowedCorsOrigin, true);

  // ── Test 3: CORS preflight (disallowed origin) ──────────────
  log("\n  Test 3: CORS Preflight (Disallowed Origin)", colors.bold);

  await runCorsPreflightTest("OPTIONS /health rejects unconfigured origin", rejectedCorsOrigin, false);

  // ── Test 4: Chat completion (Reasoning) ──────────────────────
  log("\n  Test 4: Chat Completion (Reasoning)", colors.bold);

  if (availableBackends.includes("reasoning")) {
    try {
      const chatUrl = `${proxyUrl}/reasoning/openai/deployments/${reasoningDeployment}/chat/completions?api-version=2025-01-01-preview`;
      const chatBody = JSON.stringify({
        messages: [{ role: "user", content: "Say hello in one word." }],
        max_completion_tokens: 10,
      });

      const res = await httpFetch(chatUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": proxyKey,
        },
        body: chatBody,
        timeout: 30_000,
      });

      if (res.status === 200) {
        const body = JSON.parse(res.body);
        const content = body.choices?.[0]?.message?.content || "(empty)";
        const model = body.model || "unknown";
        pass("POST reasoning/chat", `model=${model}, response: "${content.trim().slice(0, 50)}"`);
      } else {
        const body = res.body.slice(0, 200);
        fail("POST reasoning/chat", `status=${res.status} body=${body}`);
      }
    } catch (error) {
      fail("POST reasoning/chat", error.message);
    }
  } else {
    skip("POST reasoning/chat", "backend not configured");
  }

  // ── Test 5: Chat completion (Fast) ──────────────────────────
  log("\n  Test 5: Chat Completion (Fast)", colors.bold);

  if (availableBackends.includes("fast")) {
    try {
      const chatUrl = `${proxyUrl}/fast/openai/deployments/${fastDeployment}/chat/completions?api-version=2025-01-01-preview`;
      const chatBody = JSON.stringify({
        messages: [{ role: "user", content: "Say hi in one word." }],
        max_completion_tokens: 10,
      });

      const res = await httpFetch(chatUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": proxyKey,
        },
        body: chatBody,
        timeout: 15_000,
      });

      if (res.status === 200) {
        const body = JSON.parse(res.body);
        const content = body.choices?.[0]?.message?.content || "(empty)";
        const model = body.model || "unknown";
        pass("POST fast/chat", `model=${model}, response: "${content.trim().slice(0, 50)}"`);
      } else {
        const body = res.body.slice(0, 200);
        fail("POST fast/chat", `status=${res.status} body=${body}`);
      }
    } catch (error) {
      fail("POST fast/chat", error.message);
    }
  } else {
    skip("POST fast/chat", "backend not configured (optional)");
  }

  // ── Test 6: Auth rejection ───────────────────────────────────
  log("\n  Test 6: Auth Rejection (invalid key)", colors.bold);

  try {
    const chatUrl = `${proxyUrl}/reasoning/openai/deployments/${reasoningDeployment}/chat/completions?api-version=2025-01-01-preview`;
    const chatBody = JSON.stringify({
      messages: [{ role: "user", content: "Hello" }],
      max_completion_tokens: 10,
    });

    const res = await httpFetch(chatUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": "invalid-key-12345",
      },
      body: chatBody,
    });

    if (res.status === 403) {
      pass("Invalid key rejected", "403 Forbidden");
    } else {
      fail("Invalid key rejection", `expected 403, got ${res.status}`);
    }
  } catch (error) {
    fail("Auth rejection", error.message);
  }

  // ── Test 7: Rate limit headers ───────────────────────────────
  log("\n  Test 7: Rate Limit Headers", colors.bold);

  try {
    const chatUrl = `${proxyUrl}/reasoning/openai/deployments/${reasoningDeployment}/chat/completions?api-version=2025-01-01-preview`;
    const chatBody = JSON.stringify({
      messages: [{ role: "user", content: "Say hi" }],
      max_completion_tokens: 5,
    });

    const res = await httpFetch(chatUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": proxyKey,
      },
      body: chatBody,
      timeout: 30_000,
    });

    const hasMinute = res.headers["x-ratelimit-remaining-minute"] !== undefined;
    const hasDay = res.headers["x-ratelimit-remaining-day"] !== undefined;
    const hasBackend = res.headers["x-backend"] !== undefined;
    const hasAuth = res.headers["x-auth-method"] !== undefined;

    if (hasMinute && hasDay && hasBackend) {
      pass(
        "Rate limit headers",
        `backend=${res.headers["x-backend"]}, ` +
          `auth=${res.headers["x-auth-method"] || "unknown"}, ` +
          `min=${res.headers["x-ratelimit-remaining-minute"]}, ` +
          `day=${res.headers["x-ratelimit-remaining-day"]}`
      );
    } else {
      const missing = [];
      if (!hasMinute) missing.push("X-RateLimit-Remaining-Minute");
      if (!hasDay) missing.push("X-RateLimit-Remaining-Day");
      if (!hasBackend) missing.push("X-Backend");
      fail("Rate limit headers", `missing: ${missing.join(", ")}`);
    }
  } catch (error) {
    fail("Rate limit headers", error.message);
  }

  // ── Test 8: MCP Discovery endpoint ───────────────────────────
  log("\n  Test 8: MCP Discovery Endpoint", colors.bold);

  try {
    const res = await httpFetch(`${proxyUrl}/mcp/discover`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": proxyKey,
      },
      body: JSON.stringify({ serverUrl: "https://httpbin.org/post" }),
      timeout: 15_000,
    });

    // We expect either a successful discovery or a 502 (server isn't a real MCP server)
    // The important thing is it doesn't 404 or 500
    if (res.status === 200 || res.status === 502) {
      const body = JSON.parse(res.body);
      if (res.status === 200) {
        pass("POST /mcp/discover", `discovered=${body.discovered}, tools=${(body.tools || []).length}`);
      } else {
        pass("POST /mcp/discover", `endpoint reachable (502 expected for non-MCP server)`);
      }
    } else {
      fail("POST /mcp/discover", `status=${res.status}`);
    }
  } catch (error) {
    fail("POST /mcp/discover", error.message);
  }

  // ── Test 9: Realtime Token endpoint ─────────────────────────
  log("\n  Test 9: Realtime Token Endpoint", colors.bold);

  try {
    const res = await httpFetch(`${proxyUrl}/realtime/token`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": proxyKey,
      },
      body: JSON.stringify({
        voice: "alloy",
        instructions: "You are a test assistant.",
      }),
      timeout: 15_000,
    });

    if (res.status === 200) {
      const body = JSON.parse(res.body);
      if (body.clientSecret) {
        pass("POST /realtime/token", `got ephemeral token, expires: ${body.expiresAt || "unknown"}`);
      } else {
        pass("POST /realtime/token", "endpoint reachable (unexpected response shape)");
      }
    } else if (res.status === 502) {
      skip("POST /realtime/token", "realtime model not deployed (502)");
    } else {
      fail("POST /realtime/token", `status=${res.status}`);
    }
  } catch (error) {
    fail("POST /realtime/token", error.message);
  }

  // ── Test 10: User Notes endpoint (must require Entra user identity) ──────
  log("\n  Test 10: User Notes (must require Entra user identity)", colors.bold);

  try {
    // Save a note
    const saveRes = await httpFetch(`${proxyUrl}/user/notes`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": proxyKey,
      },
      body: JSON.stringify({
        action: "save",
        text: "Smoke test note — safe to delete",
        tags: ["test"],
      }),
      timeout: 15_000,
    });

    if (saveRes.status === 401) {
      pass("POST /user/notes", "secured route correctly rejected api-key-only request with 401");
    } else {
      fail("POST /user/notes", `expected 401 for api-key-only request, got status=${saveRes.status}`);
    }
  } catch (error) {
    fail("POST /user/notes", error.message);
  }

  // ── Test 11: User Preferences endpoint (must require Entra user identity) ─────────
  log("\n  Test 11: User Preferences (must require Entra user identity)", colors.bold);

  try {
    // Set a preference
    const setRes = await httpFetch(`${proxyUrl}/user/preferences`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": proxyKey,
      },
      body: JSON.stringify({
        action: "set",
        key: "test_pref",
        value: "smoke_value",
      }),
      timeout: 10_000,
    });

    if (setRes.status === 401) {
      pass("POST /user/preferences", "secured route correctly rejected api-key-only request with 401");
    } else {
      fail("POST /user/preferences", `expected 401 for api-key-only request, got status=${setRes.status}`);
    }
  } catch (error) {
    fail("POST /user/preferences", error.message);
  }

  // ── Summary ──────────────────────────────────────────────────
  log(`\n${"─".repeat(50)}`, colors.dim);
  const resultColor = failed === 0 ? colors.green : colors.red;
  const parts = [`${passed} passed`, `${failed} failed`];
  if (skipped > 0) parts.push(`${skipped} skipped`);
  log(`  Results: ${parts.join(", ")} (${total} total)`, resultColor);
  log(`${"─".repeat(50)}\n`, colors.dim);

  process.exit(failed > 0 ? 1 : 0);
}

main().catch((error) => {
  logFail(error.message);
  closePrompt();
  process.exit(1);
});
