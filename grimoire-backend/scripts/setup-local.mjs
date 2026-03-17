#!/usr/bin/env node

/**
 * Grimoire Proxy — Local Development Setup (Foundry Edition)
 *
 * Sets up local.settings.json for local development and testing.
 * Uses API key auth (not managed identity) since local dev
 * doesn't have a system-assigned identity.
 *
 * Configures two backends:
 *   - reasoning: GPT-5 Mini for cinema mode / complex tasks
 *   - fast: GPT-5 Nano (or GPT-4.1 Nano) for fast hints (optional)
 *
 * Both backends use the same Foundry endpoint with different deployments.
 *
 * Auto-detects values from .grimoire-deploy.json if available.
 *
 * Usage:  node scripts/setup-local.mjs
 *    or:  npm run setup
 */

import { existsSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  banner,
  logStep,
  logOk,
  logSkip,
  logFail,
  logInfo,
  log,
  colors,
  execLive,
  ask,
  closePrompt,
  requireCommand,
} from "./lib.mjs";
import { loadDeployConfig } from "./deploy-config.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_DIR = resolve(__dirname, "..");
const SETTINGS_PATH = resolve(PROJECT_DIR, "local.settings.json");

async function main() {
  banner("Grimoire Proxy — Local Setup (Foundry)");

  const config = loadDeployConfig();
  const prefix = config?.prefix || "grimoire";

  if (config) {
    logInfo(`Detected deployment config (prefix: ${prefix})`);
    log("");
  }

  // ── Step 1: Prerequisites ────────────────────────────────────
  logStep(1, "Checking prerequisites");

  requireCommand("node", "https://nodejs.org");
  requireCommand("func", "https://learn.microsoft.com/azure/azure-functions/functions-run-local");

  // ── Step 2: Create local.settings.json ───────────────────────
  logStep(2, "Local Settings");

  if (existsSync(SETTINGS_PATH)) {
    logSkip("local.settings.json");
    logInfo("To reset, delete local.settings.json and re-run this script.");

    const overwrite = await ask("Overwrite existing settings? (yes/no)", "no");
    if (overwrite.toLowerCase() !== "yes" && overwrite.toLowerCase() !== "y") {
      logInfo("Keeping existing settings.");

      // Still install deps and show instructions
      logStep(3, "Dependencies");
      logInfo("Installing npm dependencies...");
      execLive(`npm install --prefix "${PROJECT_DIR}"`);
      logOk("Dependencies installed");

      banner("Ready!");
      logInfo("Start the local proxy:  npm start");
      logInfo(`Then test:  npm run test:proxy -- http://localhost:7071/api YOUR_KEY`);
      closePrompt();
      return;
    }
  }

  log("\n  Configure your Microsoft Foundry connection for local dev.\n", colors.dim);
  logInfo("You need a Foundry resource (kind: AIServices) with deployed models.");
  logInfo("Get these values from Azure Portal → your Foundry resource → Keys and Endpoint.\n");

  const defaultEndpoint = config?.foundryName
    ? `https://${config.foundryName}.cognitiveservices.azure.com/`
    : "";
  const endpoint = await ask(
    "Foundry endpoint (e.g., https://your-foundry.cognitiveservices.azure.com/)",
    defaultEndpoint
  );

  if (!endpoint) {
    logFail("Endpoint is required.");
    closePrompt();
    process.exit(1);
  }

  const apiKey = await ask("Foundry API key", "");

  if (!apiKey) {
    logFail("API key is required for local development.");
    closePrompt();
    process.exit(1);
  }

  // Ask about fast hints backend
  log("\n  The proxy supports multiple backends:\n", colors.dim);
  logInfo("reasoning  — GPT-5 Mini for cinema mode / AI Builder (required)");
  logInfo("fast       — GPT-5 Nano / GPT-4.1 Nano for fast hints (optional)");
  log("");

  const enableFast = await ask(
    "Enable fast backend? (yes/no)",
    "yes"
  );
  const hasFast = enableFast.toLowerCase() === "yes" || enableFast.toLowerCase() === "y";

  const backends = hasFast ? "reasoning,fast" : "reasoning";
  const realtimeDeploymentName = config?.deploymentNames?.realtime || `${prefix}-realtime`;

  const settingsValues = {
    AzureWebJobsStorage: "UseDevelopmentStorage=true",
    FUNCTIONS_WORKER_RUNTIME: "node",
    ALLOW_PERMISSIVE_LOCAL_CORS: "true",
    REQUESTS_PER_MINUTE: "60",
    REQUESTS_PER_DAY: "5000",
    MCP_TOKEN_FORWARD_ALLOWLIST: "agent365.svc.cloud.microsoft",
    ALLOWED_ORIGINS: "",
    BACKENDS: backends,
    BACKEND_reasoning_ENDPOINT: endpoint,
    BACKEND_reasoning_KEY: apiKey,
    REALTIME_DEPLOYMENT_NAME: realtimeDeploymentName,
  };

  if (hasFast) {
    // Same Foundry endpoint, same API key, different deployment name
    settingsValues.BACKEND_fast_ENDPOINT = endpoint;
    settingsValues.BACKEND_fast_KEY = apiKey;
  }

  const settings = {
    IsEncrypted: false,
    Values: settingsValues,
  };

  writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2) + "\n");
  logOk("Created local.settings.json");
  logOk("Auth: Azure function key auth (bypassed locally by func start)");

  closePrompt();

  // ── Step 3: Install dependencies ─────────────────────────────
  logStep(3, "Dependencies");

  logInfo("Installing npm dependencies...");
  execLive(`npm install --prefix "${PROJECT_DIR}"`);
  logOk("Dependencies installed");

  // ── Summary ──────────────────────────────────────────────────
  banner("Local Setup Complete!");

  const reasoningDeployment = config?.deploymentNames?.reasoning || `${prefix}-reasoning`;
  const fastDeployment = config?.deploymentNames?.fast || `${prefix}-fast`;

  logInfo(`Endpoint:     ${endpoint}`);
  logInfo(`Backends:     ${backends}`);
  logInfo(`Auth:         Azure function key (bypassed locally)`);
  logInfo(`Prefix:       ${prefix}`);

  log("\n  Start the local proxy:\n", colors.bold);
  logInfo("npm start");
  logInfo("→ http://localhost:7071/api/health");

  log("\n  API routes:\n", colors.bold);
  logInfo(`Chat (Reasoning):   http://localhost:7071/api/reasoning/openai/deployments/${reasoningDeployment}/chat/completions`);
  if (hasFast) {
    logInfo(`Chat (Fast):        http://localhost:7071/api/fast/openai/deployments/${fastDeployment}/chat/completions`);
  }

  log("\n  Test it:\n", colors.bold);
  logInfo(`npm run test:proxy -- http://localhost:7071/api`);
}

main().catch((error) => {
  logFail(error.message);
  closePrompt();
  process.exit(1);
});
