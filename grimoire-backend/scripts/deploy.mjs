#!/usr/bin/env node

/**
 * Grimoire Proxy — Full Azure Deployment Script (Foundry Edition)
 *
 * Creates all Azure resources from scratch:
 *   1. Resource Group
 *   2. Microsoft Foundry resource (kind: AIServices) — multi-model access
 *   3. Model deployment: reasoning (cinema mode, complex reasoning)
 *   4. Model deployment: fast (fast tasks, classification, hints)
 *   5. Storage Account
 *   6. Function App with managed identity
 *   7. Role assignment (Function App → Foundry)
 *   8. Deploys the proxy code
 *   9. Configures environment variables
 *  10. Verifies with health check
 *
 * Usage:  node scripts/deploy.mjs
 *    or:  npm run deploy
 *
 * Idempotent — safe to re-run. Skips resources that already exist.
 * Cross-platform — runs on Windows, macOS, and Linux.
 */

import { randomBytes, randomUUID } from "node:crypto";
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
  exec,
  execLive,
  ask,
  closePrompt,
  requireCommand,
  httpFetch,
  normalizeOrigin,
  capitalize,
} from "./lib.mjs";
import { saveDeployConfig, loadDeployConfig, CONFIG_FILENAME } from "./deploy-config.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_DIR = resolve(__dirname, "..");

// =============================================================================
// CONSTANTS
// =============================================================================

const DEFAULT_LOCATION = "swedencentral";
const BACKEND_API_SCOPE_VALUE = "user_impersonation";

// Model catalog — available choices per slot
const MODEL_CATALOG = {
  primary: {
    "gpt-5-mini": { name: "GPT-5 Mini", version: "2025-08-07", fallback: { modelName: "gpt-4.1-mini", modelVersion: "2025-04-14" } },
    "gpt-4.1-mini": { name: "GPT-4.1 Mini", version: "2025-04-14", fallback: null },
  },
  fast: {
    "gpt-5-nano": { name: "GPT-5 Nano", version: "2025-08-07", fallback: { modelName: "gpt-4.1-nano", modelVersion: "2025-04-14" } },
    "gpt-4.1-nano": { name: "GPT-4.1 Nano", version: "2025-04-14", fallback: null },
  },
  realtime: {
    "gpt-realtime-1.5": { name: "GPT Realtime 1.5", version: "2026-02-23", fallback: { modelName: "gpt-realtime", modelVersion: "2025-08-28" } },
    "gpt-realtime": { name: "GPT Realtime (GA)", version: "2025-08-28", fallback: null },
  },
};

// Build MODELS config from admin choices (populated during Step 3)
function buildModels({ prefix, primaryModel, fastModel, realtimeModel, realtimeCapacity = 6 }) {
  const primary = MODEL_CATALOG.primary[primaryModel];
  const models = {
    reasoning: {
      name: primary.name,
      purpose: "Cinema mode, AI Builder, complex reasoning",
      deploymentName: `${prefix}-reasoning`,
      modelName: primaryModel,
      modelVersion: primary.version,
      modelFormat: "OpenAI",
      skuName: "GlobalStandard",
      skuCapacity: 100,
      fallback: primary.fallback,
    },
  };

  if (fastModel) {
    const fh = MODEL_CATALOG.fast[fastModel];
    models.fast = {
      name: fh.name,
      purpose: "Fast tasks — hints, validation, classification",
      deploymentName: `${prefix}-fast`,
      modelName: fastModel,
      modelVersion: fh.version,
      modelFormat: "OpenAI",
      skuName: "GlobalStandard",
      skuCapacity: 100,
      fallback: fh.fallback,
    };
  }

  if (realtimeModel) {
    const rt = MODEL_CATALOG.realtime[realtimeModel];
    models.realtime = {
      name: rt.name,
      purpose: "Real-time voice conversation for agent design",
      deploymentName: `${prefix}-realtime`,
      modelName: realtimeModel,
      modelVersion: rt.version,
      modelFormat: "OpenAI",
      skuName: "GlobalStandard",
      skuCapacity: realtimeCapacity,
      fallback: rt.fallback,
    };
  }

  return models;
}

// =============================================================================
// HELPERS
// =============================================================================

function shellQuote(value) {
  return `'${String(value).replace(/'/g, "'\\''")}'`;
}

function parseJsonOutput(value, fallback = null) {
  const trimmed = String(value || "").trim();
  if (!trimmed || trimmed === "null") {
    return fallback;
  }
  return JSON.parse(trimmed);
}

function buildUserScope(existingScopeId, projectName) {
  const displayName = capitalize(projectName);
  return {
    adminConsentDescription: `Allow ${displayName} to access user-owned notes and preferences.`,
    adminConsentDisplayName: `Access ${displayName} user data`,
    id: existingScopeId || randomUUID(),
    isEnabled: true,
    type: "User",
    userConsentDescription: `Allow ${displayName} to access your notes and preferences.`,
    userConsentDisplayName: `Access your ${displayName} data`,
    value: BACKEND_API_SCOPE_VALUE,
  };
}

function ensureBackendApiApplication(backendApiAppDisplayName, projectName) {
  let app = parseJsonOutput(
    exec(
      `az ad app list --display-name "${backendApiAppDisplayName}" --query "[0]" --output json`,
      { silent: true, ignoreError: true }
    )
  );

  if (!app) {
    app = JSON.parse(
      exec(
        `az ad app create ` +
          `--display-name "${backendApiAppDisplayName}" ` +
          `--sign-in-audience AzureADMyOrg ` +
          `--requested-access-token-version 2 ` +
          `--output json`,
        { silent: true }
      )
    );
    logOk(`Created Entra app registration: ${backendApiAppDisplayName}`);
  } else {
    logSkip(`${backendApiAppDisplayName} app registration`);
  }

  const graphApp = JSON.parse(
    exec(
      `az rest --method GET ` +
        `--uri "https://graph.microsoft.com/v1.0/applications/${app.id}?$select=id,appId,displayName,identifierUris,api" ` +
        `--output json`,
      { silent: true }
    )
  );

  const backendApiResource = `api://${graphApp.appId}`;
  const existingScope = Array.isArray(graphApp.api?.oauth2PermissionScopes)
    ? graphApp.api.oauth2PermissionScopes.find((scope) => scope.value === BACKEND_API_SCOPE_VALUE)
    : undefined;

  const patchBody = {
    identifierUris: [backendApiResource],
    api: {
      requestedAccessTokenVersion: 2,
      oauth2PermissionScopes: [buildUserScope(existingScope?.id, projectName)],
    },
  };

  exec(
    `az rest --method PATCH ` +
      `--uri "https://graph.microsoft.com/v1.0/applications/${graphApp.id}" ` +
      `--headers Content-Type=application/json ` +
      `--body ${shellQuote(JSON.stringify(patchBody))} ` +
      `--output none`,
    { silent: true }
  );
  logOk(`Configured API scope on ${backendApiAppDisplayName}`);

  try {
    exec(`az ad sp show --id ${graphApp.appId} --output none`, { silent: true });
    logSkip(`${backendApiAppDisplayName} service principal`);
  } catch {
    exec(`az ad sp create --id ${graphApp.appId} --output none`, { silent: true });
    logOk(`Created service principal for ${backendApiAppDisplayName}`);
  }

  return {
    appId: graphApp.appId,
    objectId: graphApp.id,
    resource: backendApiResource,
    displayName: backendApiAppDisplayName,
  };
}

// SPFx first-party app ID — constant across all tenants.
const SPFX_CLIENT_APP_ID = "08e18876-6177-487e-b8b5-cf950c1e598c";

/**
 * Ensure the SPFx service principal has an OAuth2 permission grant to the backend API.
 * This is the programmatic equivalent of approving the API permission request in
 * SharePoint Admin Center → API access.
 */
function ensureSpfxPermissionGrant(backendApiAppId) {
  // Resolve service principal IDs
  const spfxSpJson = parseJsonOutput(
    exec(
      `az ad sp list --filter "appId eq '${SPFX_CLIENT_APP_ID}'" --query "[0].id" --output json`,
      { silent: true, ignoreError: true }
    )
  );
  if (!spfxSpJson) {
    logInfo("SPFx service principal not found in tenant — permission grant skipped (will need manual approval).");
    return;
  }
  const spfxSpId = spfxSpJson;

  const backendSpId = parseJsonOutput(
    exec(
      `az ad sp list --filter "appId eq '${backendApiAppId}'" --query "[0].id" --output json`,
      { silent: true, ignoreError: true }
    )
  );
  if (!backendSpId) {
    logInfo("Backend API service principal not found — permission grant skipped.");
    return;
  }

  // Check if a grant already exists
  const existingGrants = JSON.parse(
    exec(
      `az rest --method GET ` +
        `--uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?\\$filter=clientId eq '${spfxSpId}' and resourceId eq '${backendSpId}'" ` +
        `--query "value" --output json`,
      { silent: true, ignoreError: true }
    ) || "[]"
  );

  const hasGrant = existingGrants.some((g) =>
    String(g.scope || "").split(" ").includes(BACKEND_API_SCOPE_VALUE)
  );

  if (hasGrant) {
    logSkip("SPFx → Backend API permission grant");
    return;
  }

  // Create the grant
  const grantBody = JSON.stringify({
    clientId: spfxSpId,
    consentType: "AllPrincipals",
    resourceId: backendSpId,
    scope: BACKEND_API_SCOPE_VALUE,
  });

  exec(
    `az rest --method POST ` +
      `--uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" ` +
      `--headers Content-Type=application/json ` +
      `--body ${shellQuote(grantBody)} ` +
      `--output none`,
    { silent: true }
  );
  logOk("Granted SPFx → Backend API permission (user_impersonation)");
}

/**
 * Configure Easy Auth in "pass-through" mode.
 *
 * SECURITY NOTE: requireAuthentication is false because:
 *   - LLM proxy, MCP, and other routes use Azure function key auth (authLevel: "function")
 *   - Only /api/user/* routes require the Entra bearer token (Easy Auth)
 *   - The health endpoint uses function key auth
 *
 * When Easy Auth is enabled but not required, Azure still decodes and injects
 * the x-ms-client-principal header for authenticated callers. This allows
 * MCP handlers to resolve per-user identity for session ownership and rate
 * limiting without breaking function-key-only callers.
 */
function configureFunctionAppEasyAuth({
  subscriptionId,
  tenantId,
  resourceGroup,
  functionAppName,
  backendApiAppId,
  backendApiResource,
}) {
  const authSettingsBody = {
    properties: {
      platform: {
        enabled: true,
      },
      globalValidation: {
        requireAuthentication: false,
        unauthenticatedClientAction: "AllowAnonymous",
      },
      httpSettings: {
        requireHttps: true,
      },
      identityProviders: {
        azureActiveDirectory: {
          enabled: true,
          registration: {
            clientId: backendApiAppId,
            openIdIssuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
          },
          validation: {
            allowedAudiences: [backendApiResource, backendApiAppId],
          },
        },
      },
      login: {
        tokenStore: {
          enabled: false,
        },
      },
    },
  };

  exec(
    `az rest --method PUT ` +
      `--uri "https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.Web/sites/${functionAppName}/config/authsettingsV2?api-version=2023-12-01" ` +
      `--headers Content-Type=application/json ` +
      `--body ${shellQuote(JSON.stringify(authSettingsBody))} ` +
      `--output none`,
    { silent: true }
  );
}

/**
 * Deploy a model to the Foundry resource. Returns the actual model name deployed.
 * Returns null if the user chooses to skip an optional model.
 *
 * When the preferred model is not available in the region, the user is prompted
 * to choose: use a fallback model, skip, or abort and re-deploy in another region.
 */
async function deployModel(foundryName, resourceGroup, config, location) {
  const { deploymentName, modelName, modelVersion, modelFormat, skuName, skuCapacity } = config;

  // Check if already deployed — if so, ensure capacity and model match
  try {
    const existing = JSON.parse(
      exec(
        `az cognitiveservices account deployment show ` +
          `--name ${foundryName} ` +
          `--resource-group ${resourceGroup} ` +
          `--deployment-name ${deploymentName} ` +
          `--output json`,
        { silent: true }
      )
    );
    const currentCapacity = existing.sku?.capacity;
    const existingModel = existing.properties?.model?.name;
    const needsModelUpgrade = existingModel && existingModel !== modelName;

    if (needsModelUpgrade) {
      logInfo(`${deploymentName}: upgrading ${existingModel} → ${modelName}...`);
      // Fall through to primary deploy path below
    } else if (currentCapacity !== skuCapacity) {
      logInfo(`${deploymentName}: updating capacity ${currentCapacity} → ${skuCapacity}...`);
      exec(
        `az cognitiveservices account deployment create ` +
          `--name ${foundryName} ` +
          `--resource-group ${resourceGroup} ` +
          `--deployment-name ${deploymentName} ` +
          `--model-name ${existingModel || modelName} ` +
          `--model-version "${existing.properties?.model?.version || modelVersion}" ` +
          `--model-format "${existing.properties?.model?.format || modelFormat}" ` +
          `--sku-capacity ${skuCapacity} ` +
          `--sku-name ${existing.sku?.name || skuName} ` +
          `--output none`,
        { timeout: 180_000 }
      );
      logOk(`${deploymentName}: capacity updated to ${skuCapacity}`);
      return existingModel || modelName;
    } else {
      logSkip(`${deploymentName} (${config.name})`);
      return existingModel || modelName;
    }
  } catch {
    // Doesn't exist — deploy it
  }

  // Try primary model
  logInfo(`Deploying ${modelName} as '${deploymentName}'...`);
  try {
    exec(
      `az cognitiveservices account deployment create ` +
        `--name ${foundryName} ` +
        `--resource-group ${resourceGroup} ` +
        `--deployment-name ${deploymentName} ` +
        `--model-name ${modelName} ` +
        `--model-version "${modelVersion}" ` +
        `--model-format "${modelFormat}" ` +
        `--sku-capacity ${skuCapacity} ` +
        `--sku-name ${skuName} ` +
        `--output none`,
      { timeout: 180_000 }
    );
    logOk(`Deployed: ${modelName} → ${deploymentName}`);
    return modelName;
  } catch (primaryError) {
    // Primary model deployment failed — could be region availability, quota, or other issue
    const errorDetail = primaryError.message.split("\n").pop();
    logFail(`${modelName} deployment failed in ${location}.`);
    logInfo(`Reason: ${errorDetail}`);

    const options = [];
    if (config.fallback) {
      options.push(`1. Use fallback: ${config.fallback.modelName}`);
    }
    options.push(`${config.fallback ? "2" : "1"}. Skip ${config.name} (${config.purpose})`);
    options.push(`${config.fallback ? "3" : "2"}. Abort — re-deploy in a different region`);

    log("");
    options.forEach((opt) => log(`    ${opt}`, colors.dim));
    log("");

    const defaultChoice = config.fallback ? "1" : "2";
    const choice = await ask(`What would you like to do?`, defaultChoice);

    if (config.fallback) {
      // Options: 1=fallback, 2=skip, 3=abort
      if (choice === "1") {
        const fb = config.fallback;
        logInfo(`Deploying fallback: ${fb.modelName}...`);
        try {
          exec(
            `az cognitiveservices account deployment create ` +
              `--name ${foundryName} ` +
              `--resource-group ${resourceGroup} ` +
              `--deployment-name ${deploymentName} ` +
              `--model-name ${fb.modelName} ` +
              `--model-version "${fb.modelVersion}" ` +
              `--model-format "OpenAI" ` +
              `--sku-capacity ${skuCapacity} ` +
              `--sku-name ${skuName} ` +
              `--output none`,
            { timeout: 180_000 }
          );
          logOk(`Deployed fallback: ${fb.modelName} → ${deploymentName}`);
          return fb.modelName;
        } catch (fallbackError) {
          logFail(`Fallback ${fb.modelName} also unavailable.`);
          logInfo(`Error: ${fallbackError.message.split("\n").pop()}`);
          logInfo("Try a different region or deploy the model manually in Azure AI Foundry.");
          closePrompt();
          process.exit(1);
        }
      } else if (choice === "2") {
        logInfo(`Skipping ${config.name}.`);
        return null;
      } else {
        log("\n  Deployment aborted. Re-run with a different region.", colors.yellow);
        logInfo("Tip: swedencentral and eastus2 have the widest model availability.");
        closePrompt();
        process.exit(0);
      }
    } else {
      // No fallback: 1=skip, 2=abort
      if (choice === "1") {
        logInfo(`Skipping ${config.name}.`);
        return null;
      } else {
        log("\n  Deployment aborted. Re-run with a different region.", colors.yellow);
        logInfo("Tip: swedencentral and eastus2 have the widest model availability.");
        closePrompt();
        process.exit(0);
      }
    }
  }
}

async function assignRoleWithRetry({
  principalId,
  roleName,
  scope,
  label,
}) {
  const maxAttempts = 5;
  const delaySeconds = [0, 15, 15, 20, 30];

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    if (delaySeconds[attempt] > 0) {
      logInfo(
        `Waiting ${delaySeconds[attempt]}s for identity propagation (${label}, attempt ${attempt + 1}/${maxAttempts})...`
      );
      await new Promise((r) => setTimeout(r, delaySeconds[attempt] * 1000));
    }

    try {
      exec(
        `az role assignment create ` +
          `--assignee ${principalId} ` +
          `--role "${roleName}" ` +
          `--scope ${scope} ` +
          `--output none`,
        { silent: true }
      );
      logOk(`Assigned '${roleName}' role (${label})`);
      return;
    } catch (error) {
      if (error.message.includes("already exists")) {
        logSkip(`Role assignment (${label})`);
        return;
      }
      if (error.message.includes("Cannot find user or service principal") && attempt < maxAttempts - 1) {
        continue;
      }
      throw error;
    }
  }

  throw new Error(`Role assignment failed after retries (${label}).`);
}

// =============================================================================
// MAIN
// =============================================================================

async function main() {
  banner("Grimoire — Azure Foundry Deployment");

  // ── Step 1: Check prerequisites ──────────────────────────────
  logStep(1, "Checking prerequisites");

  requireCommand("node", "https://nodejs.org");
  requireCommand("az", "https://learn.microsoft.com/cli/azure/install-azure-cli");
  requireCommand("func", "https://learn.microsoft.com/azure/azure-functions/functions-run-local");

  // ── Step 2: Check Azure login ────────────────────────────────
  logStep(2, "Checking Azure login");

  let account;
  try {
    account = JSON.parse(exec("az account show --output json", { silent: true }));
    logOk(`Logged in as: ${account.user?.name || "unknown"}`);
    logInfo(`Subscription: ${account.name} (${account.id})`);
  } catch {
    logInfo("Not logged in. Opening browser for Azure login...");
    execLive("az login");
    account = JSON.parse(exec("az account show --output json", { silent: true }));
    logOk(`Logged in as: ${account.user?.name || "unknown"}`);
  }
  const subscriptionId = account.id;
  const tenantId = account.tenantId;

  // ── Step 3: Gather configuration ─────────────────────────────
  logStep(3, "Configuration");

  // Load previous deployment config for defaults
  const previousConfig = loadDeployConfig();
  if (previousConfig) {
    logInfo(`Found previous deployment config (${CONFIG_FILENAME})`);
    logInfo(`Previous project: ${previousConfig.prefix || previousConfig.projectName || "grimoire"}`);
    log("");
  }

  // ── Phase A: Project naming ──────────────────────────────────
  log("  Project Naming:\n", colors.bold);
  log("  Choose a project name — this becomes the prefix for all Azure resources.", colors.dim);
  log("  Example: \"atlas\" creates atlas-rg, atlas-foundry, atlas-proxy, atlas-reasoning\n", colors.dim);

  const defaultProjectName = previousConfig?.projectName || "grimoire";
  const projectName = (await ask("Project name", defaultProjectName)).toLowerCase().replace(/[^a-z0-9-]/g, "");

  if (!projectName) {
    logFail("Project name is required.");
    closePrompt();
    process.exit(1);
  }

  log("\n  Add a unique suffix to avoid Azure name collisions (recommended).", colors.dim);
  log("  Press Enter for a random 4-char suffix, type your own, or \"none\" to skip.\n", colors.dim);

  const defaultSuffix = previousConfig?.suffix || "auto";
  const suffixInput = await ask("Suffix", defaultSuffix);

  let suffix = "";
  if (suffixInput === "none" || suffixInput === "skip" || suffixInput === "n") {
    suffix = "";
  } else if (suffixInput === "auto") {
    suffix = randomBytes(2).toString("hex");
  } else {
    suffix = suffixInput.toLowerCase().replace(/[^a-z0-9]/g, "").slice(0, 8);
    if (suffix.length < 2) {
      logFail("Suffix must be at least 2 alphanumeric characters (or use \"none\" to skip).");
      closePrompt();
      process.exit(1);
    }
  }

  const prefix = suffix ? `${projectName}-${suffix}` : projectName;
  // Display name must match the SPFx manifest webApiPermissionRequests resource name.
  // Always "Grimoire Backend API" regardless of prefix/suffix to avoid permission mismatch.
  const backendApiAppDisplayName = "Grimoire Backend API";

  if (suffix) {
    logOk(`Using suffix: ${suffix}`);
  } else {
    logInfo("No suffix — resource names may conflict if already taken.");
  }
  logOk(`Resource prefix: ${prefix}`);
  log("");

  // ── Phase B: Resource names (derived defaults, overridable) ──
  log("  Resource Names:\n", colors.bold);
  log("  Press Enter for derived defaults or override:\n", colors.dim);

  const defaultResourceGroup = previousConfig?.resourceGroup || `${prefix}-rg`;
  const defaultLocation = previousConfig?.location || DEFAULT_LOCATION;
  const defaultFoundryName = previousConfig?.foundryName || `${prefix}-foundry`;
  const defaultFunctionAppName = previousConfig?.functionAppName || `${prefix}-proxy`;

  const resourceGroup = await ask("Resource Group", defaultResourceGroup);
  const location = await ask("Azure Region", defaultLocation);
  const foundryName = await ask("Foundry resource name", defaultFoundryName);
  const functionAppName = await ask("Function App name", defaultFunctionAppName);

  // Storage account: must be lowercase alphanumeric, 3-24 chars
  const defaultStorage = previousConfig?.storageName || (functionAppName.replace(/[^a-z0-9]/g, "").slice(0, 20) + "stor");
  const storageName = await ask("Storage Account name", defaultStorage);

  // Resolve proxy key later (reuse existing key on re-deploy to avoid breaking clients).
  let proxyKey = "";
  let backendApiConfig = undefined;

  let sharepointOrigin = "";
  const defaultCorsOrigin = previousConfig?.sharepointOrigin || "";
  while (!sharepointOrigin) {
    const sharepointOriginInput = await ask(
      "SharePoint origin for CORS (e.g., https://contoso.sharepoint.com)",
      defaultCorsOrigin
    );

    if (!sharepointOriginInput.trim()) {
      logFail("SharePoint origin is required for deployed CORS validation.");
      continue;
    }

    try {
      sharepointOrigin = normalizeOrigin(sharepointOriginInput, "SharePoint origin for CORS");
    } catch (error) {
      logFail(error.message);
    }
  }

  // ── Model choices ──────────────────────────────────────────────
  log("\n  Model Selection:\n", colors.bold);
  log("  Choose which models to deploy. Press Enter for recommended defaults.\n", colors.dim);

  // Primary model (required)
  log("  Primary model (cinema mode, AI Builder):", colors.dim);
  log("    1. gpt-5-mini (recommended)", colors.dim);
  log("    2. gpt-4.1-mini", colors.dim);
  const primaryChoice = await ask("Primary model", "gpt-5-mini");
  const primaryModel = primaryChoice === "2" || primaryChoice === "gpt-4.1-mini" ? "gpt-4.1-mini" : "gpt-5-mini";

  // Fast model (optional)
  log("\n  Fast model (hints, validation, classification):", colors.dim);
  log("    1. gpt-5-nano (recommended)", colors.dim);
  log("    2. gpt-4.1-nano", colors.dim);
  log("    3. skip (proxy works fine without it)", colors.dim);
  const fastChoice = await ask("Fast model", "gpt-5-nano");
  const fastModel = fastChoice === "3" || fastChoice.toLowerCase() === "skip" ? null
    : fastChoice === "2" || fastChoice === "gpt-4.1-nano" ? "gpt-4.1-nano"
    : "gpt-5-nano";

  // Realtime model (optional)
  log("\n  Realtime voice model (voice-to-agent):", colors.dim);
  log("    1. gpt-realtime-1.5 (recommended — Feb 2026, +5% reasoning, +10% transcription)", colors.dim);
  log("    2. gpt-realtime (GA baseline — Aug 2025)", colors.dim);
  log("    3. skip (voice sessions will be unavailable)", colors.dim);
  const realtimeChoice = await ask("Realtime model", "gpt-realtime-1.5");
  const realtimeModel = realtimeChoice === "3" || realtimeChoice.toLowerCase() === "skip" ? null
    : realtimeChoice === "2" || realtimeChoice === "gpt-realtime" ? "gpt-realtime"
    : "gpt-realtime-1.5"; // covers "1", default, or full model name

  let realtimeCapacity = 6;
  if (realtimeModel) {
    const capInput = await ask("Realtime capacity (RPM, concurrent sessions)", "6");
    realtimeCapacity = Math.max(1, parseInt(capInput, 10) || 6);
  }

  // Build model configs from choices
  const MODELS = buildModels({ prefix, primaryModel, fastModel, realtimeModel, realtimeCapacity });

  log("\n  Configuration summary:", colors.bold);
  logInfo(`Project:            ${projectName}${suffix ? ` (suffix: ${suffix})` : ""}`);
  logInfo(`Resource prefix:    ${prefix}`);
  logInfo(`Resource Group:     ${resourceGroup}`);
  logInfo(`Region:             ${location}`);
  logInfo(`Foundry resource:   ${foundryName} (kind: AIServices)`);
  logInfo(`Function App:       ${functionAppName}`);
  logInfo(`Storage:            ${storageName}`);
  logInfo(`Backend API:        ${backendApiAppDisplayName}`);
  logInfo(`Primary model:      ${primaryModel}`);
  logInfo(`Fast model:         ${fastModel || "skip"}`);
  logInfo(`Realtime voice:     ${realtimeModel || "skip"}`);
  logInfo(`CORS origin:        ${sharepointOrigin}`);
  log("");

  const confirm = await ask("Proceed with deployment? (yes/no)", "yes");
  if (confirm.toLowerCase() !== "yes" && confirm.toLowerCase() !== "y") {
    log("\n  Deployment cancelled.", colors.yellow);
    closePrompt();
    process.exit(0);
  }

  // ── Step 4: Create Resource Group ────────────────────────────
  logStep(4, "Resource Group");

  const rgExists = exec(
    `az group exists --name ${resourceGroup}`,
    { silent: true }
  );

  if (rgExists === "true") {
    logSkip(resourceGroup);
  } else {
    exec(`az group create --name ${resourceGroup} --location ${location} --output none`);
    logOk(`Created resource group: ${resourceGroup}`);
  }

  // ── Step 5: Create Foundry Resource (AIServices) ──────────────
  logStep(5, "Microsoft Foundry Resource");

  let foundryExists = false;
  try {
    exec(
      `az cognitiveservices account show --name ${foundryName} --resource-group ${resourceGroup} --output none`,
      { silent: true }
    );
    foundryExists = true;
    logSkip(foundryName);
  } catch {
    // Doesn't exist — create it
  }

  if (!foundryExists) {
    logInfo("Creating Foundry resource (kind: AIServices). This may take 1-2 minutes...");
    logInfo("AIServices gives access to OpenAI models + Foundry catalog (Mistral, Meta, etc.)");
    try {
      exec(
        `az cognitiveservices account create ` +
          `--name ${foundryName} ` +
          `--resource-group ${resourceGroup} ` +
          `--location ${location} ` +
          `--kind AIServices ` +
          `--sku S0 ` +
          `--custom-domain ${foundryName} ` +
          `--yes ` +
          `--output none`,
        { timeout: 180_000 }
      );
      logOk(`Created Foundry resource: ${foundryName}`);
    } catch (error) {
      logFail("Failed to create Foundry resource.");
      logInfo("This may be due to region restrictions or quota limits.");
      logInfo("Try a different region or check: https://learn.microsoft.com/azure/ai-services/multi-service-resource");
      logInfo(`Error: ${error.message}`);
      process.exit(1);
    }
  }

  // Get the Foundry endpoint
  const foundryEndpoint = exec(
    `az cognitiveservices account show --name ${foundryName} --resource-group ${resourceGroup} --query properties.endpoint --output tsv`,
    { silent: true }
  );
  logOk(`Endpoint: ${foundryEndpoint}`);

  // ── Step 6: Deploy Models ─────────────────────────────────────
  logStep(6, "Model Deployments");

  log("\n  Deploying models to Foundry resource:\n", colors.dim);

  const actualModels = {};

  // Deploy reasoning model (required — cinema mode won't work without it)
  const reasoningResult = await deployModel(foundryName, resourceGroup, MODELS.reasoning, location);
  actualModels.reasoning = reasoningResult;

  // Deploy fast model (optional — falls back gracefully)
  if (MODELS.fast) {
    const fastResult = await deployModel(foundryName, resourceGroup, MODELS.fast, location);
    actualModels.fast = fastResult;
  } else {
    logSkip("Fast model (skipped by user)");
  }

  // Deploy realtime model (optional — voice session feature)
  if (MODELS.realtime) {
    const realtimeResult = await deployModel(foundryName, resourceGroup, MODELS.realtime, location);
    actualModels.realtime = realtimeResult;
  } else {
    logSkip("Realtime model (skipped by user)");
  }

  closePrompt();

  // ── Step 7: Create Storage Account ───────────────────────────
  logStep(7, "Storage Account");

  let storageExists = false;
  try {
    exec(
      `az storage account show --name ${storageName} --resource-group ${resourceGroup} --output none`,
      { silent: true }
    );
    storageExists = true;
    logSkip(storageName);
  } catch {
    // Create it
  }

  if (!storageExists) {
    exec(
      `az storage account create ` +
        `--name ${storageName} ` +
        `--resource-group ${resourceGroup} ` +
        `--location ${location} ` +
        `--sku Standard_LRS ` +
        `--output none`,
      { timeout: 120_000 }
    );
    logOk(`Created storage account: ${storageName}`);
  }

  // ── Step 8: Create Function App ──────────────────────────────
  logStep(8, "Function App");

  let funcExists = false;
  try {
    exec(
      `az functionapp show --name ${functionAppName} --resource-group ${resourceGroup} --output none`,
      { silent: true }
    );
    funcExists = true;
    logSkip(functionAppName);
  } catch {
    // Create it
  }

  if (!funcExists) {
    exec(
      `az functionapp create ` +
        `--name ${functionAppName} ` +
        `--resource-group ${resourceGroup} ` +
        `--storage-account ${storageName} ` +
        `--consumption-plan-location ${location} ` +
        `--runtime node ` +
        `--runtime-version 22 ` +
        `--functions-version 4 ` +
        `--os-type linux ` +
        `--output none`,
      { timeout: 180_000 }
    );
    logOk(`Created function app: ${functionAppName}`);
  }

  // Auth is handled by Azure Functions built-in function key auth (authLevel: "function").
  // Function keys are managed via Azure Portal or CLI — no custom ALLOWED_KEYS env var needed.
  // Retrieve the default function key for display/SPFx web part configuration.
  const functionKeyRaw = exec(
    `az functionapp keys list ` +
      `--name ${functionAppName} ` +
      `--resource-group ${resourceGroup} ` +
      `--query "functionKeys.default" ` +
      `--output tsv`,
    { silent: true, ignoreError: true }
  ).trim();
  if (functionKeyRaw && functionKeyRaw !== "null") {
    proxyKey = functionKeyRaw;
    logInfo(`Using existing Azure function key: ${proxyKey.slice(0, 16)}...`);
  } else {
    logInfo(`Function key will be auto-generated by Azure on first request.`);
    proxyKey = "(retrieve from Azure Portal after deployment)";
  }

  // ── Step 9: Backend API registration ─────────────────────────
  logStep(9, "Backend API Registration");

  try {
    backendApiConfig = ensureBackendApiApplication(backendApiAppDisplayName, projectName);
    logInfo(`Backend API resource: ${backendApiConfig.resource}`);
  } catch (error) {
    logFail(`Backend API registration failed: ${error.message}`);
    process.exit(1);
  }

  // Grant SPFx service principal (SharePoint Online Web Client Extensibility)
  // access to the backend API — this is what SharePoint admin "API access" approval does.
  ensureSpfxPermissionGrant(backendApiConfig.appId);

  // ── Step 10: Enable managed identity ────────────────────────
  logStep(10, "Managed Identity");

  const identityJson = exec(
    `az functionapp identity assign ` +
      `--name ${functionAppName} ` +
      `--resource-group ${resourceGroup} ` +
      `--output json`,
    { silent: true }
  );
  const identity = JSON.parse(identityJson);
  const principalId = identity.principalId;
  logOk(`System-assigned identity: ${principalId}`);

  // ── Step 11: Assign Cognitive Services roles ────────────────
  logStep(11, "Role Assignment");

  const foundryResourceId = exec(
    `az cognitiveservices account show --name ${foundryName} --resource-group ${resourceGroup} --query id --output tsv`,
    { silent: true }
  );

  try {
    await assignRoleWithRetry({
      principalId,
      roleName: "Cognitive Services OpenAI User",
      scope: foundryResourceId,
      label: "Foundry"
    });
  } catch (error) {
    logFail(`Role assignment failed: ${error.message}`);
    logInfo("You may need Owner or User Access Administrator permissions.");
    process.exit(1);
  }

  // ── Step 12: Deploy function code ───────────────────────────
  logStep(12, "Deploy Proxy Code");

  logInfo("Publishing to Azure with remote build...");
  execLive(`func azure functionapp publish ${functionAppName} --javascript --build remote`);
  logOk("Proxy code deployed");

  // Configure platform-level CORS (Azure Functions host)
  // This is required because Azure's platform CORS runs BEFORE application code.
  // Without it, OPTIONS preflight requests are blocked at the platform level.
  logInfo(`Configuring platform-level CORS for ${sharepointOrigin}...`);
  try {
    const existingCorsOriginsRaw = exec(
      `az functionapp cors show ` +
        `--name ${functionAppName} ` +
        `--resource-group ${resourceGroup} ` +
        `--query "allowedOrigins[]" ` +
        `--output tsv`,
      { silent: true, ignoreError: true }
    ).trim();
    const existingCorsOrigins = existingCorsOriginsRaw
      ? existingCorsOriginsRaw.split(/\r?\n/).map((origin) => origin.trim()).filter(Boolean)
      : [];

    for (const origin of new Set([...existingCorsOrigins, "*"])) {
      if (origin === sharepointOrigin) continue;
      exec(
        `az functionapp cors remove ` +
          `--name ${functionAppName} ` +
          `--resource-group ${resourceGroup} ` +
          `--allowed-origins "${origin}" ` +
          `--output none`,
        { silent: true, ignoreError: true }
      );
    }

    if (!existingCorsOrigins.includes(sharepointOrigin)) {
      exec(
        `az functionapp cors add ` +
          `--name ${functionAppName} ` +
          `--resource-group ${resourceGroup} ` +
          `--allowed-origins "${sharepointOrigin}" ` +
          `--output none`,
        { silent: true }
      );
    }

    logOk(`Platform CORS configured: ${sharepointOrigin}`);
  } catch (corsError) {
    logInfo(`CORS config warning: ${corsError.message.split("\\n").pop()}`);
    logInfo(`You may need to set CORS manually in the Azure Portal (API > CORS > add ${sharepointOrigin})`);
  }

  // ── Step 13: Configure app settings ─────────────────────────
  logStep(13, "App Settings");

  // Build backend list — only include models that were successfully deployed
  const backendNames = ["reasoning"];
  const realtimeDeploymentName = MODELS.realtime ? MODELS.realtime.deploymentName : `${prefix}-realtime`;
  const settings = [
    `BACKEND_reasoning_ENDPOINT=${foundryEndpoint}`,
    `ALLOW_PERMISSIVE_LOCAL_CORS=false`,
    `ALLOWED_ORIGINS=${sharepointOrigin}`,
    `REQUESTS_PER_MINUTE=60`,
    `REQUESTS_PER_DAY=5000`,
    // Required so the proxy can forward the Entra bearer token to Agent365 MCP servers.
    `MCP_TOKEN_FORWARD_ALLOWLIST=agent365.svc.cloud.microsoft`,
    `REALTIME_DEPLOYMENT_NAME=${realtimeDeploymentName}`,
  ];

  if (actualModels.fast) {
    backendNames.push("fast");
    settings.push(`BACKEND_fast_ENDPOINT=${foundryEndpoint}`);
  }

  settings.unshift(`BACKENDS=${backendNames.join(",")}`);
  exec(
    `az functionapp config appsettings set ` +
      `--name ${functionAppName} ` +
      `--resource-group ${resourceGroup} ` +
      `--settings ${settings.map((s) => `"${s}"`).join(" ")} ` +
      `--output none`,
    { silent: true }
  );
  logOk("App settings configured");

  // Note: no BACKEND_*_KEY set → proxy will use managed identity

  // ── Step 14: Configure Easy Auth ────────────────────────────
  logStep(14, "Easy Auth");

  try {
    configureFunctionAppEasyAuth({
      subscriptionId,
      tenantId,
      resourceGroup,
      functionAppName,
      backendApiAppId: backendApiConfig.appId,
      backendApiResource: backendApiConfig.resource,
    });
    logOk("Function App Easy Auth configured for /api/user/*");
  } catch (error) {
    logFail(`Easy Auth configuration failed: ${error.message}`);
    process.exit(1);
  }

  // ── Step 15: Verify ─────────────────────────────────────────
  logStep(15, "Verification");

  const proxyUrl = `https://${functionAppName}.azurewebsites.net/api`;
  logInfo("Waiting 10s for Azure to propagate settings...");
  await new Promise((r) => setTimeout(r, 10_000));

  try {
    const healthRes = await httpFetch(`${proxyUrl}/health`);
    if (healthRes.status === 200) {
      const health = JSON.parse(healthRes.body);
      logOk(`Health check passed: ${health.backends ? Object.keys(health.backends).join(", ") : "ok"}`);
    } else {
      logFail(`Health check returned ${healthRes.status}`);
      logInfo("The deployment may need a few more minutes to warm up. Try again shortly.");
    }
  } catch (error) {
    logFail(`Health check failed: ${error.message}`);
    logInfo("The function app may need a minute to start. Try: curl " + proxyUrl + "/health");
  }

  // ── Save deployment config ─────────────────────────────────
  const deploymentNames = { reasoning: MODELS.reasoning.deploymentName };
  if (MODELS.fast) deploymentNames.fast = MODELS.fast.deploymentName;
  if (MODELS.realtime) deploymentNames.realtime = MODELS.realtime.deploymentName;

  saveDeployConfig({
    projectName,
    suffix: suffix || "",
    prefix,
    resourceGroup,
    location,
    foundryName,
    functionAppName,
    storageName,
    sharepointOrigin,
    proxyUrl,
    proxyApiKey: proxyKey,
    deploymentNames,
    backendApiDisplayName: backendApiAppDisplayName,
    backendApiResource: backendApiConfig.resource,
    deployedAt: new Date().toISOString(),
  });
  logOk(`Saved deployment config to ${CONFIG_FILENAME}`);

  // ── Summary ──────────────────────────────────────────────────
  banner("Deployment Complete!");

  log(`  Your ${projectName} proxy is live:\n`, colors.bold);
  logInfo(`Proxy URL:          ${proxyUrl}`);
  logInfo(`Proxy API Key:      ${proxyKey}`);
  logInfo(`Backend API App:    ${backendApiConfig.displayName}`);
  logInfo(`Backend API URI:    ${backendApiConfig.resource}`);
  logInfo(`Allowed CORS:       ${sharepointOrigin}`);
  logInfo(`Foundry resource:   ${foundryName} (kind: AIServices)`);
  logInfo(`Auth:               Managed Identity + Easy Auth for /api/user/*`);
  logInfo(`Resource Group:     ${resourceGroup}`);
  logInfo(`Deployment prefix:  ${prefix}`);

  log("\n  Models deployed:\n", colors.bold);
  logInfo(`Reasoning:          ${actualModels.reasoning || "NOT DEPLOYED"} → ${MODELS.reasoning.deploymentName}`);
  logInfo(`                    Purpose: ${MODELS.reasoning.purpose}`);
  if (MODELS.fast) {
    if (actualModels.fast) {
      logInfo(`Fast:               ${actualModels.fast} → ${MODELS.fast.deploymentName}`);
      logInfo(`                    Purpose: ${MODELS.fast.purpose}`);
    } else {
      logInfo(`Fast:               SKIPPED (not available in ${location})`);
      logInfo(`                    The proxy works fine without it.`);
    }
  } else {
    logInfo(`Fast:               SKIPPED (by user choice)`);
  }
  if (MODELS.realtime) {
    if (actualModels.realtime) {
      logInfo(`Realtime:           ${actualModels.realtime} → ${MODELS.realtime.deploymentName}`);
      logInfo(`                    Purpose: ${MODELS.realtime.purpose}`);
    } else {
      logInfo(`Realtime:           SKIPPED (not available in ${location})`);
      logInfo(`                    Voice sessions will be unavailable.`);
    }
  } else {
    logInfo(`Realtime:           SKIPPED (by user choice)`);
  }
  log("\n  API routes:\n", colors.bold);
  logInfo(`Chat (Reasoning):   ${proxyUrl}/reasoning/openai/deployments/${MODELS.reasoning.deploymentName}/chat/completions`);
  if (actualModels.fast && MODELS.fast) {
    logInfo(`Chat (Fast):        ${proxyUrl}/fast/openai/deployments/${MODELS.fast.deploymentName}/chat/completions`);
  }
  logInfo(`Health:             ${proxyUrl}/health`);
  logInfo(`MCP Discovery:      ${proxyUrl}/mcp/discover`);
  logInfo(`Realtime Token:     ${proxyUrl}/realtime/token`);

  log("\n  Next steps:\n", colors.bold);
  logInfo("1. Open your SharePoint site with the Grimoire web part");
  logInfo(`2. Request tenant admin consent for ${backendApiAppDisplayName} → ${BACKEND_API_SCOPE_VALUE}`);
  logInfo("3. Edit the web part → Property Pane → AI Builder section");
  logInfo(`4. Set Proxy URL:             ${proxyUrl}`);
  logInfo(`5. Set Proxy API Key:         ${proxyKey}`);
  logInfo(`6. Set Backend API Resource:  ${backendApiConfig.resource}`);
  logInfo(`7. Set Deployment prefix:     ${prefix}`);
  logInfo("8. Set Backend:               reasoning (or fast for lightweight tasks)");
  logInfo("9. Describe your agent and watch it build!");

  log(`\n  To run smoke tests:  npm run test:proxy -- ${proxyUrl} ${proxyKey} ${sharepointOrigin}`, colors.dim);
  log(`  To remove everything:  npm run teardown\n`, colors.dim);
}

main().catch((error) => {
  logFail(error.message);
  closePrompt();
  process.exit(1);
});
