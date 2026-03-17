#!/usr/bin/env node

/**
 * Grimoire Proxy — Teardown Script
 *
 * Removes ALL Azure resources created by deploy.mjs in one command.
 * Deletes the entire resource group and everything inside it
 * (Foundry resource, model deployments, Function App, Storage).
 *
 * Auto-detects resource group from .grimoire-deploy.json if available.
 *
 * Usage:  node scripts/teardown.mjs
 *    or:  npm run teardown
 */

import {
  banner,
  logStep,
  logOk,
  logFail,
  logInfo,
  log,
  colors,
  exec,
  execLive,
  ask,
  closePrompt,
  requireCommand,
  capitalize,
} from "./lib.mjs";
import { loadDeployConfig, deleteDeployConfig, CONFIG_FILENAME } from "./deploy-config.mjs";

const config = loadDeployConfig();
const defaultResourceGroup = config?.resourceGroup || "grimoire-rg";
const projectLabel = config?.projectName ? capitalize(config.projectName) : "Grimoire";

async function main() {
  banner(`${projectLabel} — Teardown`);

  // ── Check prerequisites ──────────────────────────────────────
  logStep(1, "Checking prerequisites");
  requireCommand("az", "https://learn.microsoft.com/cli/azure/install-azure-cli");

  // ── Check Azure login ────────────────────────────────────────
  try {
    exec("az account show --output none", { silent: true });
    logOk("Azure CLI logged in");
  } catch {
    logFail("Not logged in to Azure CLI. Run: az login");
    process.exit(1);
  }

  // ── Confirm ──────────────────────────────────────────────────
  logStep(2, "Configuration");

  if (config) {
    logInfo(`Detected deployment: ${config.prefix || config.projectName || "grimoire"}`);
  }

  const resourceGroup = await ask("Resource Group to delete", defaultResourceGroup);

  // Show what will be deleted
  log("", colors.reset);
  try {
    const resources = exec(
      `az resource list --resource-group ${resourceGroup} --query "[].{name:name, type:type}" --output table`,
      { silent: true }
    );
    log("  Resources that will be deleted:\n", colors.yellow);
    log(resources, colors.dim);
  } catch {
    logFail(`Resource group '${resourceGroup}' not found.`);
    closePrompt();
    process.exit(1);
  }

  log("", colors.reset);
  log("  ⚠  This will permanently delete ALL resources above!", colors.red);
  const confirm = await ask("Type the resource group name to confirm", "");

  if (confirm !== resourceGroup) {
    log("\n  Teardown cancelled (name didn't match).", colors.yellow);
    closePrompt();
    process.exit(0);
  }

  closePrompt();

  // ── Delete ───────────────────────────────────────────────────
  logStep(3, "Deleting Resource Group");

  logInfo(`Deleting '${resourceGroup}' and all resources (this may take 2-5 minutes)...`);
  execLive(`az group delete --name ${resourceGroup} --yes --no-wait`);
  logOk(`Deletion initiated for resource group: ${resourceGroup}`);
  logInfo("Azure is deleting resources in the background. This takes a few minutes.");

  // Clean up local config if it matches the deleted resource group
  if (config && config.resourceGroup === resourceGroup) {
    deleteDeployConfig();
    logOk(`Removed local deployment config (${CONFIG_FILENAME})`);
  }

  banner("Teardown Complete");
  logInfo(`All ${projectLabel} proxy resources are being deleted.`);
  logInfo("Check status: az group show --name " + resourceGroup);

  // ── Post-teardown warnings ─────────────────────────────────
  log("\n  ⚠  Important post-teardown steps:\n", colors.yellow);

  log("  1. Entra app registrations are NOT deleted by resource group teardown.", colors.yellow);
  log("     They live at the tenant level. To clean up manually:", colors.yellow);
  log(`     az ad app list --display-name "Grimoire" --query "[].{name:displayName, appId:appId}" -o table`, colors.dim);
  log("     az ad app delete --id <appId>\n", colors.dim);

  log("  2. Cognitive Services (Foundry) resources are soft-deleted for 48 hours.", colors.yellow);
  log("     They continue to consume model deployment quota until purged.", colors.yellow);
  log("     To list soft-deleted resources:", colors.yellow);
  log(`     az cognitiveservices account list-deleted --query "[?location=='swedencentral']" -o table`, colors.dim);
  log("     To purge and free quota:", colors.yellow);
  log(`     az cognitiveservices account purge --name <name> --resource-group ${resourceGroup} --location swedencentral\n`, colors.dim);
}

main().catch((error) => {
  logFail(error.message);
  closePrompt();
  process.exit(1);
});
