/**
 * Deploy config persistence — saves and loads deployment configuration
 * so that costs.mjs, teardown.mjs, test-proxy.mjs, and setup-local.mjs
 * can auto-detect resource names from a previous deployment.
 *
 * Config file: .grimoire-deploy.json (in grimoire-backend/)
 * Added to .gitignore because it contains the proxy API key.
 */

import { existsSync, readFileSync, writeFileSync, unlinkSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_DIR = resolve(__dirname, "..");

export const CONFIG_FILENAME = ".grimoire-deploy.json";
const CONFIG_PATH = resolve(PROJECT_DIR, CONFIG_FILENAME);

/**
 * Save deployment config after a successful deploy.
 */
export function saveDeployConfig(config) {
  writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2) + "\n");
}

/**
 * Load deployment config. Returns null if no config file exists.
 */
export function loadDeployConfig() {
  if (!existsSync(CONFIG_PATH)) return null;
  try {
    return JSON.parse(readFileSync(CONFIG_PATH, "utf-8"));
  } catch {
    return null;
  }
}

/**
 * Delete the config file (e.g. after teardown).
 */
export function deleteDeployConfig() {
  if (existsSync(CONFIG_PATH)) {
    unlinkSync(CONFIG_PATH);
  }
}

/**
 * Return the absolute path of the config file (for display).
 */
export function getConfigPath() {
  return CONFIG_PATH;
}
