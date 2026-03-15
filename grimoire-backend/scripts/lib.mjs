/**
 * Shared utilities for Grimoire proxy scripts.
 * Pure Node.js — no external dependencies.
 */

import { execSync } from "node:child_process";
import { createInterface } from "node:readline";
import { request as httpsRequest } from "node:https";
import { request as httpRequest } from "node:http";

// =============================================================================
// COLORS
// =============================================================================

const supportsColor =
  process.env.FORCE_COLOR !== "0" &&
  (process.env.FORCE_COLOR || process.stdout.isTTY);

export const colors = {
  reset: supportsColor ? "\x1b[0m" : "",
  green: supportsColor ? "\x1b[32m" : "",
  red: supportsColor ? "\x1b[31m" : "",
  yellow: supportsColor ? "\x1b[33m" : "",
  blue: supportsColor ? "\x1b[34m" : "",
  cyan: supportsColor ? "\x1b[36m" : "",
  dim: supportsColor ? "\x1b[2m" : "",
  bold: supportsColor ? "\x1b[1m" : "",
};

export function log(msg, color = "") {
  console.log(`${color}${msg}${colors.reset}`);
}

export function logStep(step, msg) {
  log(`\n[${"=".repeat(60)}]`, colors.cyan);
  log(`  Step ${step}: ${msg}`, colors.bold);
  log(`[${"=".repeat(60)}]`, colors.cyan);
}

export function logOk(msg) {
  log(`  ✓ ${msg}`, colors.green);
}

export function logSkip(msg) {
  log(`  → ${msg} (already exists, skipping)`, colors.yellow);
}

export function logFail(msg) {
  log(`  ✗ ${msg}`, colors.red);
}

export function logInfo(msg) {
  log(`  ${msg}`, colors.dim);
}

// =============================================================================
// EXEC
// =============================================================================

/**
 * Execute a shell command and return stdout. Throws on failure.
 */
export function exec(cmd, options = {}) {
  try {
    return execSync(cmd, {
      encoding: "utf-8",
      stdio: options.silent ? "pipe" : ["pipe", "pipe", "pipe"],
      timeout: options.timeout || 120_000,
      ...options,
    }).trim();
  } catch (error) {
    if (options.ignoreError) return "";
    const stderr = error.stderr?.toString().trim() || error.message;
    throw new Error(`Command failed: ${cmd}\n${stderr}`);
  }
}

/**
 * Execute and stream output to console (for long-running commands).
 */
export function execLive(cmd) {
  execSync(cmd, { stdio: "inherit", encoding: "utf-8" });
}

// =============================================================================
// PROMPT
// =============================================================================

let rlInstance = null;

function getRL() {
  if (!rlInstance) {
    rlInstance = createInterface({
      input: process.stdin,
      output: process.stdout,
    });
  }
  return rlInstance;
}

/**
 * Prompt the user for input. Returns the answer (or default if empty).
 */
export function ask(question, defaultValue = "") {
  return new Promise((resolve) => {
    const suffix = defaultValue ? ` [${defaultValue}]` : "";
    getRL().question(`  ${question}${suffix}: `, (answer) => {
      resolve(answer.trim() || defaultValue);
    });
  });
}

/**
 * Prompt for a secret (still visible — Node.js readline has no hidden mode).
 */
export async function askSecret(question) {
  return ask(question);
}

/**
 * Close the readline interface.
 */
export function closePrompt() {
  if (rlInstance) {
    rlInstance.close();
    rlInstance = null;
  }
}

// =============================================================================
// PREREQUISITES
// =============================================================================

/**
 * Check if a command is available. Returns version string or null.
 */
export function checkCommand(name) {
  try {
    const version = exec(`${name} --version`, { silent: true, ignoreError: false });
    return version.split("\n")[0];
  } catch {
    return null;
  }
}

/**
 * Require a command to be installed. Exits if not found.
 */
export function requireCommand(name, installUrl = "") {
  const version = checkCommand(name);
  if (!version) {
    logFail(`'${name}' is not installed.`);
    if (installUrl) logInfo(`Install: ${installUrl}`);
    process.exit(1);
  }
  logOk(`${name}: ${version}`);
  return version;
}

// =============================================================================
// HTTP REQUESTS
// =============================================================================

/**
 * Make an HTTP/HTTPS request. Returns { status, headers, body }.
 */
export function httpFetch(url, options = {}) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const reqFn = parsedUrl.protocol === "https:" ? httpsRequest : httpRequest;

    const req = reqFn(
      url,
      {
        method: options.method || "GET",
        headers: options.headers || {},
        timeout: options.timeout || 30_000,
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          resolve({
            status: res.statusCode,
            headers: res.headers,
            body: Buffer.concat(chunks).toString("utf-8"),
          });
        });
      }
    );

    req.on("error", reject);
    req.on("timeout", () => {
      req.destroy();
      reject(new Error(`Request timeout: ${url}`));
    });

    if (options.body) {
      req.write(options.body);
    }

    req.end();
  });
}

// =============================================================================
// VALIDATION
// =============================================================================

/**
 * Validate and normalize an HTTP(S) origin (scheme + host + optional port).
 * Throws if the value is not a valid origin.
 */
export function normalizeOrigin(value, label) {
  let url;
  try {
    url = new URL(String(value || "").trim());
  } catch {
    throw new Error(`${label} must be a valid http(s) origin.`);
  }

  if (url.protocol !== "http:" && url.protocol !== "https:") {
    throw new Error(`${label} must use http:// or https://.`);
  }

  if (url.username || url.password || url.search || url.hash || url.pathname !== "/") {
    throw new Error(`${label} must be an origin only (scheme, host, optional port) with no path, query, or fragment.`);
  }

  return url.origin;
}

/**
 * Capitalize the first letter of a string.
 */
export function capitalize(str) {
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : "";
}

// =============================================================================
// BANNER
// =============================================================================

export function banner(title) {
  const line = "═".repeat(60);
  log(`\n╔${line}╗`, colors.cyan);
  log(`║  ${title.padEnd(58)}║`, colors.cyan);
  log(`╚${line}╝\n`, colors.cyan);
}
