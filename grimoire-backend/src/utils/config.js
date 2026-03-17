/**
 * Configuration — loads from environment variables.
 */

function parseCsvList(value) {
  return (value || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function parseBooleanFlag(value) {
  return String(value || "").trim().toLowerCase() === "true";
}

// Rate limits
export const REQUESTS_PER_MINUTE = parseInt(process.env.REQUESTS_PER_MINUTE || "60");
export const REQUESTS_PER_DAY = parseInt(process.env.REQUESTS_PER_DAY || "5000");

// CORS origins
export const ALLOWED_ORIGINS = parseCsvList(process.env.ALLOWED_ORIGINS);
export const ALLOW_PERMISSIVE_LOCAL_CORS = parseBooleanFlag(process.env.ALLOW_PERMISSIVE_LOCAL_CORS);

// MCP session config
export const MCP_SESSION_TIMEOUT_MINUTES = parseInt(process.env.MCP_SESSION_TIMEOUT_MINUTES || "30");
export const MAX_MCP_SESSIONS = parseInt(process.env.MAX_MCP_SESSIONS || "20");
export const MCP_ALLOWED_HOSTS = parseCsvList(process.env.MCP_ALLOWED_HOSTS).map((host) => host.toLowerCase());
export const MCP_TOKEN_FORWARD_ALLOWLIST = parseCsvList(process.env.MCP_TOKEN_FORWARD_ALLOWLIST).map((host) => host.toLowerCase());

// Outbound timeout config
export const LLM_UPSTREAM_TIMEOUT_MS = parseInt(process.env.LLM_UPSTREAM_TIMEOUT_MS || "30000");
export const REALTIME_TOKEN_TIMEOUT_MS = parseInt(process.env.REALTIME_TOKEN_TIMEOUT_MS || "15000");

// Realtime deployment name (set by deploy.mjs based on project prefix)
export const REALTIME_DEPLOYMENT_NAME = process.env.REALTIME_DEPLOYMENT_NAME || "grimoire-realtime";

/**
 * Load multi-backend configuration from env vars.
 * Pattern: BACKENDS=reasoning,fast
 *          BACKEND_reasoning_ENDPOINT=https://...
 *          BACKEND_reasoning_KEY=... (optional if using managed identity)
 */
export function loadBackends() {
  const backendNames = (process.env.BACKENDS || "")
    .split(",")
    .map((b) => b.trim())
    .filter(Boolean);
  const backends = {};

  for (const name of backendNames) {
    const endpoint = process.env[`BACKEND_${name}_ENDPOINT`];
    const key = process.env[`BACKEND_${name}_KEY`] || "";

    if (endpoint) {
      backends[name] = { endpoint: endpoint.replace(/\/$/, ""), key };
    }
  }

  return backends;
}

export const BACKENDS = loadBackends();

export function getCorsMode(config = {
  allowPermissiveLocalCors: ALLOW_PERMISSIVE_LOCAL_CORS,
}) {
  return config.allowPermissiveLocalCors ? "permissive-local" : "allowlist";
}

export function validateConfiguration(config = {
  allowedOrigins: ALLOWED_ORIGINS,
  allowPermissiveLocalCors: ALLOW_PERMISSIVE_LOCAL_CORS,
}) {
  if (!config.allowPermissiveLocalCors && config.allowedOrigins.length === 0) {
    throw new Error("ALLOWED_ORIGINS must contain at least one origin unless ALLOW_PERMISSIVE_LOCAL_CORS=true.");
  }
}
