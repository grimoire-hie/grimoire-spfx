/**
 * MCP target URL policy guard.
 * Blocks common SSRF targets while allowing public HTTPS endpoints.
 */

import { BlockList, isIP } from "node:net";
import { MCP_ALLOWED_HOSTS, MCP_TOKEN_FORWARD_ALLOWLIST } from "./config.js";

const BLOCKED_IPS = new BlockList();

// IPv4 loopback / private / link-local
BLOCKED_IPS.addSubnet("127.0.0.0", 8, "ipv4");
BLOCKED_IPS.addSubnet("10.0.0.0", 8, "ipv4");
BLOCKED_IPS.addSubnet("172.16.0.0", 12, "ipv4");
BLOCKED_IPS.addSubnet("192.168.0.0", 16, "ipv4");
BLOCKED_IPS.addSubnet("169.254.0.0", 16, "ipv4");

// IPv6 loopback / unique-local / link-local
BLOCKED_IPS.addAddress("::1", "ipv6");
BLOCKED_IPS.addSubnet("fc00::", 7, "ipv6");
BLOCKED_IPS.addSubnet("fe80::", 10, "ipv6");

const BLOCKED_HOSTNAMES = new Set([
  "localhost",
  "localhost.localdomain",
  "host.docker.internal",
]);

function normalizeHost(hostname) {
  return String(hostname || "")
    .replace(/^\[|\]$/g, "")
    .replace(/\.$/, "")
    .toLowerCase();
}

function normalizeHostPattern(pattern) {
  const input = String(pattern || "").trim().toLowerCase();
  if (!input) return "";

  const wildcard = input.startsWith("*.");
  const withoutWildcard = wildcard ? input.slice(2) : input;

  if (withoutWildcard.includes("://")) {
    try {
      const parsed = new URL(withoutWildcard);
      const host = normalizeHost(parsed.hostname);
      return wildcard ? `*.${host}` : host;
    } catch {
      return "";
    }
  }

  const host = normalizeHost(withoutWildcard);
  return wildcard ? `*.${host}` : host;
}

function matchesHostPattern(host, pattern) {
  if (!pattern) return false;
  if (pattern.startsWith("*.")) {
    const suffix = pattern.slice(2);
    return host === suffix || host.endsWith(`.${suffix}`);
  }
  return host === pattern;
}

function isBlockedHostname(host) {
  if (BLOCKED_HOSTNAMES.has(host)) return true;
  if (host.endsWith(".localhost")) return true;
  if (host.endsWith(".local")) return true;
  return false;
}

function isBlockedLiteralIp(host) {
  const version = isIP(host);
  if (version === 0) return false;

  if (version === 6 && host.startsWith("::ffff:")) {
    const mapped = host.slice("::ffff:".length);
    if (isIP(mapped) === 4) {
      return BLOCKED_IPS.check(mapped, "ipv4");
    }
  }

  return BLOCKED_IPS.check(host, version === 6 ? "ipv6" : "ipv4");
}

function checkAllowlist(host, allowlist) {
  if (allowlist.length === 0) return true;

  for (const pattern of allowlist) {
    if (matchesHostPattern(host, pattern)) {
      return true;
    }
  }
  return false;
}

const NORMALIZED_ALLOWED_HOSTS = MCP_ALLOWED_HOSTS
  .map(normalizeHostPattern)
  .filter(Boolean);

const NORMALIZED_TOKEN_FORWARD_ALLOWLIST = MCP_TOKEN_FORWARD_ALLOWLIST
  .map(normalizeHostPattern)
  .filter(Boolean);

export function validateMcpTargetUrl(serverUrl, options = {}) {
  const { requiresTokenForwarding = false } = options;

  if (!serverUrl || typeof serverUrl !== "string") {
    return { allowed: false, error: "Missing required field: serverUrl" };
  }

  let parsedUrl;
  try {
    parsedUrl = new URL(serverUrl);
  } catch {
    return { allowed: false, error: "Invalid serverUrl. Expected a valid absolute URL." };
  }

  if (parsedUrl.protocol !== "https:") {
    return { allowed: false, error: "MCP serverUrl must use HTTPS." };
  }

  const host = normalizeHost(parsedUrl.hostname);
  if (!host) {
    return { allowed: false, error: "Invalid serverUrl host." };
  }

  if (isBlockedHostname(host) || isBlockedLiteralIp(host)) {
    return {
      allowed: false,
      error: "Target host is blocked by MCP security policy (local/private/link-local hosts are not allowed).",
    };
  }

  if (!checkAllowlist(host, NORMALIZED_ALLOWED_HOSTS)) {
    return {
      allowed: false,
      error: "Target host is not allowed by MCP_ALLOWED_HOSTS.",
    };
  }

  if (requiresTokenForwarding) {
    if (NORMALIZED_TOKEN_FORWARD_ALLOWLIST.length === 0) {
      return {
        allowed: false,
        error: "Bearer token forwarding is disabled until MCP_TOKEN_FORWARD_ALLOWLIST is configured.",
      };
    }
    if (!checkAllowlist(host, NORMALIZED_TOKEN_FORWARD_ALLOWLIST)) {
      return {
        allowed: false,
        error: "Bearer token forwarding is only allowed to hosts in MCP_TOKEN_FORWARD_ALLOWLIST.",
      };
    }
  }

  return { allowed: true, host, parsedUrl };
}
