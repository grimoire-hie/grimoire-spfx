/**
 * Easy Auth user identity resolution for /api/user/* routes.
 */

const OBJECT_ID_CLAIM_TYPES = [
  "http://schemas.microsoft.com/identity/claims/objectidentifier",
  "oid",
];

const USERNAME_CLAIM_TYPES = [
  "preferred_username",
  "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn",
  "upn",
  "email",
  "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress",
];

function readHeader(headers, name) {
  if (!headers) return undefined;
  if (typeof headers.get === "function") {
    return headers.get(name) ?? headers.get(name.toLowerCase()) ?? undefined;
  }
  const direct = headers[name] ?? headers[name.toLowerCase()];
  return typeof direct === "string" ? direct : undefined;
}

function normalizePartitionIdentity(value) {
  return String(value || "").trim().toLowerCase();
}

function decodeClientPrincipal(encodedPrincipal) {
  if (!encodedPrincipal) return undefined;

  const normalized = encodedPrincipal.replace(/-/g, "+").replace(/_/g, "/");
  const padding = normalized.length % 4 === 0
    ? ""
    : "=".repeat(4 - (normalized.length % 4));

  try {
    const decodedJson = Buffer.from(`${normalized}${padding}`, "base64").toString("utf8");
    return JSON.parse(decodedJson);
  } catch {
    return undefined;
  }
}

function buildClaimMap(principal) {
  const claimMap = new Map();
  const claims = Array.isArray(principal?.claims) ? principal.claims : [];

  for (const claim of claims) {
    const claimType = String(claim?.typ ?? claim?.type ?? "").trim().toLowerCase();
    const claimValue = String(claim?.val ?? claim?.value ?? "").trim();
    if (claimType && claimValue && !claimMap.has(claimType)) {
      claimMap.set(claimType, claimValue);
    }
  }

  return claimMap;
}

function getFirstClaimValue(claimMap, claimTypes) {
  for (const claimType of claimTypes) {
    const value = claimMap.get(claimType.toLowerCase());
    if (value) {
      return value;
    }
  }
  return undefined;
}

export function extractAuthenticatedUser(request) {
  const principalHeader = readHeader(request?.headers, "x-ms-client-principal");
  const principal = decodeClientPrincipal(principalHeader);
  const claimMap = buildClaimMap(principal);

  const objectId = normalizePartitionIdentity(
    getFirstClaimValue(claimMap, OBJECT_ID_CLAIM_TYPES)
      || readHeader(request?.headers, "x-ms-client-principal-id")
  );

  if (!objectId) {
    return undefined;
  }

  const upnOrEmail = String(
    getFirstClaimValue(claimMap, USERNAME_CLAIM_TYPES)
      || principal?.userDetails
      || readHeader(request?.headers, "x-ms-client-principal-name")
      || ""
  ).trim();

  return {
    objectId,
    upnOrEmail,
  };
}

export function requireAuthenticatedUser(request, corsHeaders) {
  const user = extractAuthenticatedUser(request);
  if (!user) {
    return {
      authenticated: false,
      errorResponse: {
        status: 401,
        headers: corsHeaders,
        jsonBody: { error: "Authenticated user identity required." },
      },
    };
  }

  return { authenticated: true, user };
}
