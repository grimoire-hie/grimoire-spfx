/**
 * Backend authentication — managed identity or API key.
 */

import { DefaultAzureCredential } from "@azure/identity";

const AZURE_OPENAI_SCOPE = "https://cognitiveservices.azure.com/.default";
let credential = null;
let tokenCache = { token: null, expiresAt: 0 };

async function getAzureToken() {
  const now = Date.now();

  // Return cached token if still valid (with 5 min buffer)
  if (tokenCache.token && tokenCache.expiresAt > now + 5 * 60 * 1000) {
    return tokenCache.token;
  }

  // Lazy-init credential
  if (!credential) {
    credential = new DefaultAzureCredential();
  }

  const tokenResponse = await credential.getToken(AZURE_OPENAI_SCOPE);
  tokenCache = {
    token: tokenResponse.token,
    expiresAt: tokenResponse.expiresOnTimestamp,
  };

  return tokenCache.token;
}

/**
 * Get auth headers for Azure OpenAI backend.
 * Uses API key if configured, managed identity otherwise.
 */
export async function getBackendAuthHeaders(backend, context) {
  const startedAt = Date.now();

  if (backend.key) {
    return {
      headers: { "api-key": backend.key },
      durationMs: Date.now() - startedAt,
      source: "api-key",
    };
  }

  try {
    const token = await getAzureToken();
    return {
      headers: { Authorization: `Bearer ${token}` },
      durationMs: Date.now() - startedAt,
      source: "managed-identity",
    };
  } catch (error) {
    context.error(`Managed identity token error: ${error.message}`);
    throw new Error(
      "No API key configured and managed identity token failed. " +
      "Set BACKEND_{name}_KEY or enable system-assigned managed identity."
    );
  }
}
