/**
 * Realtime Token handler.
 * Issues ephemeral tokens for WebRTC GPT Realtime sessions.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { validateAuth } from "../middleware/auth.js";
import { enforceRateLimit } from "../middleware/rateLimit.js";
import { getBackendAuthHeaders } from "../llm/backendAuth.js";
import { BACKENDS, REALTIME_TOKEN_TIMEOUT_MS, REALTIME_DEPLOYMENT_NAME } from "../utils/config.js";
import { createTimingHeaders } from "../utils/diagnostics.js";

const SUPPORTED_REALTIME_VOICES = new Set(["alloy", "echo", "shimmer", "marin", "cedar"]);
const DEFAULT_REALTIME_VOICE = "alloy";

function normalizeRealtimeVoice(voice) {
  return typeof voice === "string" && SUPPORTED_REALTIME_VOICES.has(voice)
    ? voice
    : DEFAULT_REALTIME_VOICE;
}

export async function realtimeTokenHandler(request, context) {
  const requestStartedAt = Date.now();
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  // Auth + rate limit
  const auth = validateAuth(request, corsHeaders);
  if (!auth.authenticated) return auth.errorResponse;
  const rateLimitError = enforceRateLimit(auth.apiKey, corsHeaders);
  if (rateLimitError) return rateLimitError;

  // Parse body
  let body;
  try {
    body = await request.json();
  } catch {
    return { status: 400, headers: corsHeaders, jsonBody: { error: "Invalid JSON body." } };
  }

  const { voice, instructions } = body;
  const normalizedVoice = normalizeRealtimeVoice(voice);

  // Use reasoning backend to derive the Foundry endpoint
  const backend = BACKENDS["reasoning"];
  if (!backend) {
    return {
      status: 500,
      headers: corsHeaders,
      jsonBody: { error: "No reasoning backend configured. Cannot derive Foundry endpoint." },
    };
  }

  const foundryEndpoint = backend.endpoint.replace(/\/$/, "");
  const tokenUrl = `${foundryEndpoint}/openai/v1/realtime/client_secrets`;

  context.log(`[realtime] Requesting ephemeral token from: ${foundryEndpoint}`);

  try {
    const auth = await getBackendAuthHeaders(backend, context);
    const authHeaders = auth.headers;

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), REALTIME_TOKEN_TIMEOUT_MS);

    let tokenResponse;
    const upstreamStartedAt = Date.now();
    try {
      tokenResponse = await fetch(tokenUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...authHeaders },
        body: JSON.stringify({
          session: {
            type: "realtime",
            model: REALTIME_DEPLOYMENT_NAME,
            instructions: instructions || "You are a helpful M365 assistant. Always start the conversation in English.",
            audio: { output: { voice: normalizedVoice } },
          },
        }),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timeout);
    }
    const upstreamDurationMs = Date.now() - upstreamStartedAt;

    const diagnosticHeaders = createTimingHeaders(corsHeaders, {
      totalDurationMs: Date.now() - requestStartedAt,
      authDurationMs: auth.durationMs,
      upstreamDurationMs,
    });

    if (!tokenResponse.ok) {
      const errorText = await tokenResponse.text();
      context.warn(`[realtime] Token request failed: ${tokenResponse.status}`);
      return {
        status: 502,
        headers: diagnosticHeaders,
        jsonBody: {
          error: `Realtime API returned ${tokenResponse.status}. The realtime model may not be deployed.`,
          detail: errorText.slice(0, 500),
        },
      };
    }

    const tokenData = await tokenResponse.json();

    return {
      headers: diagnosticHeaders,
      jsonBody: {
        clientSecret: tokenData.client_secret?.value || tokenData.value,
        expiresAt: tokenData.client_secret?.expires_at || tokenData.expires_at,
        endpoint: foundryEndpoint,
        voice: normalizedVoice,
      },
    };
  } catch (error) {
    if (error.name === "AbortError") {
      context.error(`[realtime] Token request timed out after ${REALTIME_TOKEN_TIMEOUT_MS}ms`);
      return {
        status: 504,
        headers: createTimingHeaders(corsHeaders, {
          totalDurationMs: Date.now() - requestStartedAt,
        }),
        jsonBody: { error: `Realtime token request timed out after ${REALTIME_TOKEN_TIMEOUT_MS}ms` },
      };
    }
    context.error(`[realtime] Error: ${error.message}`);
    return {
      status: 500,
      headers: createTimingHeaders(corsHeaders, {
        totalDurationMs: Date.now() - requestStartedAt,
      }),
      jsonBody: { error: "Failed to obtain realtime token: " + error.message },
    };
  }
}
