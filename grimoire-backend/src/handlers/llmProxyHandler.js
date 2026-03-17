/**
 * LLM Proxy handler.
 * Routes /api/{backend}/openai/* to Azure OpenAI with rate limiting.
 * Supports streaming (SSE) and regular JSON responses.
 */

import { getCorsHeaders, handlePreflight } from "../middleware/cors.js";
import { resolveCallerId } from "../middleware/callerIdentity.js";
import { checkRateLimit } from "../middleware/rateLimit.js";
import { getBackendAuthHeaders } from "../llm/backendAuth.js";
import {
  BACKENDS,
  REQUESTS_PER_MINUTE,
  REQUESTS_PER_DAY,
  LLM_UPSTREAM_TIMEOUT_MS,
} from "../utils/config.js";
import { createTimingHeaders } from "../utils/diagnostics.js";

export async function llmProxyHandler(request, context) {
  const requestStartedAt = Date.now();
  const preflight = handlePreflight(request);
  if (preflight) return preflight;

  const corsHeaders = getCorsHeaders(request);

  // Extract backend name from path
  const url = new URL(request.url);
  const pathMatch = url.pathname.match(/\/api\/([^/]+)\/(openai\/.*)$/);

  if (!pathMatch) {
    return {
      status: 400,
      headers: corsHeaders,
      jsonBody: {
        error: "Invalid path format.",
        usage: "/api/{backend}/openai/deployments/{deployment}/chat/completions",
        availableBackends: Object.keys(BACKENDS),
      },
    };
  }

  const backendName = pathMatch[1];
  const openaiPath = "/" + pathMatch[2] + url.search;

  // Validate backend exists
  const backend = BACKENDS[backendName];
  if (!backend) {
    return {
      status: 404,
      headers: corsHeaders,
      jsonBody: {
        error: `Backend '${backendName}' not found.`,
        availableBackends: Object.keys(BACKENDS),
      },
    };
  }

  // Rate limit (keyed on caller identity)
  const callerId = resolveCallerId(request);
  const rateCheck = checkRateLimit(callerId);
  if (!rateCheck.allowed) {
    return { status: 429, headers: corsHeaders, jsonBody: { error: rateCheck.error } };
  }

  // Build target URL
  const targetUrl = backend.endpoint + openaiPath;
  const authMethod = backend.key ? "api-key" : "managed-identity";

  context.log(`[${backendName}] Proxying: ${openaiPath} (auth: ${authMethod})`);

  try {
    const auth = await getBackendAuthHeaders(backend, context);
    const authHeaders = auth.headers;
    const body = request.method !== "GET" ? await request.text() : undefined;

    // Check if streaming
    let isStreaming = false;
    try {
      isStreaming = body && JSON.parse(body).stream === true;
    } catch { /* not JSON */ }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), LLM_UPSTREAM_TIMEOUT_MS);

    let response;
    const upstreamStartedAt = Date.now();
    try {
      response = await fetch(targetUrl, {
        method: request.method,
        headers: { "Content-Type": "application/json", ...authHeaders },
        body,
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timeout);
    }
    const upstreamDurationMs = Date.now() - upstreamStartedAt;

    const rateLimitHeaders = createTimingHeaders({
      ...corsHeaders,
      "X-Backend": backendName,
      "X-Auth-Method": authMethod,
      "X-RateLimit-Remaining-Minute": String(REQUESTS_PER_MINUTE - rateCheck.minuteCount),
      "X-RateLimit-Remaining-Day": String(REQUESTS_PER_DAY - rateCheck.dayCount),
    }, {
      totalDurationMs: Date.now() - requestStartedAt,
      authDurationMs: auth.durationMs,
      upstreamDurationMs,
    });

    // Streaming response
    if (isStreaming) {
      return {
        status: response.status,
        body: response.body,
        headers: {
          ...rateLimitHeaders,
          "Content-Type": "text/event-stream",
          "Cache-Control": "no-cache",
          Connection: "keep-alive",
        },
      };
    }

    // Regular response
    const responseContentType = response.headers.get("content-type") || "";
    const responseText = await response.text();

    if (responseContentType.includes("application/json")) {
      try {
        const data = JSON.parse(responseText);
        if (data.usage?.total_tokens) {
          context.log(`[${backendName}] Tokens used: ${data.usage.total_tokens}`);
        }
        return { status: response.status, jsonBody: data, headers: rateLimitHeaders };
      } catch { /* invalid JSON */ }
    }

    return {
      status: response.status,
      body: responseText,
      headers: {
        ...rateLimitHeaders,
        ...(responseContentType ? { "Content-Type": responseContentType } : {}),
      },
    };
  } catch (error) {
    if (error.name === "AbortError") {
      context.error(`[${backendName}] Proxy timeout after ${LLM_UPSTREAM_TIMEOUT_MS}ms`);
      return {
        status: 504,
        headers: createTimingHeaders(corsHeaders, {
          totalDurationMs: Date.now() - requestStartedAt,
        }),
        jsonBody: { error: `Upstream LLM request timed out after ${LLM_UPSTREAM_TIMEOUT_MS}ms.` },
      };
    }
    context.error(`[${backendName}] Proxy error: ${error.message}`);
    return {
      status: 500,
      headers: createTimingHeaders(corsHeaders, {
        totalDurationMs: Date.now() - requestStartedAt,
      }),
      jsonBody: { error: "Proxy error: " + error.message },
    };
  }
}
