/**
 * Rate limiting — in-memory, resets on function restart.
 *
 * SCALING NOTE: This implementation is per-instance. In a multi-instance
 * Azure Functions deployment, each instance has its own rate limit store.
 * For strict global rate limiting, replace with Azure Cache for Redis.
 * For the current single-instance consumption plan, this is sufficient.
 */

import { REQUESTS_PER_MINUTE, REQUESTS_PER_DAY } from "../utils/config.js";

const rateLimitStore = new Map();

export function checkRateLimit(key) {
  const now = Date.now();
  const minuteAgo = now - 60 * 1000;
  const dayAgo = now - 24 * 60 * 60 * 1000;

  if (!rateLimitStore.has(key)) {
    rateLimitStore.set(key, []);
  }

  const requests = rateLimitStore.get(key);

  // Clean old entries
  const validRequests = requests.filter((ts) => ts > dayAgo);
  rateLimitStore.set(key, validRequests);

  // Count recent requests
  const minuteCount = validRequests.filter((ts) => ts > minuteAgo).length;
  const dayCount = validRequests.length;

  if (minuteCount >= REQUESTS_PER_MINUTE) {
    return {
      allowed: false,
      error: `Rate limit exceeded. Max ${REQUESTS_PER_MINUTE} requests per minute.`,
    };
  }

  if (dayCount >= REQUESTS_PER_DAY) {
    return {
      allowed: false,
      error: `Daily limit exceeded. Max ${REQUESTS_PER_DAY} requests per day.`,
    };
  }

  // Record this request
  validRequests.push(now);

  return { allowed: true, minuteCount: minuteCount + 1, dayCount: dayCount + 1 };
}

/**
 * Check rate limit and return error response if exceeded, or null if ok.
 */
export function enforceRateLimit(apiKey, corsHeaders) {
  const rateCheck = checkRateLimit(apiKey);
  if (!rateCheck.allowed) {
    return {
      status: 429,
      headers: corsHeaders,
      jsonBody: { error: rateCheck.error },
    };
  }
  return null;
}
