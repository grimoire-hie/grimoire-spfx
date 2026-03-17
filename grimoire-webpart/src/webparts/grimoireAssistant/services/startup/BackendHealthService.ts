import type { IProxyConfig } from '../../store/useGrimoireStore';
import {
  BACKEND_HEALTH_CACHE_TTL_MS,
  readCachedBackendHealth,
  resolveProxySessionScope,
  writeCachedBackendHealth
} from './StartupSessionCache';

export interface IBackendHealthResult {
  backendOk: boolean;
  checkedAt: number;
  source: 'network' | 'session-cache';
  durationMs: number;
}

export async function fetchBackendHealth(
  proxyConfig: IProxyConfig,
  options?: {
    allowSessionCache?: boolean;
    ttlMs?: number;
  }
): Promise<IBackendHealthResult> {
  const ttlMs = options?.ttlMs ?? BACKEND_HEALTH_CACHE_TTL_MS;
  const cacheScope = resolveProxySessionScope(proxyConfig);

  if (options?.allowSessionCache !== false) {
    const cached = readCachedBackendHealth(cacheScope, ttlMs);
    if (cached) {
      return {
        backendOk: cached.backendOk,
        checkedAt: cached.checkedAt,
        source: 'session-cache',
        durationMs: 0
      };
    }
  }

  const startedAt = typeof performance !== 'undefined' ? performance.now() : Date.now();
  const response = await fetch(`${proxyConfig.proxyUrl}/health`, {
    headers: { 'x-functions-key': proxyConfig.proxyApiKey }
  });
  const durationMs = Math.max(
    0,
    Math.round((typeof performance !== 'undefined' ? performance.now() : Date.now()) - startedAt)
  );
  const checkedAt = Date.now();
  const result: IBackendHealthResult = {
    backendOk: response.ok,
    checkedAt,
    source: 'network',
    durationMs
  };

  writeCachedBackendHealth(cacheScope, {
    backendOk: result.backendOk,
    checkedAt
  });

  return result;
}
