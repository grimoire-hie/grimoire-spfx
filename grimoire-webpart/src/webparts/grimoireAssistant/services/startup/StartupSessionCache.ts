import type { PublicWebSearchCapabilityStatus, IProxyConfig } from '../../store/useGrimoireStore';

const SESSION_KEY_PREFIX = 'grimoire.startup.v1';

export const BACKEND_HEALTH_CACHE_TTL_MS = 30_000;
export const PUBLIC_WEB_PROBE_CACHE_TTL_MS = 60_000;

export interface ISessionPreferenceSeed {
  avatarRaw?: string;
  assistantRaw?: string;
}

export interface IBackendHealthCacheEntry {
  backendOk: boolean;
  checkedAt: number;
}

export interface IPublicWebProbeCacheEntry {
  status: PublicWebSearchCapabilityStatus;
  detail?: string;
  checkedAt: number;
}

interface IScopedCacheEntry<TPayload> {
  scope: string;
  storedAt: number;
  payload: TPayload;
}

function getSessionStorage(): Storage | undefined {
  if (typeof window === 'undefined') {
    return undefined;
  }

  try {
    return window.sessionStorage;
  } catch {
    return undefined;
  }
}

function readScopedEntry<TPayload>(key: string, scope: string): IScopedCacheEntry<TPayload> | undefined {
  const storage = getSessionStorage();
  if (!storage) {
    return undefined;
  }

  const raw = storage.getItem(key);
  if (!raw) {
    return undefined;
  }

  try {
    const parsed = JSON.parse(raw) as IScopedCacheEntry<TPayload>;
    if (!parsed || parsed.scope !== scope) {
      return undefined;
    }
    return parsed;
  } catch {
    return undefined;
  }
}

function writeScopedEntry<TPayload>(key: string, scope: string, payload: TPayload): void {
  const storage = getSessionStorage();
  if (!storage) {
    return;
  }

  const entry: IScopedCacheEntry<TPayload> = {
    scope,
    storedAt: Date.now(),
    payload
  };
  storage.setItem(key, JSON.stringify(entry));
}

function normalizeScopePart(value: string | undefined): string {
  return encodeURIComponent((value || '').trim().toLowerCase() || 'default');
}

export function resolveUserSessionScope(userIdentity: string | undefined): string {
  return `user:${normalizeScopePart(userIdentity)}`;
}

export function resolveProxySessionScope(proxyConfig: IProxyConfig | undefined): string {
  return `proxy:${normalizeScopePart(
    proxyConfig
      ? `${proxyConfig.proxyUrl}|${proxyConfig.backend}|${proxyConfig.backendApiResource || ''}`
      : 'none'
  )}`;
}

export function readSessionPreferenceSeed(scope: string): ISessionPreferenceSeed | undefined {
  return readScopedEntry<ISessionPreferenceSeed>(`${SESSION_KEY_PREFIX}.preferences`, scope)?.payload;
}

export function writeSessionPreferenceSeed(
  scope: string,
  avatarRaw: string,
  assistantRaw: string
): void {
  writeScopedEntry(`${SESSION_KEY_PREFIX}.preferences`, scope, {
    avatarRaw,
    assistantRaw
  });
}

export function readCachedBackendHealth(
  scope: string,
  ttlMs: number = BACKEND_HEALTH_CACHE_TTL_MS
): IBackendHealthCacheEntry | undefined {
  const entry = readScopedEntry<IBackendHealthCacheEntry>(`${SESSION_KEY_PREFIX}.health`, scope);
  if (!entry) {
    return undefined;
  }
  if ((Date.now() - entry.storedAt) > ttlMs) {
    return undefined;
  }
  return entry.payload;
}

export function writeCachedBackendHealth(scope: string, payload: IBackendHealthCacheEntry): void {
  writeScopedEntry(`${SESSION_KEY_PREFIX}.health`, scope, payload);
}

export function readCachedPublicWebProbe(
  scope: string,
  ttlMs: number = PUBLIC_WEB_PROBE_CACHE_TTL_MS
): IPublicWebProbeCacheEntry | undefined {
  const entry = readScopedEntry<IPublicWebProbeCacheEntry>(`${SESSION_KEY_PREFIX}.public-web`, scope);
  if (!entry) {
    return undefined;
  }
  if ((Date.now() - entry.storedAt) > ttlMs) {
    return undefined;
  }
  return entry.payload;
}

export function writeCachedPublicWebProbe(scope: string, payload: IPublicWebProbeCacheEntry): void {
  writeScopedEntry(`${SESSION_KEY_PREFIX}.public-web`, scope, payload);
}
