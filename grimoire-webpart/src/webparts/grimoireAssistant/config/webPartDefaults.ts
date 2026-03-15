export const DEFAULT_PROXY_BACKEND = 'reasoning';
export const DEFAULT_DEPLOYMENT_PREFIX = 'grimoire';
export const GRIMOIRE_API_VERSION = '2025-01-01-preview';
export const PROXY_BACKEND_KEYS = ['reasoning', 'fast'] as const;
export type ProxyBackendKey = typeof PROXY_BACKEND_KEYS[number];

export const PROXY_BACKEND_ROUTE_MAP: Record<ProxyBackendKey, string> = {
  reasoning: 'reasoning',
  fast: 'fast'
};

/**
 * Build a deployment name from a prefix and backend key.
 * e.g., getDeploymentName('atlas-x7k2', 'reasoning') → 'atlas-x7k2-reasoning'
 */
export function getDeploymentName(prefix: string, backend: ProxyBackendKey): string {
  return (prefix || DEFAULT_DEPLOYMENT_PREFIX) + '-' + backend;
}

export interface IProxyBackendOption {
  key: ProxyBackendKey;
  text: string;
}

export function getProxyBackendOptions(labels: Record<ProxyBackendKey, string>): IProxyBackendOption[] {
  return PROXY_BACKEND_KEYS.map((key) => ({
    key,
    text: labels[key]
  }));
}
