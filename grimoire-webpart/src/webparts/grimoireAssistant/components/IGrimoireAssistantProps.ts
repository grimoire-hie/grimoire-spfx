/**
 * Props for the GrimoireAssistant root component.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  DEFAULT_PROXY_BACKEND,
  DEFAULT_DEPLOYMENT_PREFIX,
  GRIMOIRE_API_VERSION,
  type ProxyBackendKey,
  getDeploymentName,
  PROXY_BACKEND_ROUTE_MAP
} from '../config/webPartDefaults';
import type { IProxyConfig, ISpThemeColors } from '../store/useGrimoireStore';

export type { IProxyConfig };

export interface IGrimoireAssistantProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  context: WebPartContext;
  proxyUrl?: string;
  proxyApiKey?: string;
  backendApiResource?: string;
  proxyBackend?: string;
  deploymentPrefix?: string;
  mcpEnvironmentId?: string;
  spThemeColors?: ISpThemeColors;
}

/**
 * Check if proxy is configured with required fields
 */
export function hasValidProxyConfig(props: IGrimoireAssistantProps): boolean {
  return !!(
    props.proxyUrl?.trim() &&
    props.proxyApiKey?.trim()
  );
}

/**
 * Build proxy config object from props
 */
export function getProxyConfig(props: IGrimoireAssistantProps): IProxyConfig | undefined {
  if (!hasValidProxyConfig(props)) {
    return undefined;
  }

  const backend = (props.proxyBackend as ProxyBackendKey | undefined) || DEFAULT_PROXY_BACKEND;
  const prefix = props.deploymentPrefix?.trim() || DEFAULT_DEPLOYMENT_PREFIX;

  return {
    proxyUrl: props.proxyUrl!,
    proxyApiKey: props.proxyApiKey!,
    backendApiResource: props.backendApiResource?.trim() || undefined,
    backend: PROXY_BACKEND_ROUTE_MAP[backend] || DEFAULT_PROXY_BACKEND,
    deployment: getDeploymentName(prefix, backend),
    apiVersion: GRIMOIRE_API_VERSION
  };
}
