/**
 * Host environment detection and responsive utilities
 */

export type HostEnvironment = 'sharepoint' | 'teams' | 'office' | 'outlook' | 'local';

export const BREAKPOINTS = {
  mobile: 768,
  tablet: 1024,
  desktop: 1440
} as const;

/**
 * Detect the host environment based on URL patterns and Teams context
 */
export function detectHostEnvironment(hasTeamsContext: boolean): HostEnvironment {
  if (hasTeamsContext) {
    return 'teams';
  }

  if (typeof window === 'undefined') {
    return 'local';
  }

  const url = window.location.href.toLowerCase();

  if (url.includes('.sharepoint.com') || url.includes('/_layouts/')) {
    return 'sharepoint';
  }
  if (url.includes('office.com') || url.includes('office365.com')) {
    return 'office';
  }
  if (url.includes('outlook.')) {
    return 'outlook';
  }

  return 'local';
}

/**
 * Get the estimated height of the host's chrome (header bars, etc.)
 */
export function getHostChromeHeight(environment: HostEnvironment): number {
  switch (environment) {
    case 'sharepoint':
    case 'office':
    case 'outlook':
      return 48;
    case 'teams':
    case 'local':
    default:
      return 0;
  }
}

export function isMobileViewport(): boolean {
  if (typeof window === 'undefined') return false;
  return window.innerWidth < BREAKPOINTS.mobile;
}

export function isTabletViewport(): boolean {
  if (typeof window === 'undefined') return false;
  return window.innerWidth >= BREAKPOINTS.mobile && window.innerWidth < BREAKPOINTS.tablet;
}

