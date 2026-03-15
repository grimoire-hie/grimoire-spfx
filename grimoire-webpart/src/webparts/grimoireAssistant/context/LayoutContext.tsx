/**
 * LayoutContext
 * Manages layout state (sidebar, responsive breakpoints, available height).
 */

import * as React from 'react';
import {
  HostEnvironment,
  detectHostEnvironment,
  getHostChromeHeight,
  isMobileViewport,
  isTabletViewport
} from '../utils/hostEnvironment';

export interface ILayoutState {
  availableHeight: number;
  hostEnvironment: HostEnvironment;
  isMobile: boolean;
  isTablet: boolean;
}

export interface ILayoutActions {
  updateLayout: () => void;
}

const LayoutContext = React.createContext<(ILayoutState & ILayoutActions) | undefined>(undefined);

export interface ILayoutProviderProps {
  hasTeamsContext: boolean;
  children: React.ReactNode;
}

export const LayoutProvider: React.FC<ILayoutProviderProps> = ({
  hasTeamsContext,
  children
}) => {
  const [state, setState] = React.useState<ILayoutState>(() => {
    const env = detectHostEnvironment(hasTeamsContext);
    const chromeHeight = getHostChromeHeight(env);
    return {
      availableHeight: typeof window !== 'undefined' ? window.innerHeight - chromeHeight : 800,
      hostEnvironment: env,
      isMobile: isMobileViewport(),
      isTablet: isTabletViewport()
    };
  });

  const updateLayout = React.useCallback(() => {
    const env = detectHostEnvironment(hasTeamsContext);
    const chromeHeight = getHostChromeHeight(env);
    setState({
      availableHeight: window.innerHeight - chromeHeight,
      hostEnvironment: env,
      isMobile: isMobileViewport(),
      isTablet: isTabletViewport()
    });
  }, [hasTeamsContext]);

  React.useEffect(() => {
    const handleResize = (): void => updateLayout();
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, [updateLayout]);

  const value = React.useMemo(
    () => ({ ...state, updateLayout }),
    [state, updateLayout]
  );

  return (
    <LayoutContext.Provider value={value}>
      {children}
    </LayoutContext.Provider>
  );
};

export const useLayout = (): ILayoutState & ILayoutActions => {
  const context = React.useContext(LayoutContext);
  if (!context) {
    throw new Error('useLayout must be used within a LayoutProvider');
  }
  return context;
};
