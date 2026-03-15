/**
 * ErrorBoundary
 * React class component (React 17 requires class-based error boundaries).
 * Catches render crashes, logs to LogService, and renders a visible fallback
 * with the actual error message + stack trace instead of SPFx's "[object Object]".
 */

import * as React from 'react';
import * as strings from 'GrimoireAssistantWebPartStrings';
import { logService } from '../services/logging/LogService';

interface IErrorBoundaryState {
  hasError: boolean;
  error: Error | undefined;
}

export class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  IErrorBoundaryState
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, error: undefined };
  }

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    const detail = [
      error.message,
      error.stack || '(no stack)',
      info.componentStack || ''
    ].join('\n');

    logService.error('system', `Render crash: ${error.message}`, detail);
  }

  public render(): React.ReactNode {
    if (this.state.hasError) {
      const { error } = this.state;
      return (
        <div
          style={{
            padding: 24,
            backgroundColor: '#1a0000',
            color: '#ff6b6b',
            fontFamily: '"Cascadia Code", "Consolas", monospace',
            fontSize: 13,
            minHeight: 200,
            overflow: 'auto'
          }}
        >
          <h3 style={{ margin: '0 0 12px', color: '#ff6b6b' }}>
            {strings.RenderError}
          </h3>
          <p style={{ color: '#e0e0e0', margin: '0 0 12px' }}>
            {error?.message || String(error)}
          </p>
          {error?.stack && (
            <pre
              style={{
                fontSize: 11,
                color: '#888888',
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-all',
                margin: '0 0 16px',
                maxHeight: 300,
                overflow: 'auto'
              }}
            >
              {error.stack}
            </pre>
          )}
          <button
            onClick={() => this.setState({ hasError: false, error: undefined })}
            style={{
              padding: '6px 16px',
              backgroundColor: '#a8d8ff',
              color: '#0a0a1a',
              border: 'none',
              borderRadius: 4,
              cursor: 'pointer',
              fontWeight: 600
            }}
          >
            {strings.RetryButton}
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
