import type { ITextChatCallbacks, ITextChatRateLimitInfo } from './TextChatService';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';
import { createCorrelationId } from '../hie/HAEContracts';

export const RATE_LIMIT_EXHAUSTED_STATUS = 'Rate-limited. Automatic retries exhausted. Please try again.';
export const RATE_LIMIT_STATUS_CLEAR_MS = 8000;

export function formatRateLimitRetryStatus(info: ITextChatRateLimitInfo): string {
  const delaySeconds = Math.max(1, Math.ceil(info.delayMs / 1000));
  return `Rate-limited, retrying in ${delaySeconds}s... (${info.attempt}/${info.maxRetries})`;
}

interface IRateLimitLifecycleCallbacks extends Pick<ITextChatCallbacks, 'onRateLimitRetry' | 'onRateLimitResolved' | 'onRateLimitExhausted'> {
  clear: () => void;
  isExhausted: () => boolean;
}

interface ICreateRateLimitLifecycleCallbacksOptions {
  resolveTurnId?: () => string | undefined;
  clearDelayMs?: number;
  setActivityStatus?: (status: string) => void;
  getActivityStatus?: () => string;
  setTimeoutFn?: typeof setTimeout;
  clearTimeoutFn?: typeof clearTimeout;
}

export function createRateLimitLifecycleCallbacks(
  options?: ICreateRateLimitLifecycleCallbacksOptions
): IRateLimitLifecycleCallbacks {
  const setActivityStatus = options?.setActivityStatus || ((status: string) => {
    useGrimoireStore.getState().setActivityStatus(status);
  });
  const getActivityStatus = options?.getActivityStatus || (() => useGrimoireStore.getState().activityStatus);
  const setTimeoutFn = options?.setTimeoutFn || setTimeout;
  const clearTimeoutFn = options?.clearTimeoutFn || clearTimeout;
  const clearDelayMs = options?.clearDelayMs ?? RATE_LIMIT_STATUS_CLEAR_MS;
  let exhaustionTimer: ReturnType<typeof setTimeout> | undefined;
  let exhausted = false;

  const emitRateLimitEvent = (
    eventName: 'llm.rate_limit.retry_scheduled' | 'llm.rate_limit.recovered' | 'llm.rate_limit.exhausted',
    info: ITextChatRateLimitInfo
  ): void => {
    const turnId = options?.resolveTurnId?.();
    const lineage = hybridInteractionEngine.getTurnLineage(turnId);
    hybridInteractionEngine.emitEvent({
      eventName,
      source: 'system',
      surface: 'app-shell',
      correlationId: createCorrelationId('rl'),
      turnId,
      rootTurnId: lineage?.rootTurnId || turnId,
      parentTurnId: lineage?.parentTurnId,
      payload: {
        attempt: info.attempt,
        maxRetries: info.maxRetries,
        delayMs: info.delayMs,
        headerSource: info.headerSource,
        status: info.status
      },
      exposurePolicy: { mode: 'store-only', relevance: 'background' }
    });
  };

  const clearExhaustionTimer = (): void => {
    if (exhaustionTimer) {
      clearTimeoutFn(exhaustionTimer);
      exhaustionTimer = undefined;
    }
  };

  return {
    onRateLimitRetry: (info: ITextChatRateLimitInfo): void => {
      exhausted = false;
      clearExhaustionTimer();
      setActivityStatus(formatRateLimitRetryStatus(info));
      emitRateLimitEvent('llm.rate_limit.retry_scheduled', info);
    },
    onRateLimitResolved: (info: ITextChatRateLimitInfo): void => {
      exhausted = false;
      clearExhaustionTimer();
      setActivityStatus('');
      emitRateLimitEvent('llm.rate_limit.recovered', info);
    },
    onRateLimitExhausted: (info: ITextChatRateLimitInfo): void => {
      exhausted = true;
      clearExhaustionTimer();
      setActivityStatus(RATE_LIMIT_EXHAUSTED_STATUS);
      emitRateLimitEvent('llm.rate_limit.exhausted', info);
      exhaustionTimer = setTimeoutFn(() => {
        exhaustionTimer = undefined;
        exhausted = false;
        if (getActivityStatus() === RATE_LIMIT_EXHAUSTED_STATUS) {
          setActivityStatus('');
        }
      }, clearDelayMs);
    },
    clear: (): void => {
      exhausted = false;
      clearExhaustionTimer();
      setActivityStatus('');
    },
    isExhausted: (): boolean => exhausted
  };
}
