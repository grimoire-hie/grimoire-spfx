jest.mock('../hie/HybridInteractionEngine', () => ({
  hybridInteractionEngine: {
    emitEvent: jest.fn(),
    getTurnLineage: jest.fn()
  }
}));

import {
  createRateLimitLifecycleCallbacks,
  formatRateLimitRetryStatus,
  RATE_LIMIT_EXHAUSTED_STATUS
} from './TextChatRateLimitUi';
import { hybridInteractionEngine } from '../hie/HybridInteractionEngine';

describe('TextChatRateLimitUi', () => {
  beforeEach(() => {
    jest.useFakeTimers();
    (hybridInteractionEngine.getTurnLineage as jest.Mock).mockReturnValue({
      turnId: 'turn-1',
      rootTurnId: 'root-1',
      parentTurnId: 'parent-1'
    });
  });

  afterEach(() => {
    jest.useRealTimers();
    jest.clearAllMocks();
  });

  it('formats retry status using rounded-up seconds', () => {
    expect(formatRateLimitRetryStatus({
      attempt: 2,
      maxRetries: 3,
      delayMs: 1501,
      headerSource: 'retry-after-ms',
      status: 'retrying'
    })).toBe('Rate-limited, retrying in 2s... (2/3)');
  });

  it('emits store-only HIE events and clears exhausted status after the timeout', () => {
    let activityStatus = '';
    const setActivityStatus = jest.fn((status: string) => {
      activityStatus = status;
    });
    const callbacks = createRateLimitLifecycleCallbacks({
      resolveTurnId: () => 'turn-1',
      setActivityStatus,
      getActivityStatus: () => activityStatus,
      clearDelayMs: 8000
    });

    callbacks.onRateLimitRetry?.({
      attempt: 1,
      maxRetries: 3,
      delayMs: 5000,
      headerSource: 'fallback-exponential',
      status: 'retrying'
    });

    expect(setActivityStatus).toHaveBeenLastCalledWith('Rate-limited, retrying in 5s... (1/3)');
    expect(hybridInteractionEngine.emitEvent).toHaveBeenNthCalledWith(1, expect.objectContaining({
      eventName: 'llm.rate_limit.retry_scheduled',
      turnId: 'turn-1',
      rootTurnId: 'root-1',
      parentTurnId: 'parent-1',
      exposurePolicy: { mode: 'store-only', relevance: 'background' },
      payload: expect.objectContaining({
        attempt: 1,
        delayMs: 5000,
        headerSource: 'fallback-exponential',
        status: 'retrying'
      })
    }));

    callbacks.onRateLimitResolved?.({
      attempt: 1,
      maxRetries: 3,
      delayMs: 5000,
      headerSource: 'fallback-exponential',
      status: 'resolved'
    });

    expect(setActivityStatus).toHaveBeenLastCalledWith('');
    expect(hybridInteractionEngine.emitEvent).toHaveBeenNthCalledWith(2, expect.objectContaining({
      eventName: 'llm.rate_limit.recovered',
      payload: expect.objectContaining({
        attempt: 1,
        status: 'resolved'
      })
    }));

    callbacks.onRateLimitExhausted?.({
      attempt: 3,
      maxRetries: 3,
      delayMs: 30000,
      headerSource: 'fallback-exponential',
      status: 'exhausted'
    });

    expect(callbacks.isExhausted()).toBe(true);
    expect(setActivityStatus).toHaveBeenLastCalledWith(RATE_LIMIT_EXHAUSTED_STATUS);
    expect(hybridInteractionEngine.emitEvent).toHaveBeenNthCalledWith(3, expect.objectContaining({
      eventName: 'llm.rate_limit.exhausted',
      payload: expect.objectContaining({
        attempt: 3,
        delayMs: 30000,
        status: 'exhausted'
      })
    }));

    jest.advanceTimersByTime(7999);
    expect(setActivityStatus).toHaveBeenCalledTimes(3);

    jest.advanceTimersByTime(1);
    expect(setActivityStatus).toHaveBeenLastCalledWith('');
    expect(callbacks.isExhausted()).toBe(false);
  });
});
