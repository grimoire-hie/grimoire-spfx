import { logService } from '../logging/LogService';
import {
  type IStartupMetric,
  type StartupMetricName,
  type StartupPhase,
  useGrimoireStore
} from '../../store/useGrimoireStore';

const activeMetricStarts = new Map<StartupMetricName, number>();

function getNowMs(): number {
  if (typeof performance !== 'undefined' && typeof performance.now === 'function') {
    return performance.now();
  }
  return Date.now();
}

function buildCompletedMetric(
  name: StartupMetricName,
  endMs: number,
  detail?: string
): IStartupMetric {
  const current = useGrimoireStore.getState().startupMetrics[name];
  const startedAtMs = activeMetricStarts.get(name) ?? current?.startedAtMs ?? endMs;
  const durationMs = Math.max(0, Math.round(endMs - startedAtMs));

  return {
    status: 'completed',
    startedAtMs,
    durationMs,
    completedAtMs: endMs,
    detail
  };
}

export function setStartupPhase(phase: StartupPhase, detail?: string): void {
  const store = useGrimoireStore.getState();
  if (store.startupPhase === phase) {
    return;
  }

  store.setStartupPhase(phase);
  logService.info('system', `Startup phase: ${phase}`, detail);
}

export function beginStartupMetric(name: StartupMetricName, detail?: string): void {
  const nowMs = getNowMs();
  activeMetricStarts.set(name, nowMs);
  useGrimoireStore.getState().updateStartupMetric(name, {
    status: 'running',
    startedAtMs: nowMs,
    detail
  });
}

export function completeStartupMetric(name: StartupMetricName, detail?: string): void {
  const store = useGrimoireStore.getState();
  if (store.startupMetrics[name]?.status === 'completed') {
    return;
  }

  const completedMetric = buildCompletedMetric(name, getNowMs(), detail);
  activeMetricStarts.delete(name);
  store.updateStartupMetric(name, completedMetric);
  logService.info(
    'system',
    `Startup metric: ${name}`,
    detail,
    completedMetric.durationMs
  );
}

export function recordStartupMetric(name: StartupMetricName, detail?: string): void {
  const store = useGrimoireStore.getState();
  if (store.startupMetrics[name]?.status === 'completed') {
    return;
  }

  const nowMs = getNowMs();
  activeMetricStarts.delete(name);
  store.updateStartupMetric(name, {
    status: 'completed',
    startedAtMs: nowMs,
    durationMs: 0,
    completedAtMs: nowMs,
    detail
  });
  logService.info('system', `Startup metric: ${name}`, detail, 0);
}

export function resetStartupInstrumentation(): void {
  activeMetricStarts.clear();
}
