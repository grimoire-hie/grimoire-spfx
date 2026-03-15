export function createTimingHeaders(existingHeaders = {}, diagnostics = {}) {
  const headers = { ...existingHeaders };

  if (Number.isFinite(diagnostics.totalDurationMs)) {
    headers["X-Grimoire-Duration-Ms"] = String(Math.max(0, Math.round(diagnostics.totalDurationMs)));
  }
  if (Number.isFinite(diagnostics.authDurationMs)) {
    headers["X-Grimoire-Auth-Duration-Ms"] = String(Math.max(0, Math.round(diagnostics.authDurationMs)));
  }
  if (Number.isFinite(diagnostics.upstreamDurationMs)) {
    headers["X-Grimoire-Upstream-Duration-Ms"] = String(Math.max(0, Math.round(diagnostics.upstreamDurationMs)));
  }
  if (Number.isFinite(diagnostics.storageInitDurationMs) && diagnostics.storageInitDurationMs > 0) {
    headers["X-Grimoire-Storage-Init-Duration-Ms"] = String(Math.max(0, Math.round(diagnostics.storageInitDurationMs)));
  }

  return headers;
}
