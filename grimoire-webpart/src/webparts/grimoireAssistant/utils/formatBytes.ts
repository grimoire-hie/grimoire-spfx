/**
 * Format a byte count into a human-readable string (B, KB, MB).
 * Returns fallback string when bytes is undefined/0.
 */
export function formatBytes(bytes: number | undefined, fallback: string = ''): string {
  if (!bytes) return fallback;
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
