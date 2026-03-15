/**
 * URL helpers for SharePoint document links.
 *
 * SharePoint's `webUrl` for files is the raw file path which triggers
 * a download in the browser. Appending `?web=1` tells SharePoint to
 * open Office documents in the web viewer (Word Online, Excel Online, etc.)
 * instead of downloading them.
 */

const OFFICE_EXTENSIONS = new Set([
  'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt',
  'dotx', 'dotm', 'xlsm', 'xltx', 'pptm', 'potx',
  'one', 'vsdx'
]);

/**
 * Transform a SharePoint file URL into a web viewer URL.
 * For Office documents, appends `?web=1` so they open in Word Online / Excel Online etc.
 * Non-Office files (PDF, images, etc.) are returned as-is — browsers handle them natively.
 */
export function toWebViewerUrl(url: string): string {
  if (!url) return url;

  try {
    const parsed = new URL(url);
    const ext = parsed.pathname.split('.').pop()?.toLowerCase() || '';

    if (OFFICE_EXTENSIONS.has(ext)) {
      parsed.searchParams.set('web', '1');
      return parsed.toString();
    }
  } catch {
    // Not a valid URL — return as-is
  }

  return url;
}
