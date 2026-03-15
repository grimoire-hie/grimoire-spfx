function safeParseUrl(url: string | undefined): URL | undefined {
  if (!url) {
    return undefined;
  }
  try {
    return new URL(url);
  } catch {
    return undefined;
  }
}

function getSitePathAndRemainder(parsed: URL): { sitePath: string; remainder: string[] } {
  const segments = parsed.pathname.split('/').filter(Boolean);
  if (segments.length >= 2 && /^(sites|teams|personal)$/i.test(segments[0])) {
    return {
      sitePath: `/${segments[0]}/${segments[1]}`,
      remainder: segments.slice(2)
    };
  }

  return {
    sitePath: '',
    remainder: segments
  };
}

function joinUrlPath(baseUrl: string, name: string): string {
  return `${baseUrl.replace(/\/$/, '')}/${encodeURIComponent(name).replace(/%2F/gi, '/')}`;
}

export function isSharePointViewerUrl(url: string | undefined): boolean {
  const parsed = safeParseUrl(url);
  if (!parsed) {
    return false;
  }
  return /\/_layouts\/15\/Doc\.aspx$/i.test(parsed.pathname);
}

export function extractSharePointSiteUrl(url: string | undefined): string | undefined {
  const parsed = safeParseUrl(url);
  if (!parsed) {
    return undefined;
  }

  const { sitePath } = getSitePathAndRemainder(parsed);
  return `${parsed.origin}${sitePath}`;
}

export function inferDocumentLibraryBaseUrl(urls: ReadonlyArray<string | undefined>): string | undefined {
  for (let i = 0; i < urls.length; i++) {
    const parsed = safeParseUrl(urls[i]);
    if (!parsed || isSharePointViewerUrl(parsed.toString())) {
      continue;
    }

    const { sitePath, remainder } = getSitePathAndRemainder(parsed);
    if (remainder.length === 0) {
      continue;
    }

    return `${parsed.origin}${sitePath}/${remainder[0]}`;
  }

  return undefined;
}

export function resolveDocumentLibraryItemUrl(
  itemUrl: string | undefined,
  itemName: string | undefined,
  siblingUrls: ReadonlyArray<string | undefined>
): string | undefined {
  if (!itemUrl) {
    return undefined;
  }
  if (!isSharePointViewerUrl(itemUrl)) {
    return itemUrl;
  }

  const baseUrl = inferDocumentLibraryBaseUrl(siblingUrls);
  if (baseUrl && itemName) {
    return joinUrlPath(baseUrl, itemName);
  }

  const parsed = safeParseUrl(itemUrl);
  const fileName = itemName || parsed?.searchParams.get('file') || undefined;
  if (!parsed || !fileName) {
    return itemUrl;
  }

  const { sitePath } = getSitePathAndRemainder(parsed);
  const fallbackLibrarySegment = parsed.searchParams.get('id');
  if (fallbackLibrarySegment) {
    const decoded = fallbackLibrarySegment.replace(/^\/+/, '');
    return `${parsed.origin}/${decoded.replace(/\/$/, '')}/${encodeURIComponent(fileName).replace(/%2F/gi, '/')}`;
  }

  return `${parsed.origin}${sitePath}/Shared%20Documents/${encodeURIComponent(fileName).replace(/%2F/gi, '/')}`;
}
