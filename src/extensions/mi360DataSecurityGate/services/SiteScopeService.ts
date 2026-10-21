const TECHNICAL_EXCLUSION_SUFFIXES = [
  '/_layouts/15/validation.aspx',
  '/_layouts/15/error.aspx',
  '/_layouts/15/authenticate.aspx',
  '/_layouts/15/signin.aspx',
  '/_layouts/15/settings.aspx'
];

function getPathname(url: string): string | null {
  try {
    return new URL(url).pathname.toLowerCase();
  } catch {
    return null;
  }
}

/**
 * Determines whether the given URL falls within the gated site collection.
 */
export function isWithinSite(url: string, siteRootPath: string): boolean {
  const pathname = getPathname(url);
  if (pathname === null) {
    return false;
  }
  const normalizedRoot = siteRootPath.toLowerCase();
  return pathname === normalizedRoot || pathname.startsWith(`${normalizedRoot}/`);
}

/**
 * Narrow, explicit set of technical pages required to avoid redirect loops during
 * authentication, error handling, and validation. This must never be widened to a broad
 * pattern that could allow normal MI360 content to bypass the gate.
 */
export function isTechnicalExclusion(url: string, siteRootPath: string): boolean {
  const pathname = getPathname(url);
  if (pathname === null) {
    return false;
  }
  const normalizedRoot = siteRootPath.toLowerCase();
  return TECHNICAL_EXCLUSION_SUFFIXES.some((suffix) => pathname === `${normalizedRoot}${suffix}`);
}

export function isAcknowledgementPage(url: string, acknowledgementPageRelativeUrl: string): boolean {
  const pathname = getPathname(url);
  if (pathname === null) {
    return false;
  }
  return pathname === acknowledgementPageRelativeUrl.toLowerCase();
}
