/**
 * Validates that a candidate return URL is a safe, site-relative path within the gated site
 * collection. Used to prevent open-redirect attacks (external hosts, protocol-relative URLs,
 * `javascript:`/`data:` schemes, encoded payloads, and path traversal out of the site root).
 */

const CONTROL_CHARACTER_PATTERN = /[\u0000-\u001f\u007f]/;
const SCHEME_PATTERN = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;
const MAX_DECODE_ATTEMPTS = 3;

function safeDecode(candidate: string): string | null {
  let current = candidate;
  for (let attempt = 0; attempt < MAX_DECODE_ATTEMPTS; attempt++) {
    let next: string;
    try {
      next = decodeURIComponent(current);
    } catch {
      return null;
    }
    if (next === current) {
      return current;
    }
    current = next;
  }
  return current;
}

/**
 * Returns a validated, site-relative return URL, or the site root if the candidate is missing,
 * malformed, or points outside the gated site collection.
 *
 * @param candidate - The raw, untrusted return URL supplied by the browser/query string.
 * @param siteRootPath - The site-relative root, e.g. `/sites/MI360` (no trailing slash).
 * @param origin - The current window origin, used to resolve and re-validate the final path.
 */
export function validateReturnUrl(candidate: string | null | undefined, siteRootPath: string, origin: string): string {
  const normalizedRoot = siteRootPath.replace(/\/+$/, '').toLowerCase();
  const fallback = siteRootPath;

  if (!candidate || !candidate.trim()) {
    return fallback;
  }

  const decoded = safeDecode(candidate.trim());
  if (decoded === null) {
    return fallback;
  }

  const normalizedSlashes = decoded.replace(/\\/g, '/');

  if (CONTROL_CHARACTER_PATTERN.test(normalizedSlashes)) {
    return fallback;
  }

  if (SCHEME_PATTERN.test(normalizedSlashes)) {
    return fallback;
  }

  if (!normalizedSlashes.startsWith('/') || normalizedSlashes.startsWith('//')) {
    return fallback;
  }

  const lowerPath = normalizedSlashes.toLowerCase();
  if (lowerPath !== normalizedRoot && !lowerPath.startsWith(`${normalizedRoot}/`)) {
    return fallback;
  }

  try {
    const resolved = new URL(normalizedSlashes, origin);
    if (resolved.origin !== origin) {
      return fallback;
    }
    const resolvedLowerPath = resolved.pathname.toLowerCase();
    if (resolvedLowerPath !== normalizedRoot && !resolvedLowerPath.startsWith(`${normalizedRoot}/`)) {
      return fallback;
    }
    return `${resolved.pathname}${resolved.search}${resolved.hash}`;
  } catch {
    return fallback;
  }
}
