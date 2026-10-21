import { validateReturnUrl } from './ReturnUrlValidator';

const ORIGIN = 'https://trw1.sharepoint.com';
const ROOT = '/sites/MI360';

describe('validateReturnUrl', () => {
  it('returns the site root when the candidate is empty', () => {
    expect(validateReturnUrl('', ROOT, ORIGIN)).toBe(ROOT);
    expect(validateReturnUrl(null, ROOT, ORIGIN)).toBe(ROOT);
    expect(validateReturnUrl(undefined, ROOT, ORIGIN)).toBe(ROOT);
  });

  it('accepts the site root itself', () => {
    expect(validateReturnUrl('/sites/MI360', ROOT, ORIGIN)).toBe('/sites/MI360');
  });

  it('accepts a nested site-relative path', () => {
    expect(validateReturnUrl('/sites/MI360/SitePages/Home.aspx', ROOT, ORIGIN)).toBe('/sites/MI360/SitePages/Home.aspx');
  });

  it('preserves query string and hash on a valid path', () => {
    expect(validateReturnUrl('/sites/MI360/Lists/Foo/AllItems.aspx?View=1#top', ROOT, ORIGIN)).toBe(
      '/sites/MI360/Lists/Foo/AllItems.aspx?View=1#top'
    );
  });

  it('is case-insensitive when matching the site root', () => {
    expect(validateReturnUrl('/SITES/mi360/SitePages/Home.aspx', ROOT, ORIGIN)).toBe('/SITES/mi360/SitePages/Home.aspx');
  });

  it('rejects a path outside the site root', () => {
    expect(validateReturnUrl('/sites/OtherSite/Home.aspx', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects a path that merely starts with the root as a text prefix', () => {
    expect(validateReturnUrl('/sites/MI360EVIL/Home.aspx', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects protocol-relative URLs', () => {
    expect(validateReturnUrl('//evil.com/phish', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects absolute external URLs', () => {
    expect(validateReturnUrl('https://evil.com/sites/MI360', ROOT, ORIGIN)).toBe(ROOT);
    expect(validateReturnUrl('http://evil.com', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects javascript: and data: URLs', () => {
    expect(validateReturnUrl('javascript:alert(1)', ROOT, ORIGIN)).toBe(ROOT);
    expect(validateReturnUrl('data:text/html,<script>alert(1)</script>', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects backslash-based protocol-relative bypass attempts', () => {
    expect(validateReturnUrl('/\\evil.com', ROOT, ORIGIN)).toBe(ROOT);
    expect(validateReturnUrl('\\\\evil.com', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects single-encoded open-redirect attempts', () => {
    expect(validateReturnUrl('%2F%2Fevil.com', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects double-encoded open-redirect attempts', () => {
    expect(validateReturnUrl('%252F%252Fevil.com', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects path traversal that escapes the site root', () => {
    expect(validateReturnUrl('/sites/MI360/../../evil', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects malformed paths', () => {
    expect(validateReturnUrl('not-a-path', ROOT, ORIGIN)).toBe(ROOT);
  });

  it('rejects embedded control characters', () => {
    expect(validateReturnUrl('/sites/MI360/%0d%0aSet-Cookie:evil', ROOT, ORIGIN)).toBe(ROOT);
  });
});
