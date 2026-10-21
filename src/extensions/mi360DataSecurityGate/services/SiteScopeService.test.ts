import { isAcknowledgementPage, isTechnicalExclusion, isWithinSite } from './SiteScopeService';

const ORIGIN = 'https://trw1.sharepoint.com';
const ROOT = '/sites/MI360';

describe('isWithinSite', () => {
  it('matches the site root itself', () => {
    expect(isWithinSite(`${ORIGIN}/sites/MI360`, ROOT)).toBe(true);
  });

  it('matches nested pages under the root', () => {
    expect(isWithinSite(`${ORIGIN}/sites/MI360/SitePages/Home.aspx`, ROOT)).toBe(true);
  });

  it('is case-insensitive', () => {
    expect(isWithinSite(`${ORIGIN}/Sites/MI360/SitePages/Home.aspx`, ROOT)).toBe(true);
  });

  it('rejects a different site collection', () => {
    expect(isWithinSite(`${ORIGIN}/sites/OtherSite`, ROOT)).toBe(false);
  });

  it('rejects a site whose name merely starts with the root', () => {
    expect(isWithinSite(`${ORIGIN}/sites/MI360Evil`, ROOT)).toBe(false);
  });

  it('rejects an unparseable URL', () => {
    expect(isWithinSite('not-a-url', ROOT)).toBe(false);
  });
});

describe('isTechnicalExclusion', () => {
  it('matches the approved technical pages', () => {
    expect(isTechnicalExclusion(`${ORIGIN}/sites/MI360/_layouts/15/authenticate.aspx`, ROOT)).toBe(true);
    expect(isTechnicalExclusion(`${ORIGIN}/sites/MI360/_layouts/15/error.aspx`, ROOT)).toBe(true);
  });

  it('does not match normal MI360 content pages', () => {
    expect(isTechnicalExclusion(`${ORIGIN}/sites/MI360/SitePages/Home.aspx`, ROOT)).toBe(false);
  });
});

describe('isAcknowledgementPage', () => {
  const ackPage = '/sites/mi360/sitepages/mi360-data-security-acknowledgement.aspx';

  it('matches the configured acknowledgement page', () => {
    expect(isAcknowledgementPage(`${ORIGIN}/sites/MI360/SitePages/MI360-Data-Security-Acknowledgement.aspx`, ackPage)).toBe(true);
  });

  it('does not match other pages', () => {
    expect(isAcknowledgementPage(`${ORIGIN}/sites/MI360/SitePages/Home.aspx`, ackPage)).toBe(false);
  });
});
