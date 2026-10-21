import { buildAcknowledgementPayload, IAcknowledgementRequest } from './AcknowledgementListService';

const ENTITY_TYPE = 'SP.Data.MI360_x0020_Data_x0020_Security_x0020_AcknowledgementsListItem';

function createRequest(overrides: Partial<IAcknowledgementRequest> = {}): IAcknowledgementRequest {
  return {
    identity: {
      entraObjectId: 'a1b2c3d4-0000-0000-0000-000000000000',
      displayName: 'Jane Doe',
      userPrincipalName: 'jane.doe@zf.com'
    },
    policyTitle: 'ZF Data Security Policy',
    source: 'MI360-SPFx-ApplicationCustomizer',
    requestedPage: '/sites/MI360/SitePages/Home.aspx',
    ...overrides
  };
}

describe('buildAcknowledgementPayload', () => {
  it('carries identity fields from the authenticated context, never from form input', () => {
    const payload = buildAcknowledgementPayload(createRequest(), ENTITY_TYPE);
    expect(payload.EntraObjectId).toBe('a1b2c3d4-0000-0000-0000-000000000000');
    expect(payload.UserDisplayName).toBe('Jane Doe');
    expect(payload.UserPrincipalName).toBe('jane.doe@zf.com');
    expect(payload.Title).toBe('Jane Doe');
  });

  it('uses the resolved list entity type name in __metadata', () => {
    const payload = buildAcknowledgementPayload(createRequest(), ENTITY_TYPE);
    expect(payload.__metadata).toEqual({ type: ENTITY_TYPE });
  });

  it('marks the record Active and stamps a UTC ISO timestamp', () => {
    const payload = buildAcknowledgementPayload(createRequest(), ENTITY_TYPE);
    expect(payload.RecordStatus).toBe('Active');
    expect(typeof payload.AcknowledgedAtUtc).toBe('string');
    expect(new Date(payload.AcknowledgedAtUtc as string).toISOString()).toBe(payload.AcknowledgedAtUtc);
  });

  it('records the requested page and policy title supplied by configuration', () => {
    const payload = buildAcknowledgementPayload(createRequest({ requestedPage: '/sites/MI360/Lists/Foo/AllItems.aspx' }), ENTITY_TYPE);
    expect(payload.RequestedPage).toBe('/sites/MI360/Lists/Foo/AllItems.aspx');
    expect(payload.PolicyTitle).toBe('ZF Data Security Policy');
  });
});
