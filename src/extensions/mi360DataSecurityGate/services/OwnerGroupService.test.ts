import { isMemberOfGroup } from './OwnerGroupService';

describe('isMemberOfGroup', () => {
  it('returns true when a group title matches exactly', () => {
    expect(isMemberOfGroup([{ Id: 1, Title: 'MI360 Owners' }], 'MI360 Owners')).toBe(true);
  });

  it('matches case-insensitively and trims whitespace', () => {
    expect(isMemberOfGroup([{ Id: 1, Title: '  mi360 owners  ' }], 'MI360 Owners')).toBe(true);
  });

  it('returns false when no group matches', () => {
    expect(isMemberOfGroup([{ Id: 1, Title: 'MI360 Members' }], 'MI360 Owners')).toBe(false);
  });

  it('returns false for an empty group list', () => {
    expect(isMemberOfGroup([], 'MI360 Owners')).toBe(false);
  });

  it('handles groups missing a Title without throwing', () => {
    expect(isMemberOfGroup([{ Id: 1 }], 'MI360 Owners')).toBe(false);
  });
});
