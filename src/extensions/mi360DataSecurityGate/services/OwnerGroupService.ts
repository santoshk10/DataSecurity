import type { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';

interface ISiteGroup {
  Id?: number;
  Title?: string;
}

/**
 * Pure membership check, separated from the network call so it can be unit tested
 * without a SharePoint context.
 */
export function isMemberOfGroup(groups: ISiteGroup[], groupName: string): boolean {
  const normalizedTarget = groupName.trim().toLowerCase();
  return groups.some((group) => (group.Title || '').trim().toLowerCase() === normalizedTarget);
}

/**
 * Determines whether the current signed-in user belongs to the configured owners group.
 * Fails closed: any non-OK response or network failure throws, so the caller treats the
 * user as a non-owner and continues through the acknowledgement gate rather than bypassing it.
 */
export class OwnerGroupService {
  public constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly httpClientConfiguration: SPHttpClientConfiguration,
    private readonly webUrl: string
  ) {}

  public async isCurrentUserOwner(groupName: string): Promise<boolean> {
    const requestUrl = `${this.webUrl}/_api/web/currentuser/groups?$select=Id,Title`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(requestUrl, this.httpClientConfiguration);
    if (!response.ok) {
      throw new Error(`Unable to resolve current user group membership (status ${response.status}).`);
    }
    const json = await response.json();
    const groups: ISiteGroup[] = Array.isArray(json.value) ? json.value : [];
    return isMemberOfGroup(groups, groupName);
  }
}
