import type { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IAuthenticatedIdentity {
  entraObjectId: string;
  displayName: string;
  userPrincipalName: string;
}

interface IUserProfileProperty {
  Key?: string;
  Value?: string;
}

/**
 * Retrieves the current user's permanent Entra Object ID and profile details from the
 * authenticated SharePoint context. Values are never accepted from editable form controls.
 */
export class EntraIdentityService {
  public constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly httpClientConfiguration: SPHttpClientConfiguration,
    private readonly webUrl: string
  ) {}

  public async getCurrentEntraObjectId(): Promise<string | null> {
    const requestUrl = `${this.webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=UserProfileProperties`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(requestUrl, this.httpClientConfiguration);
    if (!response.ok) {
      throw new Error(`Unable to read user profile properties (status ${response.status}).`);
    }
    const properties = await response.json();
    const entries: IUserProfileProperty[] = Array.isArray(properties.UserProfileProperties) ? properties.UserProfileProperties : [];
    const entry = entries.find((item) => item.Key === 'SPS-ObjectId' || item.Key === 'SPS-ObjectId-Encrypted');
    return entry?.Value ? String(entry.Value).trim() : null;
  }

  public async getCurrentIdentity(): Promise<IAuthenticatedIdentity> {
    const objectId = await this.getCurrentEntraObjectId();
    if (!objectId) {
      throw new Error('Authenticated user is missing an Entra Object ID.');
    }

    const requestUrl = `${this.webUrl}/_api/web/currentuser?$select=Title,Email,LoginName`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(requestUrl, this.httpClientConfiguration);
    if (!response.ok) {
      throw new Error(`Unable to determine current user details (status ${response.status}).`);
    }
    const currentUser = await response.json();
    const displayName = (currentUser.Title || currentUser.LoginName || '').trim();
    const userPrincipalName = (currentUser.Email || currentUser.LoginName || '').trim();

    if (!displayName || !userPrincipalName) {
      throw new Error('Authenticated user context did not return a display name or principal name.');
    }

    return { entraObjectId: objectId, displayName, userPrincipalName };
  }
}
