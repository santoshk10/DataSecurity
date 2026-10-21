import type { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';
import { IAuthenticatedIdentity } from './EntraIdentityService';

export interface IAcknowledgementRequest {
  identity: IAuthenticatedIdentity;
  policyTitle: string;
  source: string;
  requestedPage: string;
}

/**
 * Pure payload builder, separated from the network call so it can be unit tested without a
 * SharePoint context. Every value originates from the authenticated identity or fixed
 * configuration, never from editable form controls.
 */
export function buildAcknowledgementPayload(request: IAcknowledgementRequest, entityTypeFullName: string): Record<string, unknown> {
  return {
    __metadata: { type: entityTypeFullName },
    Title: request.identity.displayName,
    EntraObjectId: request.identity.entraObjectId,
    UserDisplayName: request.identity.displayName,
    UserPrincipalName: request.identity.userPrincipalName,
    AcknowledgedAtUtc: new Date().toISOString(),
    AcknowledgementSource: request.source,
    PolicyTitle: request.policyTitle,
    RequestedPage: request.requestedPage,
    RecordStatus: 'Active'
  };
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Reads and writes the MI360 Data Security Acknowledgements list. Duplicate submissions,
 * repeated selections from multiple tabs, and concurrent requests are treated as a single
 * successful acknowledgement: an existing active record for the same Entra Object ID is never
 * duplicated, and a write conflict is resolved by re-checking for an existing record.
 */
export class AcknowledgementListService {
  private entityTypeFullName: string | undefined;

  public constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly httpClientConfiguration: SPHttpClientConfiguration,
    private readonly webUrl: string,
    private readonly listTitle: string
  ) {}

  public async hasActiveAcknowledgement(entraObjectId: string): Promise<boolean> {
    const listTitle = escapeODataString(this.listTitle);
    const filterObjectId = escapeODataString(entraObjectId);
    const requestUrl = `${this.webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$filter=EntraObjectId eq '${filterObjectId}' and RecordStatus eq 'Active'&$select=Id&$top=1`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(requestUrl, this.httpClientConfiguration);
    if (!response.ok) {
      throw new Error(`Unable to query acknowledgement list status (status ${response.status}).`);
    }
    const json = await response.json();
    return Array.isArray(json.value) && json.value.length > 0;
  }

  /**
   * Creates the acknowledgement record if one does not already exist. Returns without error
   * whether the record was newly created or already existed, satisfying the "treat an existing
   * record as a successful acknowledgement" requirement.
   */
  public async ensureAcknowledgement(request: IAcknowledgementRequest): Promise<void> {
    if (await this.hasActiveAcknowledgement(request.identity.entraObjectId)) {
      return;
    }

    const entityTypeFullName = await this.getEntityTypeFullName();
    const payload = buildAcknowledgementPayload(request, entityTypeFullName);
    const listTitle = escapeODataString(this.listTitle);
    const createResponse: SPHttpClientResponse = await this.spHttpClient.post(
      `${this.webUrl}/_api/web/lists/getByTitle('${listTitle}')/items`,
      this.httpClientConfiguration,
      { body: JSON.stringify(payload) }
    );

    if (createResponse.ok) {
      return;
    }

    // A conflict, or a transient failure that still resulted in a write, is only acceptable if
    // an active record now exists; otherwise this is a genuine failure and must fail closed.
    if (await this.hasActiveAcknowledgement(request.identity.entraObjectId)) {
      return;
    }

    throw new Error(`Acknowledgement save failed (status ${createResponse.status}).`);
  }

  private async getEntityTypeFullName(): Promise<string> {
    if (this.entityTypeFullName) {
      return this.entityTypeFullName;
    }
    const listTitle = escapeODataString(this.listTitle);
    const requestUrl = `${this.webUrl}/_api/web/lists/getByTitle('${listTitle}')?$select=ListItemEntityTypeFullName`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(requestUrl, this.httpClientConfiguration);
    if (!response.ok) {
      throw new Error(`Unable to resolve acknowledgement list schema (status ${response.status}).`);
    }
    const json = await response.json();
    if (!json.ListItemEntityTypeFullName) {
      throw new Error('Acknowledgement list did not return an entity type name.');
    }
    this.entityTypeFullName = json.ListItemEntityTypeFullName;
    return this.entityTypeFullName as string;
  }
}
