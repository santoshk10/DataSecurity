import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Log, SPEventArgs } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AcknowledgementPage } from './components/AcknowledgementPage';
import { OwnerGroupService } from './services/OwnerGroupService';
import { EntraIdentityService } from './services/EntraIdentityService';
import { AcknowledgementListService } from './services/AcknowledgementListService';
import { isAcknowledgementPage, isTechnicalExclusion, isWithinSite } from './services/SiteScopeService';
import { validateReturnUrl } from './utils/ReturnUrlValidator';

export interface IMi360DataSecurityGateApplicationCustomizerProperties {
  siteRelativeRootPath?: string;
  ownersGroupName?: string;
  acknowledgementListTitle?: string;
  acknowledgementPageRelativeUrl?: string;
  policyTitle?: string;
  policyUrl?: string;
  supportUrl?: string;
}

const LOG_SOURCE = 'Mi360DataSecurityGateApplicationCustomizer';
const DEFAULT_SITE_ROOT = '/sites/MI360';
const DEFAULT_LIST_TITLE = 'MI360 Data Security Acknowledgements';
const ACKNOWLEDGEMENT_SOURCE = 'MI360-SPFx-ApplicationCustomizer';
const CONTAINER_ID = 'mi360DataSecurityGateContainer';

export default class Mi360DataSecurityGateApplicationCustomizer extends BaseApplicationCustomizer<IMi360DataSecurityGateApplicationCustomizerProperties> {
  private _hostElement: HTMLElement | null = null;
  private _isEvaluating = false;
  private _pendingReEvaluation = false;
  /** Non-authoritative, in-memory only. Never persisted to storage; cleared on full reload. */
  private readonly _acknowledgedObjectIds = new Set<string>();
  private readonly _onNavigated = (): void => {
    this._runGate().catch((error) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });
  };

  public async onInit(): Promise<void> {
    this.context.application.navigatedEvent.add(this, this._onNavigated as (eventArgs: SPEventArgs) => void);
    await this._runGate();
  }

  public onDispose(): void {
    this.context.application.navigatedEvent.remove(this, this._onNavigated as (eventArgs: SPEventArgs) => void);
    this._unmountAcknowledgementUi();
  }

  private get siteRootPath(): string {
    return this.properties.siteRelativeRootPath || DEFAULT_SITE_ROOT;
  }

  private get acknowledgementPageRelativeUrl(): string {
    return (this.properties.acknowledgementPageRelativeUrl || `${this.siteRootPath}/SitePages/MI360-Data-Security-Acknowledgement.aspx`).toLowerCase();
  }

  private get webUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private async _runGate(): Promise<void> {
    // Coalesce overlapping navigation events into a single follow-up run rather than
    // allowing concurrent evaluations to race against each other.
    if (this._isEvaluating) {
      this._pendingReEvaluation = true;
      return;
    }
    this._isEvaluating = true;
    try {
      await this._evaluateGate();
    } finally {
      this._isEvaluating = false;
      if (this._pendingReEvaluation) {
        this._pendingReEvaluation = false;
        await this._runGate();
      }
    }
  }

  private async _evaluateGate(): Promise<void> {
    try {
      const currentUrl = window.location.href;
      if (!isWithinSite(currentUrl, this.siteRootPath)) {
        return;
      }
      if (isTechnicalExclusion(currentUrl, this.siteRootPath)) {
        return;
      }
      if (isAcknowledgementPage(currentUrl, this.acknowledgementPageRelativeUrl)) {
        this._renderAcknowledgementUi();
        return;
      }

      this._unmountAcknowledgementUi();

      const ownerGroupService = new OwnerGroupService(this.context.spHttpClient, SPHttpClient.configurations.v1, this.webUrl);
      const isOwner = await ownerGroupService.isCurrentUserOwner(this.properties.ownersGroupName || 'MI360 Owners');
      if (isOwner) {
        return;
      }

      const identityService = new EntraIdentityService(this.context.spHttpClient, SPHttpClient.configurations.v1, this.webUrl);
      const objectId = await identityService.getCurrentEntraObjectId();
      if (!objectId) {
        this._redirectToAcknowledgementPage(this.siteRootPath);
        return;
      }

      if (this._acknowledgedObjectIds.has(objectId)) {
        return;
      }

      const listService = new AcknowledgementListService(
        this.context.spHttpClient,
        SPHttpClient.configurations.v1,
        this.webUrl,
        this.properties.acknowledgementListTitle || DEFAULT_LIST_TITLE
      );
      if (await listService.hasActiveAcknowledgement(objectId)) {
        this._acknowledgedObjectIds.add(objectId);
        return;
      }

      this._redirectToAcknowledgementPage(currentUrl);
    } catch (error) {
      // Fail closed: any failure while evaluating owner status or acknowledgement status
      // keeps the user in the acknowledgement experience instead of allowing access.
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Unknown gate failure'));
      this._renderAcknowledgementUi();
    }
  }

  private _redirectToAcknowledgementPage(returnUrlCandidate: string): void {
    const validatedReturnUrl = validateReturnUrl(returnUrlCandidate, this.siteRootPath, window.location.origin);
    const target = new URL(this.acknowledgementPageRelativeUrl, window.location.origin);
    target.searchParams.set('returnUrl', validatedReturnUrl);
    window.location.assign(target.toString());
  }

  private _unmountAcknowledgementUi(): void {
    if (this._hostElement) {
      ReactDOM.unmountComponentAtNode(this._hostElement);
      this._hostElement.remove();
      this._hostElement = null;
    }
  }

  private _renderAcknowledgementUi(): void {
    const currentReturn = validateReturnUrl(
      new URL(window.location.href).searchParams.get('returnUrl') || this.siteRootPath,
      this.siteRootPath,
      window.location.origin
    );

    this._unmountAcknowledgementUi();
    const container = document.createElement('div');
    container.id = CONTAINER_ID;
    container.style.position = 'relative';
    container.style.zIndex = '1000';
    (document.body || document.documentElement).insertBefore(container, (document.body || document.documentElement).firstChild);
    this._hostElement = container;

    ReactDOM.render(
      <AcknowledgementPage
        returnUrl={currentReturn}
        policyUrl={this.properties.policyUrl || 'https://www.zf.com/corporate/en/our-company/data-protection.html'}
        supportUrl={this.properties.supportUrl || `${this.siteRootPath}/SitePages/MI360-Support.aspx`}
        onAcknowledge={async (redirectTarget: string) => {
          const identityService = new EntraIdentityService(this.context.spHttpClient, SPHttpClient.configurations.v1, this.webUrl);
          const identity = await identityService.getCurrentIdentity();

          const listService = new AcknowledgementListService(
            this.context.spHttpClient,
            SPHttpClient.configurations.v1,
            this.webUrl,
            this.properties.acknowledgementListTitle || DEFAULT_LIST_TITLE
          );

          await listService.ensureAcknowledgement({
            identity,
            policyTitle: this.properties.policyTitle || 'ZF Data Security Policy',
            source: ACKNOWLEDGEMENT_SOURCE,
            requestedPage: redirectTarget || this.siteRootPath
          });

          this._acknowledgedObjectIds.add(identity.entraObjectId);
          const destination = validateReturnUrl(redirectTarget, this.siteRootPath, window.location.origin);
          window.location.assign(destination);
        }}
      />,
      container
    );
  }
}