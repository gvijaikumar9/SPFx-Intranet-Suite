import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import Footer, { IFooterProps } from './components/Footer';
import { groupLinks, parseSocial, IFooterLink } from './models/IFooter';
import { getFooterLinks } from './services/footerStore';

export interface IIntranetFooterProperties {
  brandText?: string;
  blurb?: string;
  copyright?: string;
  listTitle?: string;   // FooterLinks list; default 'FooterLinks'
  social?: string;      // "linkedin|https://..., email|mailto:..."
  accent?: string;      // hex; matches the site theme accent
}

const HOST_ID = 'intranet-suite-footer-host';

// The page's scrollable content region, in priority order. We append the footer as the
// LAST child of this region so it flows after all content and only shows at the bottom
// (NOT the pinned Bottom placeholder, which stays fixed on the screen).
const REGION_SELECTORS = [
  '[data-automation-id="contentScrollRegion"]',
  '#spPageCanvasContent',
  'div.CanvasZoneContainer',
  'main'
];

// Site-wide intranet footer. Injected at the end of the page canvas so it scrolls with
// the content. Link columns come from a SharePoint list; brand/blurb/copyright/social/
// accent are set as extension properties. Never breaks the page.
export default class IntranetFooterApplicationCustomizer extends BaseApplicationCustomizer<IIntranetFooterProperties> {

  private _links: IFooterLink[] = [];
  private _observer: MutationObserver | undefined = undefined;
  private _pending: number | undefined = undefined;

  public onInit(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = this.properties.listTitle || 'FooterLinks';
    return getFooterLinks(this.context.spHttpClient, webUrl, list).then((links) => {
      this._links = links;
      this._ensure();
      // The canvas renders asynchronously and re-renders on SPA navigation, so keep the
      // footer placed. A cheap, debounced, idempotent check on body mutations covers it.
      this._observer = new MutationObserver(() => this._schedule());
      this._observer.observe(document.body, { childList: true, subtree: true });
    });
  }

  private _schedule(): void {
    if (this._pending !== undefined) { return; }
    this._pending = window.setTimeout(() => {
      this._pending = undefined;
      this._ensure();
    }, 300);
  }

  private _findRegion(): HTMLElement | undefined {
    for (let i = 0; i < REGION_SELECTORS.length; i++) {
      const el = document.querySelector(REGION_SELECTORS[i]) as HTMLElement | null;
      if (el) { return el; }
    }
    return undefined;
  }

  // Idempotent: place the host as the region's last child (if missing/moved) and render.
  private _ensure(): void {
    const region = this._findRegion();
    if (!region) { return; }

    let host = document.getElementById(HOST_ID) as HTMLDivElement | null;
    if (!host) {
      host = document.createElement('div');
      host.id = HOST_ID;
    }
    if (host.parentElement !== region || region.lastElementChild !== host) {
      region.appendChild(host); // move/append as the LAST child = after all content
    }
    this._render(host);
  }

  private _render(host: HTMLElement): void {
    const element: React.ReactElement<IFooterProps> = React.createElement(Footer, {
      brandText: this.properties.brandText || 'Contoso',
      blurb: this.properties.blurb || '',
      copyright: this.properties.copyright || '',
      groups: groupLinks(this._links),
      social: parseSocial(this.properties.social),
      accent: this.properties.accent || '#0f6cbd'
    });
    ReactDom.render(element, host);
  }

  protected onDispose(): void {
    if (this._observer) { this._observer.disconnect(); this._observer = undefined; }
    if (this._pending !== undefined) { window.clearTimeout(this._pending); this._pending = undefined; }
    const host = document.getElementById(HOST_ID);
    if (host) {
      ReactDom.unmountComponentAtNode(host);
      if (host.parentElement) { host.parentElement.removeChild(host); }
    }
  }
}
