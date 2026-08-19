import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
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

// The page's scroll container (modern SharePoint scrolls an inner region, not the window).
const REGION_SELECTORS = [
  '[data-automation-id="contentScrollRegion"]',
  '#spPageCanvasContent'
];
const REVEAL_THRESHOLD = 40; // px from the bottom at which the footer appears

// Site-wide intranet footer. Renders in the Bottom placeholder (which modern SharePoint
// pins to the viewport). To match "only show at the bottom of the page", the footer is
// hidden until the page is scrolled to the bottom, then revealed. Link columns come from
// a SharePoint list; brand/blurb/copyright/social/accent are extension properties.
export default class IntranetFooterApplicationCustomizer extends BaseApplicationCustomizer<IIntranetFooterProperties> {

  private _bottom: PlaceholderContent | undefined = undefined;
  private _links: IFooterLink[] = [];
  private _pending: number | undefined = undefined;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._render);
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = this.properties.listTitle || 'FooterLinks';
    return getFooterLinks(this.context.spHttpClient, webUrl, list).then((links) => {
      this._links = links;
      this._render();
      // scroll happens on an inner region, so capture scroll events from any element
      document.addEventListener('scroll', this._schedule, true);
      window.addEventListener('resize', this._schedule);
    });
  }

  private _render = (): void => {
    if (!this._bottom) {
      this._bottom = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );
    }
    if (!this._bottom || !this._bottom.domElement) { return; } // placeholder not ready yet

    const element: React.ReactElement<IFooterProps> = React.createElement(Footer, {
      brandText: this.properties.brandText || 'Contoso',
      blurb: this.properties.blurb || '',
      copyright: this.properties.copyright || '',
      groups: groupLinks(this._links),
      social: parseSocial(this.properties.social),
      accent: this.properties.accent || '#0f6cbd'
    });
    ReactDom.render(element, this._bottom.domElement);
    // start hidden; a short delay lets the canvas lay out before the first check
    this._bottom.domElement.style.display = 'none';
    window.setTimeout(() => this._apply(), 500);
  };

  private _schedule = (): void => {
    if (this._pending !== undefined) { return; }
    this._pending = window.setTimeout(() => { this._pending = undefined; this._apply(); }, 120);
  };

  private _scroller(): HTMLElement | undefined {
    for (let i = 0; i < REGION_SELECTORS.length; i++) {
      const el = document.querySelector(REGION_SELECTORS[i]) as HTMLElement | null;
      if (el && el.scrollHeight > el.clientHeight + 2) { return el; }
    }
    return undefined;
  }

  private _atBottom(): boolean {
    const region = this._scroller();
    if (region) {
      return region.scrollTop + region.clientHeight >= region.scrollHeight - REVEAL_THRESHOLD;
    }
    const doc = document.documentElement;
    const bottom = window.innerHeight + (window.pageYOffset || doc.scrollTop);
    return bottom >= doc.scrollHeight - REVEAL_THRESHOLD;
  }

  private _apply(): void {
    if (!this._bottom || !this._bottom.domElement) { return; }
    this._bottom.domElement.style.display = this._atBottom() ? '' : 'none';
  }

  private _onDispose = (): void => {
    if (this._bottom && this._bottom.domElement) {
      ReactDom.unmountComponentAtNode(this._bottom.domElement);
    }
  };

  protected onDispose(): void {
    document.removeEventListener('scroll', this._schedule, true);
    window.removeEventListener('resize', this._schedule);
    if (this._pending !== undefined) { window.clearTimeout(this._pending); }
    this._onDispose();
  }
}
