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

const HOST_ID = 'intranet-suite-footer';
const PAD_STYLE_ID = 'intranet-suite-footer-pad';
// The page scrolls an inner region in modern SharePoint; pad it so content clears the
// fixed footer. Selectors tried in order.
const REGION_SELECTORS = ['[data-automation-id="contentScrollRegion"]', '#spPageCanvasContent'];

// Site-wide intranet footer. It is a position:fixed overlay mounted on <body>, so it
// shows on every page and occupies NO layout space - which lets us hide/show it on scroll
// with zero flicker (toggling a fixed element never reflows the page). It is hidden until
// the page is scrolled to the bottom, and the scroll region is padded by the footer's
// height so the last content is never covered.
export default class IntranetFooterApplicationCustomizer extends BaseApplicationCustomizer<IIntranetFooterProperties> {

  private _links: IFooterLink[] = [];
  private _host: HTMLDivElement | undefined = undefined;
  private _height: number = 0;
  private _pending: number | undefined = undefined;

  public onInit(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = this.properties.listTitle || 'FooterLinks';
    return getFooterLinks(this.context.spHttpClient, webUrl, list).then((links) => {
      this._links = links;
      this._mount();
      document.addEventListener('scroll', this._schedule, true); // capture inner-region scroll too
      window.addEventListener('resize', this._onResize);
    });
  }

  private _mount(): void {
    let host = document.getElementById(HOST_ID) as HTMLDivElement | null;
    if (!host) {
      host = document.createElement('div');
      host.id = HOST_ID;
      host.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:30;display:none;';
      document.body.appendChild(host);
    }
    this._host = host;

    const element: React.ReactElement<IFooterProps> = React.createElement(Footer, {
      brandText: this.properties.brandText || 'Contoso',
      blurb: this.properties.blurb || '',
      copyright: this.properties.copyright || '',
      groups: groupLinks(this._links),
      social: parseSocial(this.properties.social),
      accent: this.properties.accent || '#0f6cbd'
    });
    ReactDom.render(element, host);

    // measure once laid out, pad the scroll region, then apply the scroll state
    window.setTimeout(() => { this._measure(); this._apply(); }, 500);
  }

  private _measure(): void {
    if (!this._host) { return; }
    // measure without flashing it on screen
    const prev = this._host.style.display;
    this._host.style.visibility = 'hidden';
    this._host.style.display = 'block';
    this._height = this._host.offsetHeight || 0;
    this._host.style.display = prev;
    this._host.style.visibility = '';
    this._pad();
  }

  // Static padding on the scroll region (via a stylesheet so React re-renders keep it),
  // so the last content sits clear of the fixed footer. Static = no reflow = no flicker.
  private _pad(): void {
    if (this._height <= 0) { return; }
    let style = document.getElementById(PAD_STYLE_ID) as HTMLStyleElement | null;
    if (!style) {
      style = document.createElement('style');
      style.id = PAD_STYLE_ID;
      document.head.appendChild(style);
    }
    const sel = REGION_SELECTORS.join(', ');
    style.textContent = `${sel} { padding-bottom: ${this._height}px !important; }`;
  }

  private _scroller(): HTMLElement | undefined {
    for (let i = 0; i < REGION_SELECTORS.length; i++) {
      const el = document.querySelector(REGION_SELECTORS[i]) as HTMLElement | null;
      if (el && el.scrollHeight > el.clientHeight + 2) { return el; }
    }
    return undefined;
  }

  private _atBottom(): boolean {
    const reveal = Math.max(this._height, 40); // reveal as the padded zone comes into view
    const region = this._scroller();
    if (region) {
      return region.scrollHeight - region.scrollTop - region.clientHeight <= reveal;
    }
    const doc = document.documentElement;
    return doc.scrollHeight - (window.innerHeight + (window.pageYOffset || doc.scrollTop)) <= reveal;
  }

  private _schedule = (): void => {
    if (this._pending !== undefined) { return; }
    this._pending = window.setTimeout(() => { this._pending = undefined; this._apply(); }, 80);
  };

  private _onResize = (): void => { this._measure(); this._apply(); };

  private _apply(): void {
    if (!this._host) { return; }
    // toggling a position:fixed element never reflows the page, so this cannot flicker
    this._host.style.display = this._atBottom() ? 'block' : 'none';
  }

  protected onDispose(): void {
    document.removeEventListener('scroll', this._schedule, true);
    window.removeEventListener('resize', this._onResize);
    if (this._pending !== undefined) { window.clearTimeout(this._pending); }
    if (this._host) {
      ReactDom.unmountComponentAtNode(this._host);
      if (this._host.parentElement) { this._host.parentElement.removeChild(this._host); }
    }
    const pad = document.getElementById(PAD_STYLE_ID);
    if (pad && pad.parentElement) { pad.parentElement.removeChild(pad); }
  }
}
