import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import Footer, { IFooterProps } from './components/Footer';
import { parseSocial, IFooterLink } from './models/IFooter';
import { getFooterLinks } from './services/footerStore';

export interface IIntranetFooterProperties {
  brandText?: string;
  copyright?: string;
  listTitle?: string;   // FooterLinks list; default 'FooterLinks'
  social?: string;      // "linkedin|https://..., email|mailto:..."
  accent?: string;      // hex; matches the site theme accent
}

const REGION_SELECTORS = ['[data-automation-id="contentScrollRegion"]', '#spPageCanvasContent'];
const REVEAL = 44; // px from the bottom at which the slim bar appears

// Site-wide intranet footer: a slim full-width bar in the Bottom placeholder, shown on
// every page. With the built-in site footer disabled (Enable-Footer does this), the
// placeholder is position:fixed, so it occupies no layout space - which lets us hide it
// until the page is scrolled to the bottom with NO flicker (toggling a fixed element can
// never reflow the page).
export default class IntranetFooterApplicationCustomizer extends BaseApplicationCustomizer<IIntranetFooterProperties> {

  private _bottom: PlaceholderContent | undefined = undefined;
  private _links: IFooterLink[] = [];
  private _pending: number | undefined = undefined;
  private _placed: boolean = false;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._render);
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = this.properties.listTitle || 'FooterLinks';
    return getFooterLinks(this.context.spHttpClient, webUrl, list).then((links) => {
      this._links = links;
      this._render();
      document.addEventListener('scroll', this._schedule, true); // capture inner-region scroll
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
      copyright: this.properties.copyright || '',
      links: this._links,   // already sorted by FooterOrder
      social: parseSocial(this.properties.social),
      accent: this.properties.accent || '#0f6cbd'
    });
    ReactDom.render(element, this._bottom.domElement);
    if (!this._placed) {
      this._placed = true;
      this._bottom.domElement.style.display = 'none'; // hide only on first placement
      window.setTimeout(() => this._apply(), 400);
    }
  };

  // Only scroll-reveal when the footer is actually pinned (a fixed/sticky context). If it
  // renders in-flow, keep it always rendered - an in-flow footer already shows only at the
  // end of the page, and never hiding it means no reflow and no flicker.
  private _isPinned(): boolean {
    let el: HTMLElement | null = this._bottom ? this._bottom.domElement : null;
    let depth = 0;
    while (el && depth < 6) {
      const pos = window.getComputedStyle(el).position;
      if (pos === 'fixed' || pos === 'sticky') { return true; }
      el = el.parentElement;
      depth++;
    }
    return false;
  }

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
      return region.scrollHeight - region.scrollTop - region.clientHeight <= REVEAL;
    }
    const doc = document.documentElement;
    // no inner scroll region (short page): treat as "at bottom" so it shows
    return doc.scrollHeight - (window.innerHeight + (window.pageYOffset || doc.scrollTop)) <= REVEAL;
  }

  private _schedule = (): void => {
    if (this._pending !== undefined) { return; }
    this._pending = window.setTimeout(() => { this._pending = undefined; this._apply(); }, 80);
  };

  private _apply(): void {
    if (!this._bottom || !this._bottom.domElement) { return; }
    if (!this._isPinned()) {
      this._bottom.domElement.style.display = ''; // in-flow: always render, never toggle
      return;
    }
    // pinned (fixed): toggling never reflows the page, so this cannot flicker
    this._bottom.domElement.style.display = this._atBottom() ? '' : 'none';
  }

  private _onDispose = (): void => {
    if (this._bottom && this._bottom.domElement) {
      ReactDom.unmountComponentAtNode(this._bottom.domElement);
    }
  };

  protected onDispose(): void {
    this.context.placeholderProvider.changedEvent.remove(this, this._render);
    document.removeEventListener('scroll', this._schedule, true);
    window.removeEventListener('resize', this._schedule);
    if (this._pending !== undefined) { window.clearTimeout(this._pending); }
    this._onDispose();
  }
}
