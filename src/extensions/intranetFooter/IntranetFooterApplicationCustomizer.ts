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

// Site-wide intranet footer. Renders in the Bottom placeholder, which SharePoint places
// in-flow at the end of the page (so it only shows when you scroll to the bottom) - as
// long as the built-in site footer stays enabled to anchor it. Link columns come from a
// SharePoint list; brand/blurb/copyright/social/accent are extension properties. Never
// breaks the page.
export default class IntranetFooterApplicationCustomizer extends BaseApplicationCustomizer<IIntranetFooterProperties> {

  private _bottom: PlaceholderContent | undefined = undefined;
  private _links: IFooterLink[] = [];

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._render);
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = this.properties.listTitle || 'FooterLinks';
    return getFooterLinks(this.context.spHttpClient, webUrl, list).then((links) => {
      this._links = links;
      this._render();
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
  };

  private _onDispose = (): void => {
    if (this._bottom && this._bottom.domElement) {
      ReactDom.unmountComponentAtNode(this._bottom.domElement);
    }
  };

  protected onDispose(): void {
    this._onDispose();
  }
}
