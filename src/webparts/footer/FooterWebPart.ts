import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FooterWebPartStrings';
import Footer, { IFooterProps } from './components/Footer';
import { groupLinks, parseSocial, IFooterLink } from './models/IFooter';
import { getFooterLinks } from './services/footerStore';

export interface IFooterWebPartProps {
  brandText: string;
  blurb: string;
  copyright: string;
  listTitle: string;   // FooterLinks list
  social: string;      // "linkedin|https://..., email|mailto:..."
}

// A full-width footer web part. Place it in the last (full-width) section of a page so it
// sits at the end of the content and only shows when you scroll to the bottom. Link
// columns come from the FooterLinks list; brand/blurb/copyright/social are properties;
// the accent follows the site theme. Never breaks the page on a data failure.
export default class FooterWebPart extends BaseClientSideWebPart<IFooterWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _links: IFooterLink[] = [];

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
      const list = this.properties.listTitle || 'FooterLinks';
      this._links = await getFooterLinks(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, list);
    });
  }

  public render(): void {
    const element: React.ReactElement<IFooterProps> = React.createElement(Footer, {
      brandText: this.properties.brandText || 'Contoso',
      blurb: this.properties.blurb || '',
      copyright: this.properties.copyright || '',
      groups: groupLinks(this._links),
      social: parseSocial(this.properties.social),
      accent: this._accent
    });
    ReactDom.render(element, this.domElement);
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'listTitle') {
      getFooterLinks(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, this.properties.listTitle || 'FooterLinks')
        .then((links) => { this._links = links; this._safeRender(); })
        .catch(() => this._safeRender());
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) { return; }
    if (currentTheme.palette && currentTheme.palette.themePrimary) {
      this._accent = currentTheme.palette.themePrimary;
    }
    this._safeRender();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BrandGroupName,
              groupFields: [
                PropertyPaneTextField('brandText', { label: strings.BrandFieldLabel }),
                PropertyPaneTextField('blurb', { label: strings.BlurbFieldLabel, multiline: true }),
                PropertyPaneTextField('copyright', { label: strings.CopyrightFieldLabel })
              ]
            },
            {
              groupName: strings.LinksGroupName,
              groupFields: [
                PropertyPaneTextField('listTitle', { label: strings.ListFieldLabel }),
                PropertyPaneTextField('social', { label: strings.SocialFieldLabel, multiline: true }),
                PropertyPaneLabel('fieldHint', { text: strings.FieldHint })
              ]
            }
          ]
        }
      ]
    };
  }
}
