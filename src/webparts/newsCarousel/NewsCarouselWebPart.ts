import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'NewsCarouselWebPartStrings';
import NewsCarousel from './components/NewsCarousel';
import { INewsCarouselProps } from './components/INewsCarouselProps';
import { INews, urlValue } from './models/INews';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface INewsCarouselWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  rotateSeconds: number;
  listTitle: string;
  useDemoData: boolean;
}

const DEMO_ITEMS: INews[] = [
  { id: 1, title: 'Contoso hits 99.98% platform uptime this quarter', summary: 'Engineering and operations share how the reliability program paid off.', category: 'Company news' },
  { id: 2, title: 'Meet the 2026 Innovation Award winners', summary: 'Five teams recognised for standout customer impact this year.', category: 'People' },
  { id: 3, title: 'The new HQ collaboration floors open Monday', summary: 'Book a tour and see what changed on levels 4 through 6.', category: 'Workplace' }
];

export default class NewsCarouselWebPart extends BaseClientSideWebPart<INewsCarouselWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: INews[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<INewsCarouselProps> = React.createElement(
      NewsCarousel,
      {
        title: this.properties.title || 'News',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        rotateSeconds: this.properties.rotateSeconds || 6,
        isDemo: isDemo,
        accent: this._accent,
        items: this._items,
        loading: this._loading,
        error: this._error
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
      await this._loadListOptions().catch(() => { /* dropdown stays empty; not fatal */ });
      await this._loadItems();
    });
  }

  private async _loadListOptions(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100`
      + `&$select=Title&$orderby=Title`;
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return; }
    const json = await res.json();
    this._listOptions = (json.value || []).map((l: { Title: string }) => ({ key: l.Title, text: l.Title }));
  }

  private async _loadItems(): Promise<void> {
    this._loading = true;
    this._error = undefined;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._items = DEMO_ITEMS;
      this._loading = false;
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=Id,Title,Summary,Category,NewsImage,NewsLink&$top=10&$orderby=Modified desc`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load news (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>): INews => ({
          id: Number(row.Id ?? 0),
          title: (row.Title ?? '').toString(),
          summary: row.Summary ? row.Summary.toString() : undefined,
          category: row.Category ? row.Category.toString() : undefined,
          imageUrl: urlValue(row.NewsImage) || undefined,
          linkUrl: urlValue(row.NewsLink) || undefined
        })).filter((n: INews) => n.title.length > 0);
      }
    } catch {
      this._error = 'Could not load news.';
      this._items = [];
    }

    this._loading = false;
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh();
    }
    const dataProps = ['listTitle', 'useDemoData'];
    if (dataProps.indexOf(path) >= 0) {
      if (path === 'useDemoData') {
        this.context.propertyPane.refresh();
      }
      this._loading = true;
      this._safeRender();
      this._loadItems().then(() => this._safeRender()).catch(() => this._safeRender());
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

  private _resetToDefault(): void {
    this.properties.title = 'News';
    this.properties.layout = 'card';
    this.properties.rotateSeconds = 6;
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.useDemoData = true;
    this.context.propertyPane.refresh();
    this._loading = true;
    this._safeRender();
    this._loadItems().then(() => this._safeRender()).catch(() => this._safeRender());
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const demo = this.properties.useDemoData;
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                PropertyPaneDropdown('layout', {
                  label: strings.LayoutFieldLabel,
                  options: [
                    { key: 'card', text: 'Card' },
                    { key: 'minimal', text: 'Minimal' },
                    { key: 'bold', text: 'Bold' },
                    { key: 'compact', text: 'Compact' }
                  ]
                }),
                PropertyPaneSlider('rotateSeconds', { label: strings.RotateFieldLabel, min: 3, max: 20, step: 1 })
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneToggle('useDemoData', { label: strings.UseDemoLabel, onText: 'On', offText: 'Off' }),
                PropertyPaneDropdown('listTitle', {
                  label: strings.ListFieldLabel,
                  options: this._listOptions,
                  disabled: demo
                }),
                PropertyPaneLabel('fieldHint', { text: strings.FieldHint })
              ]
            },
            {
              groupName: 'Container',
              groupFields: standardPaneFields(this.properties.backgroundMode || 'transparent', () => this._resetToDefault())
            }
          ]
        }
      ]
    };
  }
}
