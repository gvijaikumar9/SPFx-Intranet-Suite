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

import * as strings from 'ImageGalleryWebPartStrings';
import ImageGallery from './components/ImageGallery';
import { IImageGalleryProps } from './components/IImageGalleryProps';
import { IImage, isImageName } from './models/IImage';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IImageGalleryWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  columns: number;
  listTitle: string;
  useDemoData: boolean;
}

const DEMO_ITEMS: IImage[] = [
  { id: 1, name: 'Team offsite', url: '' },
  { id: 2, name: 'Award night', url: '' },
  { id: 3, name: 'Volunteer day', url: '' },
  { id: 4, name: 'Product launch', url: '' },
  { id: 5, name: 'Summer social', url: '' },
  { id: 6, name: 'New office', url: '' },
  { id: 7, name: 'Hackathon', url: '' },
  { id: 8, name: 'Town hall', url: '' }
];

export default class ImageGalleryWebPart extends BaseClientSideWebPart<IImageGalleryWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: IImage[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IImageGalleryProps> = React.createElement(
      ImageGallery,
      {
        title: this.properties.title || 'Gallery',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        columns: this.properties.columns || 4,
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
    // BaseTemplate 101 = document library.
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 101`
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
      + `?$select=Id,FileLeafRef,FileRef,FileSystemObjectType&$top=60&$orderby=Modified desc`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `Library "${this.properties.listTitle}" was not found.`
          : `Could not load photos (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || [])
          .filter((row: Record<string, unknown>) =>
            Number(row.FileSystemObjectType ?? 1) === 0 && isImageName((row.FileLeafRef ?? '').toString()))
          .map((row: Record<string, unknown>): IImage => ({
            id: Number(row.Id ?? 0),
            name: (row.FileLeafRef ?? '').toString(),
            url: (row.FileRef ?? '').toString()
          }));
      }
    } catch {
      this._error = 'Could not load photos.';
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
    this.properties.title = 'Gallery';
    this.properties.layout = 'card';
    this.properties.columns = 4;
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
                PropertyPaneSlider('columns', { label: strings.ColumnsFieldLabel, min: 2, max: 6, step: 1 })
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
