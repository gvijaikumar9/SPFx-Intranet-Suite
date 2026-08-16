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

import * as strings from 'CelebrationsWebPartStrings';
import Celebrations from './components/Celebrations';
import { ICelebrationsProps } from './components/ICelebrationsProps';
import { ICelebration } from './models/ICelebration';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface ICelebrationsWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  maxItems: number;
  showPhoto: boolean;
  listTitle: string;
  useDemoData: boolean;
}

const DEMO_ITEMS: ICelebration[] = [
  { id: 1, name: 'Priya Nair', date: '2026-09-04', type: 'Birthday' },
  { id: 2, name: 'Marcus Lee', date: '2026-09-12', type: 'Work anniversary' },
  { id: 3, name: 'Elena Petrova', date: '2026-10-02', type: 'Birthday' },
  { id: 4, name: 'Arjun Rao', date: '2026-10-19', type: 'Work anniversary' }
];

export default class CelebrationsWebPart extends BaseClientSideWebPart<ICelebrationsWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: ICelebration[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<ICelebrationsProps> = React.createElement(
      Celebrations,
      {
        title: this.properties.title || 'Celebrations',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        showPhoto: this.properties.showPhoto !== false,
        webUrl: this.context.pageContext.web.absoluteUrl,
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

  /** Order by the next upcoming occurrence (birthdays/anniversaries recur each year). */
  private _sortUpcoming(items: ICelebration[]): ICelebration[] {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const keyed = items.map((c) => {
      const d = new Date(c.date);
      let next = today + 1e14; // invalid dates sort to the end
      if (!isNaN(d.getTime())) {
        let occ = new Date(now.getFullYear(), d.getMonth(), d.getDate()).getTime();
        if (occ < today) { occ = new Date(now.getFullYear() + 1, d.getMonth(), d.getDate()).getTime(); }
        next = occ;
      }
      return { c: c, next: next };
    });
    keyed.sort((a, b) => a.next - b.next);
    return keyed.map((k) => k.c);
  }

  private async _loadItems(): Promise<void> {
    this._loading = true;
    this._error = undefined;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._items = this._sortUpcoming(DEMO_ITEMS).slice(0, this.properties.maxItems || 6);
      this._loading = false;
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=Id,Title,CelebrationDate,CelebrationType,Person/EMail&$expand=Person&$top=500`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load celebrations (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        const mapped = (json.value || []).map((row: Record<string, unknown>): ICelebration => {
          const person = row.Person as { EMail?: string } | undefined;
          return {
            id: Number(row.Id ?? 0),
            name: (row.Title ?? '').toString(),
            date: (row.CelebrationDate ?? '').toString(),
            type: (row.CelebrationType ?? '').toString(),
            email: person && person.EMail ? person.EMail : undefined
          };
        }).filter((c: ICelebration) => c.name.length > 0);
        this._items = this._sortUpcoming(mapped).slice(0, this.properties.maxItems || 6);
      }
    } catch {
      this._error = 'Could not load celebrations.';
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
    const dataProps = ['listTitle', 'maxItems', 'useDemoData'];
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
    this.properties.title = 'Celebrations';
    this.properties.layout = 'card';
    this.properties.maxItems = 6;
    this.properties.showPhoto = true;
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
                PropertyPaneToggle('showPhoto', { label: strings.ShowPhotoLabel, onText: 'Photo', offText: 'Initials' }),
                PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 1, max: 12, step: 1 })
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
