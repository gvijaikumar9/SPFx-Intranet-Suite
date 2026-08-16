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

import * as strings from 'HolidaysListWebPartStrings';
import HolidaysList from './components/HolidaysList';
import { IHolidaysListProps } from './components/IHolidaysListProps';
import { IHoliday } from './models/IHoliday';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IHolidaysListWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  maxItems: number;
  listTitle: string;
  useDemoData: boolean;
}

export default class HolidaysListWebPart extends BaseClientSideWebPart<IHolidaysListWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: IHoliday[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IHolidaysListProps> = React.createElement(
      HolidaysList,
      {
        title: this.properties.title || 'Upcoming holidays',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
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
    const top = this.properties.maxItems || 6;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      const base = new Date();
      const mk = (days: number): string => {
        const d = new Date(base.getTime());
        d.setDate(d.getDate() + days);
        d.setHours(0, 0, 0, 0);
        return d.toISOString();
      };
      this._items = [
        { id: 1, name: 'Thanksgiving', date: mk(20), region: 'US' },
        { id: 2, name: 'Diwali', date: mk(35), region: 'IN' },
        { id: 3, name: 'Christmas Day', date: mk(60), region: 'Global' }
      ].slice(0, top);
      this._loading = false;
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const nowIso = now.toISOString().replace(/\.\d+Z$/, 'Z');
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=Id,Title,HolidayDate,Region`
      + `&$filter=HolidayDate ge datetime'${nowIso}'&$orderby=HolidayDate&$top=${top}`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load holidays (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>): IHoliday => ({
          id: Number(row.Id ?? 0),
          name: (row.Title ?? '').toString(),
          date: (row.HolidayDate ?? '').toString(),
          region: row.Region ? row.Region.toString() : undefined
        })).filter((h: IHoliday) => h.name.length > 0);
      }
    } catch {
      this._error = 'Could not load holidays.';
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
    this.properties.title = 'Upcoming holidays';
    this.properties.layout = 'card';
    this.properties.maxItems = 6;
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
                PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 1, max: 15, step: 1 })
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
