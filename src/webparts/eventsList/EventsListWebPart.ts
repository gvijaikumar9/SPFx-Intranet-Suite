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

import * as strings from 'EventsListWebPartStrings';
import EventsList from './components/EventsList';
import { IEventsListProps } from './components/IEventsListProps';
import { IEvent } from './models/IEvent';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IEventsListWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  maxItems: number;
  listTitle: string;
  useDemoData: boolean;
}

export default class EventsListWebPart extends BaseClientSideWebPart<IEventsListWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: IEvent[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IEventsListProps> = React.createElement(
      EventsList,
      {
        title: this.properties.title || 'Upcoming events',
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
    // BaseTemplate 106 = calendar/events list.
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists?$filter=Hidden eq false and (BaseTemplate eq 106 or BaseTemplate eq 100)`
      + `&$select=Title&$orderby=Title`;
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return; }
    const json = await res.json();
    this._listOptions = (json.value || []).map((l: { Title: string }) => ({ key: l.Title, text: l.Title }));
  }

  private async _loadItems(): Promise<void> {
    this._loading = true;
    this._error = undefined;
    const top = this.properties.maxItems || 5;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      const base = new Date();
      const mk = (days: number, hour: number): string => {
        const d = new Date(base.getTime());
        d.setDate(d.getDate() + days);
        d.setHours(hour, 0, 0, 0);
        return d.toISOString();
      };
      this._items = [
        { id: 1, title: 'All-hands town hall', start: mk(3, 10), end: mk(3, 11), location: 'Auditorium' },
        { id: 2, title: 'Lunch and learn', start: mk(7, 12), end: mk(7, 13), location: 'Cafe' },
        { id: 3, title: 'Quarterly review', start: mk(12, 14), end: mk(12, 16), location: 'Room 4B' }
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
      + `?$select=Id,Title,EventDate,EndDate,Location`
      + `&$filter=EventDate ge datetime'${nowIso}'&$orderby=EventDate&$top=${top}`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load events (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>): IEvent => ({
          id: Number(row.Id ?? 0),
          title: (row.Title ?? '').toString(),
          start: (row.EventDate ?? '').toString(),
          end: row.EndDate ? row.EndDate.toString() : undefined,
          location: row.Location ? row.Location.toString() : undefined
        })).filter((e: IEvent) => e.title.length > 0);
      }
    } catch {
      this._error = 'Could not load events.';
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
    this.properties.title = 'Upcoming events';
    this.properties.layout = 'card';
    this.properties.maxItems = 5;
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
