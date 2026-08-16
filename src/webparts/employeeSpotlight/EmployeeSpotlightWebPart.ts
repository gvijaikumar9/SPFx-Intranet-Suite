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

import * as strings from 'EmployeeSpotlightWebPartStrings';
import EmployeeSpotlight from './components/EmployeeSpotlight';
import { IEmployeeSpotlightProps } from './components/IEmployeeSpotlightProps';
import { IEmployee } from './models/IEmployee';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IEmployeeSpotlightWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  showPhoto: boolean;
  listTitle: string;
  maxItems: number;
  useDemoData: boolean;
}

const DEMO_ITEMS: IEmployee[] = [
  { id: 1, name: 'Arjun Rao', month: 'This month', citation: 'For steady delivery and mentoring across two teams.' },
  { id: 2, name: 'Elena Petrova', month: 'Last month', citation: 'Rebuilt the onboarding guide and cut setup time in half.' },
  { id: 3, name: 'Marcus Lee', month: 'Two months ago', citation: 'Kept the release calm under real pressure.' }
];

export default class EmployeeSpotlightWebPart extends BaseClientSideWebPart<IEmployeeSpotlightWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: IEmployee[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;
  private _emptyNote: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IEmployeeSpotlightProps> = React.createElement(
      EmployeeSpotlight,
      {
        title: this.properties.title || 'Employee of the month',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        isDemo: isDemo,
        accent: this._accent,
        webUrl: this.context.pageContext.web.absoluteUrl,
        showPhoto: this.properties.showPhoto !== false,
        items: this._items,
        loading: this._loading,
        error: this._error,
        emptyNote: this._emptyNote
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
    this._emptyNote = undefined;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._items = DEMO_ITEMS.slice(0, this.properties.maxItems || 4);
      this._loading = false;
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const top = this.properties.maxItems || 4;
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=Id,Title,EomMonth,Citation,Employee/Title,Employee/EMail`
      + `&$expand=Employee&$top=${top}&$orderby=Created desc`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load recognition (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        const rows = json.value || [];
        // The Employee person field is the source of truth for the name and photo.
        // Items with no Employee selected are not valid recognition entries, so skip them
        // rather than falling back to the Title text (which is a headline, not a name).
        this._items = rows
          .map((row: Record<string, unknown>): IEmployee | undefined => {
            const emp = row.Employee as { Title?: string; EMail?: string } | undefined;
            if (!emp || !emp.Title) { return undefined; }
            return {
              id: Number(row.Id ?? 0),
              name: emp.Title,
              month: (row.EomMonth ?? '').toString(),
              citation: (row.Citation ?? '').toString(),
              email: emp.EMail || undefined
            };
          })
          .filter((e: IEmployee | undefined): e is IEmployee => e !== undefined);

        // Help the admin: rows exist but none have a person selected.
        if (rows.length > 0 && this._items.length === 0) {
          this._emptyNote = 'These recognition items have no Employee selected yet. Set the Employee field on the item.';
        }
      }
    } catch {
      this._error = 'Could not load recognition.';
      this._items = [];
    }

    this._loading = false;
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh(); // enable/disable the custom-colour field
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

  /** Restore every property to its shipped default and reload. */
  private _resetToDefault(): void {
    this.properties.title = 'Employee of the month';
    this.properties.layout = 'card';
    this.properties.showPhoto = true;
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.maxItems = 4;
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
                PropertyPaneToggle('showPhoto', { label: strings.ShowPhotoLabel, onText: 'Profile photo', offText: 'Initials' }),
                PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 1, max: 10, step: 1 })
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
