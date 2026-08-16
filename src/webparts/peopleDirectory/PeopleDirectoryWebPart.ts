import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'PeopleDirectoryWebPartStrings';
import PeopleDirectory from './components/PeopleDirectory';
import { IPeopleDirectoryProps } from './components/IPeopleDirectoryProps';
import { IPerson } from './models/IPerson';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IPeopleDirectoryWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  showPhoto: boolean;
  useDemoData: boolean;
}

const DEMO_PEOPLE: IPerson[] = [
  { login: 'd1', name: 'Priya Nair', title: 'Engineering Manager', department: 'Platform' },
  { login: 'd2', name: 'Sam Okoro', title: 'Senior Developer', department: 'Platform' },
  { login: 'd3', name: 'Liam Chen', title: 'Product Designer', department: 'Design' },
  { login: 'd4', name: 'Maria Gomez', title: 'People Partner', department: 'HR' },
  { login: 'd5', name: 'Ahmed Ali', title: 'Support Lead', department: 'IT' },
  { login: 'd6', name: 'Elena Petrova', title: 'Data Analyst', department: 'Finance' }
];

export default class PeopleDirectoryWebPart extends BaseClientSideWebPart<IPeopleDirectoryWebPartProps> {

  private _accent: string = '#0f6cbd';

  public render(): void {
    const isDemo = this.properties.useDemoData;
    const element: React.ReactElement<IPeopleDirectoryProps> = React.createElement(
      PeopleDirectory,
      {
        title: this.properties.title || 'Find a colleague',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        showPhoto: this.properties.showPhoto !== false,
        webUrl: this.context.pageContext.web.absoluteUrl,
        isDemo: isDemo,
        accent: this._accent,
        demoPeople: isDemo ? DEMO_PEOPLE : [],
        onSearch: (query: string) => this._onSearch(query)
      }
    );
    ReactDom.render(element, this.domElement);
  }

  private _onSearch(query: string): Promise<IPerson[]> {
    if (this.properties.useDemoData) {
      const q = query.toLowerCase();
      return Promise.resolve(DEMO_PEOPLE.filter((p) => p.name.toLowerCase().indexOf(q) >= 0));
    }
    return this._searchPeople(query);
  }

  /** Live people search against the site's client people picker service. */
  private async _searchPeople(query: string): Promise<IPerson[]> {
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const body = JSON.stringify({
      queryParams: {
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        MaximumEntitySuggestions: 12,
        PrincipalSource: 15,
        PrincipalType: 1,
        QueryString: query
      }
    });
    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata', 'Content-type': 'application/json;odata=nometadata' },
        body: body
      });
      if (!res.ok) { return []; }
      const json = await res.json();
      const payload = json.value ?? (json.d && json.d.ClientPeoplePickerSearchUser);
      if (!payload) { return []; }
      const entities: Array<{ Key: string; DisplayText: string; EntityData?: { Email?: string; Title?: string; Department?: string } }> = JSON.parse(payload);
      return entities.map((e): IPerson => ({
        login: e.Key,
        name: e.DisplayText,
        email: e.EntityData && e.EntityData.Email ? e.EntityData.Email : undefined,
        title: e.EntityData && e.EntityData.Title ? e.EntityData.Title : undefined,
        department: e.EntityData && e.EntityData.Department ? e.EntityData.Department : undefined
      }));
    } catch {
      return [];
    }
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh();
    }
    if (path === 'useDemoData') {
      this._safeRender();
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
    this.properties.title = 'Find a colleague';
    this.properties.layout = 'card';
    this.properties.showPhoto = true;
    this.properties.useDemoData = true;
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.context.propertyPane.refresh();
    this._safeRender();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneToggle('showPhoto', { label: strings.ShowPhotoLabel, onText: 'Profile photo', offText: 'Initials' })
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneToggle('useDemoData', { label: strings.UseDemoLabel, onText: 'On', offText: 'Off' })
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
