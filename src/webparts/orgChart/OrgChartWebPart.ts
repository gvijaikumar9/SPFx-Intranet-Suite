import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'OrgChartWebPartStrings';
import OrgChart from './components/OrgChart';
import { IOrgChartProps } from './components/IOrgChartProps';
import { IOrgNode } from './models/IOrgNode';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IOrgChartWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  showPhoto: boolean;
  useDemoData: boolean;
}

const DEMO_MANAGER: IOrgNode = { name: 'Dana Whitfield', title: 'Director of Engineering' };
const DEMO_ME: IOrgNode = { name: 'Priya Nair', title: 'Engineering Manager' };
const DEMO_REPORTS: IOrgNode[] = [
  { name: 'Sam Okoro', title: 'Senior Developer' },
  { name: 'Liam Chen', title: 'Developer' },
  { name: 'Ahmed Ali', title: 'QA Engineer' },
  { name: 'Elena Petrova', title: 'Data Analyst' }
];

/** Some SP collection endpoints return a bare array, others { results: [] }. Normalise. */
function asArray(v: unknown): string[] {
  if (Array.isArray(v)) { return v as string[]; }
  if (v && typeof v === 'object' && Array.isArray((v as { results?: unknown[] }).results)) {
    return (v as { results: string[] }).results;
  }
  return [];
}

export default class OrgChartWebPart extends BaseClientSideWebPart<IOrgChartWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _manager: IOrgNode | undefined = undefined;
  private _me: IOrgNode | undefined = undefined;
  private _reports: IOrgNode[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData;
    const element: React.ReactElement<IOrgChartProps> = React.createElement(
      OrgChart,
      {
        title: this.properties.title || 'Team',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        showPhoto: this.properties.showPhoto !== false,
        webUrl: this.context.pageContext.web.absoluteUrl,
        isDemo: isDemo,
        accent: this._accent,
        manager: isDemo ? DEMO_MANAGER : this._manager,
        me: isDemo ? DEMO_ME : this._me,
        reports: isDemo ? DEMO_REPORTS : this._reports,
        loading: this._loading,
        error: this._error
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
      if (!this.properties.useDemoData) {
        await this._loadOrg();
      } else {
        this._loading = false;
      }
    });
  }

  private _get(url: string): Promise<SPHttpClientResponse> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' }
    });
  }

  private async _resolve(account: string): Promise<IOrgNode | undefined> {
    if (!account) { return undefined; }
    const v = encodeURIComponent(account);
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${v}'`
      + `&$select=DisplayName,Email,Title`;
    const res = await this._get(url);
    if (!res.ok) { return undefined; }
    const p = await res.json();
    if (!p || !p.DisplayName) { return undefined; }
    return { name: p.DisplayName, title: p.Title || undefined, email: p.Email || undefined };
  }

  private async _loadOrg(): Promise<void> {
    this._loading = true;
    this._error = undefined;
    try {
      const meUrl = `${this.context.pageContext.web.absoluteUrl}`
        + `/_api/SP.UserProfiles.PeopleManager/GetMyProperties`
        + `?$select=DisplayName,Email,Title,ExtendedManagers,DirectReports`;
      const meRes = await this._get(meUrl);
      if (!meRes.ok) {
        this._error = `Could not read your profile (HTTP ${meRes.status}).`;
        this._me = undefined;
        this._loading = false;
        return;
      }
      const me = await meRes.json();
      this._me = { name: me.DisplayName || 'You', title: me.Title || undefined, email: me.Email || undefined };

      const managers = asArray(me.ExtendedManagers);
      const reports = asArray(me.DirectReports).slice(0, 8);

      // The immediate manager is the last entry in the top-down manager chain.
      const mgrAccount = managers.length > 0 ? managers[managers.length - 1] : '';
      this._manager = mgrAccount ? await this._resolve(mgrAccount) : undefined;

      const resolved = await Promise.all(reports.map((a) => this._resolve(a)));
      this._reports = resolved.filter((r): r is IOrgNode => !!r);
    } catch {
      this._error = 'Could not load the org chart.';
      this._me = undefined;
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
    if (path === 'useDemoData') {
      this.context.propertyPane.refresh();
      if (!this.properties.useDemoData) {
        this._loading = true;
        this._safeRender();
        this._loadOrg().then(() => this._safeRender()).catch(() => this._safeRender());
      } else {
        this._safeRender();
      }
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
    this.properties.title = 'Team';
    this.properties.layout = 'card';
    this.properties.showPhoto = true;
    this.properties.useDemoData = true;
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.context.propertyPane.refresh();
    this._loading = false;
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
                PropertyPaneToggle('useDemoData', { label: strings.UseDemoLabel, onText: 'On', offText: 'Off' }),
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
