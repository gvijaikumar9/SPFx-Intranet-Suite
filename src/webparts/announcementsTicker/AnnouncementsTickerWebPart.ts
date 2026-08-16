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

import * as strings from 'AnnouncementsTickerWebPartStrings';
import AnnouncementsTicker from './components/AnnouncementsTicker';
import { IAnnouncementsTickerProps } from './components/IAnnouncementsTickerProps';
import { IAnnouncement, normalizeSeverity } from './models/IAnnouncement';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IAnnouncementsTickerWebPartProps extends IStandardWebPartProps {
  label: string;
  layout: string;
  listTitle: string;
  messageField: string;
  severityField: string;
  linkField: string;
  maxItems: number;
  rotateSeconds: number;
  showControls: boolean;
  useDemoData: boolean;
}

const DEMO_ITEMS: IAnnouncement[] = [
  { id: 'd1', message: 'Payroll cut-off is this Friday at 5pm.', severity: 'warning' },
  { id: 'd2', message: 'All-hands town hall on Thursday 10am, main auditorium.', severity: 'info' },
  { id: 'd3', message: 'VPN maintenance Saturday 6 to 8am, expect brief drops.', severity: 'critical' },
  { id: 'd4', message: 'Open enrolment for benefits starts Nov 1.', severity: 'info' }
];

export default class AnnouncementsTickerWebPart extends BaseClientSideWebPart<IAnnouncementsTickerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _accent: string = '#0f6cbd';
  private _items: IAnnouncement[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;
  private _listOptions: { key: string; text: string }[] = [];

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IAnnouncementsTickerProps> = React.createElement(
      AnnouncementsTicker,
      {
        label: this.properties.label || 'Alerts',
        items: this._items,
        rotateSeconds: this.properties.rotateSeconds || 6,
        showControls: this.properties.showControls !== false,
        loading: this._loading,
        error: this._error,
        isDemo: isDemo,
        accent: this._accent,
        isDarkTheme: this._isDarkTheme,
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent)
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

  /** Populate the property-pane list dropdown with the site's non-hidden document/custom lists. */
  private async _loadListOptions(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists?$filter=Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 104)`
      + `&$select=Title&$orderby=Title`;
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return; }
    const json = await res.json();
    this._listOptions = (json.value || []).map((l: { Title: string }) => ({ key: l.Title, text: l.Title }));
  }

  /** Load announcements from the configured list, or fall back to demo data. */
  private async _loadItems(): Promise<void> {
    this._loading = true;
    this._error = undefined;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._items = DEMO_ITEMS.slice(0, this.properties.maxItems || 8);
      this._loading = false;
      return;
    }

    const msgField = (this.properties.messageField || 'Title').trim();
    const sevField = (this.properties.severityField || '').trim();
    const linkField = (this.properties.linkField || '').trim();
    const top = this.properties.maxItems || 8;

    const select = ['Id', msgField]
      .concat(sevField ? [sevField] : [])
      .concat(linkField ? [linkField] : [])
      .join(',');
    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=${encodeURIComponent(select)}&$top=${top}&$orderby=Modified desc`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load announcements (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>, i: number): IAnnouncement => {
          const rawLink = linkField ? row[linkField] : undefined;
          let href: string | undefined;
          if (rawLink && typeof rawLink === 'object' && 'Url' in (rawLink as object)) {
            href = (rawLink as { Url?: string }).Url;
          } else if (typeof rawLink === 'string' && rawLink) {
            href = rawLink;
          }
          return {
            id: (row.Id ?? i).toString(),
            message: (row[msgField] ?? '').toString(),
            severity: sevField ? normalizeSeverity(row[sevField]) : 'info',
            href: href
          };
        }).filter((a: IAnnouncement) => a.message.length > 0);
      }
    } catch {
      this._error = 'Could not load announcements.';
      this._items = [];
    }

    this._loading = false;
  }

  /**
   * Render only when BOTH the DOM element and the properties bag exist. onThemeChanged
   * fires during _internalInitialize, before this.properties is set, so guarding on
   * domElement alone still let render() dereference an undefined this.properties.
   */
  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh(); // enable/disable the custom-colour field
    }
    // Re-fetch when anything affecting the data changes.
    const dataProps = ['listTitle', 'messageField', 'severityField', 'linkField', 'maxItems', 'useDemoData'];
    if (dataProps.indexOf(path) >= 0) {
      if (path === 'useDemoData') {
        this.context.propertyPane.refresh(); // toggles the disabled state of the list picker
      }
      this._loading = true;
      this._safeRender();
      this._loadItems().then(() => this._safeRender()).catch(() => this._safeRender());
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) { return; }
    this._isDarkTheme = !!currentTheme.isInverted;
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
    this.properties.label = 'Alerts';
    this.properties.layout = 'card';
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.messageField = 'Title';
    this.properties.severityField = '';
    this.properties.linkField = '';
    this.properties.maxItems = 8;
    this.properties.rotateSeconds = 6;
    this.properties.showControls = true;
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
                PropertyPaneTextField('label', { label: strings.LabelFieldLabel }),
                PropertyPaneDropdown('layout', {
                  label: strings.LayoutFieldLabel,
                  options: [
                    { key: 'card', text: 'Card' },
                    { key: 'minimal', text: 'Minimal' },
                    { key: 'bold', text: 'Bold' },
                    { key: 'compact', text: 'Compact' }
                  ]
                }),
                PropertyPaneSlider('rotateSeconds', { label: strings.RotateFieldLabel, min: 3, max: 20, step: 1 }),
                PropertyPaneToggle('showControls', { label: strings.ShowControlsLabel })
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
                PropertyPaneTextField('messageField', { label: strings.MessageFieldLabel, disabled: demo }),
                PropertyPaneTextField('severityField', { label: strings.SeverityFieldLabel, disabled: demo }),
                PropertyPaneTextField('linkField', { label: strings.LinkFieldLabel, disabled: demo }),
                PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 1, max: 20, step: 1 }),
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
