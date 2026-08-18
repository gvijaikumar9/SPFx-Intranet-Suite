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

import * as strings from 'TabsContainerWebPartStrings';
import TabsContainer from './components/TabsContainer';
import { ITabsContainerProps } from './components/ITabsContainerProps';
import { ITabItem, tabsOf } from './models/ITabItem';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';
import { resolveTileColors, collectCustomKeys, tileColorModeOptions, tileColorPickerOptions, MAX_TILE_COLORS } from '../shared/tileColors';

export interface ITabsContainerWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  listTitle: string;
  tabField: string;
  subtitleField: string;
  linkField: string;
  useDemoData: boolean;
  tileColorMode: string;   // none | palette | fluro | water | dark | custom
  tileColor1: string;
  tileColor2: string;
  tileColor3: string;
  tileColor4: string;
  tileColor5: string;
  tileColor6: string;
  tileColor7: string;
  tileColor8: string;
}

const DEMO_ITEMS: ITabItem[] = [
  { id: 1, tab: 'Company news', title: 'Q3 all-hands recap and the roadmap ahead', subtitle: 'Riya Thomas, 2 hours ago' },
  { id: 2, tab: 'Company news', title: 'Open enrolment for 2027 benefits starts Nov 1', subtitle: 'People team, yesterday' },
  { id: 3, tab: 'Events', title: 'All-hands town hall', subtitle: 'Thu 10:00, Auditorium' },
  { id: 4, tab: 'Events', title: 'Lunch and learn: Copilot for analysts', subtitle: 'Fri 12:30, Level 4' },
  { id: 5, tab: 'Links', title: 'Book leave', subtitle: 'HR portal', href: '#' },
  { id: 6, tab: 'Links', title: 'IT support', subtitle: 'Service desk', href: '#' }
];

function urlValue(raw: unknown): string | undefined {
  if (!raw) { return undefined; }
  let url: string;
  if (typeof raw === 'object' && 'Url' in (raw as object)) {
    url = ((raw as { Url?: string }).Url || '').trim();
  } else {
    const s = raw.toString();
    const comma = s.indexOf(',');
    url = (comma >= 0 ? s.substring(0, comma) : s).trim();
  }
  if (!url) { return undefined; }
  // Allow only safe schemes; drop javascript:/data:/vbscript: etc.
  return /^(https?:|mailto:|#|\/)/i.test(url) ? url : undefined;
}

export default class TabsContainerWebPart extends BaseClientSideWebPart<ITabsContainerWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: ITabItem[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<ITabsContainerProps> = React.createElement(
      TabsContainer,
      {
        title: this.properties.title || "What's happening",
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        isDemo: isDemo,
        accent: this._accent,
        items: this._items,
        tabColors: this._tabColors(),
        loading: this._loading,
        error: this._error
      }
    );
    ReactDom.render(element, this.domElement);
  }

  private _tabColors(): (string | undefined)[] {
    const n = tabsOf(this._items || []).length;
    return resolveTileColors(
      this.properties.tileColorMode || 'none',
      n,
      collectCustomKeys(this.properties, n)
    );
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

    const tabF = (this.properties.tabField || 'Category').trim();
    const subF = (this.properties.subtitleField || '').trim();
    const linkF = (this.properties.linkField || '').trim();
    const select = ['Id', 'Title', tabF]
      .concat(subF ? [subF] : [])
      .concat(linkF ? [linkF] : [])
      .join(',');
    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=${encodeURIComponent(select)}&$top=200&$orderby=Modified desc`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load items (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>): ITabItem => ({
          id: Number(row.Id ?? 0),
          tab: (row[tabF] ?? 'General').toString(),
          title: (row.Title ?? '').toString(),
          subtitle: subF && row[subF] ? String(row[subF]) : undefined,
          href: linkF ? urlValue(row[linkF]) : undefined
        })).filter((it: ITabItem) => it.title.length > 0);
      }
    } catch {
      this._error = 'Could not load items.';
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
    const dataProps = ['listTitle', 'tabField', 'subtitleField', 'linkField', 'useDemoData'];
    if (dataProps.indexOf(path) >= 0) {
      if (path === 'useDemoData') {
        this.context.propertyPane.refresh();
      }
      this._loading = true;
      this._safeRender();
      this._loadItems().then(() => this._safeRender()).catch(() => this._safeRender());
    }
    if (path === 'tileColorMode') {
      this.context.propertyPane.refresh(); // show or hide the per-tab pickers
    }
    if (path.indexOf('tileColor') === 0) {
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
    this.properties.title = "What's happening";
    this.properties.layout = 'card';
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.tabField = 'Category';
    this.properties.subtitleField = '';
    this.properties.linkField = '';
    this.properties.useDemoData = true;
    this.properties.tileColorMode = 'none';
    const p = this.properties as unknown as { [k: string]: string };
    for (let i = 1; i <= MAX_TILE_COLORS; i++) { p['tileColor' + i] = ''; }
    this.context.propertyPane.refresh();
    this._loading = true;
    this._safeRender();
    this._loadItems().then(() => this._safeRender()).catch(() => this._safeRender());
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const demo = this.properties.useDemoData;

    // Tab colours group: a mode dropdown, plus a picker per tab (named) in Custom mode.
    const tabNames = tabsOf(this._items || []);
    const tabColorFields = [
      PropertyPaneDropdown('tileColorMode', { label: 'Tab colours', options: tileColorModeOptions(false) })
    ];
    if ((this.properties.tileColorMode || 'none') === 'custom') {
      const pickerOptions = tileColorPickerOptions();
      const count = Math.min(MAX_TILE_COLORS, tabNames.length);
      for (let i = 1; i <= count; i++) {
        tabColorFields.push(PropertyPaneDropdown('tileColor' + i, { label: tabNames[i - 1], options: pickerOptions }));
      }
    }

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
                })
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
                PropertyPaneTextField('tabField', { label: strings.TabFieldLabel, disabled: demo }),
                PropertyPaneTextField('subtitleField', { label: strings.SubtitleFieldLabel, disabled: demo }),
                PropertyPaneTextField('linkField', { label: strings.LinkFieldLabel, disabled: demo }),
                PropertyPaneLabel('fieldHint', { text: strings.FieldHint })
              ]
            },
            {
              groupName: 'Tab colours',
              groupFields: tabColorFields
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
