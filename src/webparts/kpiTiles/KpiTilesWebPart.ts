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

import * as strings from 'KpiTilesWebPartStrings';
import KpiTiles from './components/KpiTiles';
import { IKpiTilesProps } from './components/IKpiTilesProps';
import { IKpi, normalizeTrend } from './models/IKpi';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';
import { resolveTileColors, collectCustomKeys, tileColorModeOptions, tileColorPickerOptions, MAX_TILE_COLORS } from '../shared/tileColors';

export interface IKpiTilesWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  columns: number;
  listTitle: string;
  useDemoData: boolean;
  tileColorMode: string;   // none | palette | trend | custom
  tileColor1: string;
  tileColor2: string;
  tileColor3: string;
  tileColor4: string;
  tileColor5: string;
  tileColor6: string;
  tileColor7: string;
  tileColor8: string;
}

const DEMO_ITEMS: IKpi[] = [
  { id: 1, label: 'Open IT tickets', value: '42', trend: 'down', delta: '12% vs last week' },
  { id: 2, label: 'Employees', value: '1,204', trend: 'up', delta: '18 new hires' },
  { id: 3, label: 'Platform uptime', value: '99.98%', trend: 'up', delta: 'above target' },
  { id: 4, label: 'Engagement', value: '72', trend: 'up', delta: '+4 pts' }
];

export default class KpiTilesWebPart extends BaseClientSideWebPart<IKpiTilesWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: IKpi[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IKpiTilesProps> = React.createElement(
      KpiTiles,
      {
        title: this.properties.title || 'Key numbers',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        columns: this.properties.columns || 4,
        isDemo: isDemo,
        accent: this._accent,
        items: this._items,
        tileColors: this._tileColors(),
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
      + `?$select=Id,Title,KpiValue,KpiTrend,KpiDelta&$top=20`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load numbers (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>): IKpi => ({
          id: Number(row.Id ?? 0),
          label: (row.Title ?? '').toString(),
          value: (row.KpiValue ?? '').toString(),
          trend: normalizeTrend(row.KpiTrend),
          delta: row.KpiDelta ? row.KpiDelta.toString() : undefined
        })).filter((k: IKpi) => k.label.length > 0);
      }
    } catch {
      this._error = 'Could not load numbers.';
      this._items = [];
    }

    this._loading = false;
  }

  private _tileColors(): (string | undefined)[] {
    const items = this._items || [];
    const trends: string[] = [];
    for (let i = 0; i < items.length; i++) { trends.push(items[i].trend); }
    return resolveTileColors(
      this.properties.tileColorMode || 'none',
      items.length,
      collectCustomKeys(this.properties, items.length),
      trends
    );
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
    if (path === 'tileColorMode') {
      this.context.propertyPane.refresh(); // show or hide the per-tile pickers
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
    this.properties.title = 'Key numbers';
    this.properties.layout = 'card';
    this.properties.columns = 4;
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
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

    // Tile colours group: a mode dropdown, plus one picker per tile in Custom mode.
    const tileColorFields = [
      PropertyPaneDropdown('tileColorMode', { label: 'Tile colours', options: tileColorModeOptions(true) })
    ];
    if ((this.properties.tileColorMode || 'none') === 'custom') {
      const tileCount = Math.min(MAX_TILE_COLORS, Math.max((this._items && this._items.length) || 0, 4));
      const pickerOptions = tileColorPickerOptions();
      for (let i = 1; i <= tileCount; i++) {
        tileColorFields.push(PropertyPaneDropdown('tileColor' + i, { label: 'Tile ' + i, options: pickerOptions }));
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
              groupName: 'Tile colours',
              groupFields: tileColorFields
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
