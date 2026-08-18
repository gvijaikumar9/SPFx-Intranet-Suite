import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneDropdownOptionType,
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

// Tile fills, grouped by family. Soft/Fluorescent/Water are light (dark text stays
// legible); Dark fills flip the tile text to light automatically (see the component).
const MILD: { [key: string]: string } = {
  // soft pastels
  blue: '#eaf2fb', green: '#eaf6ee', amber: '#fdf3e3', rose: '#fbeef0',
  purple: '#f2edfb', teal: '#e8f5f3', grey: '#f1f3f5',
  // fluorescent (brighter, still dark-text readable)
  lime: '#ddf99a', coral: '#ffcdbf', sky: '#cfe9ff', lemon: '#fff0a6',
  lilac: '#e4d4ff', mint: '#bff4d6',
  // water / aqua
  aqua: '#d6f5f0', seafoam: '#dcf3e6', ocean: '#c2e6f0',
  // dark / deep (light text)
  slate: '#334155', indigo: '#3730a3', forest: '#14532d', maroon: '#7f1d1d',
  plum: '#4c1d95', charcoal: '#1f2937'
};
// Auto palettes rotate one family across the tiles, so each tile is a unique colour.
const PALETTES: { [mode: string]: string[] } = {
  palette: ['blue', 'green', 'amber', 'rose', 'purple', 'teal', 'grey'],   // soft (legacy key)
  fluro: ['lime', 'coral', 'sky', 'lemon', 'lilac', 'mint'],
  water: ['aqua', 'seafoam', 'ocean'],
  dark: ['slate', 'indigo', 'forest', 'maroon', 'plum', 'charcoal']
};
const MAX_TILE_COLORS = 8;

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
    const mode = this.properties.tileColorMode || 'none';
    const items = this._items || [];
    const props = this.properties as unknown as { [k: string]: string };
    const order = PALETTES[mode];
    const out: (string | undefined)[] = [];
    for (let i = 0; i < items.length; i++) {
      if (order) {
        out.push(MILD[order[i % order.length]]);
      } else if (mode === 'trend') {
        const t = items[i].trend;
        out.push(t === 'up' ? MILD.green : (t === 'down' ? MILD.rose : MILD.grey));
      } else if (mode === 'custom') {
        const key = (props['tileColor' + (i + 1)] || '').toString();
        out.push(key && MILD[key] ? MILD[key] : undefined);
      } else {
        out.push(undefined);
      }
    }
    return out;
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
    const H = PropertyPaneDropdownOptionType.Header;
    const colorOptions = [
      { key: '', text: 'Default (theme)' },
      { key: 'h-soft', text: 'Soft', type: H },
      { key: 'blue', text: 'Mild blue' },
      { key: 'green', text: 'Mild green' },
      { key: 'amber', text: 'Mild amber' },
      { key: 'rose', text: 'Mild rose' },
      { key: 'purple', text: 'Mild purple' },
      { key: 'teal', text: 'Mild teal' },
      { key: 'grey', text: 'Mild grey' },
      { key: 'h-fluro', text: 'Fluorescent', type: H },
      { key: 'lime', text: 'Lime' },
      { key: 'coral', text: 'Coral' },
      { key: 'sky', text: 'Sky' },
      { key: 'lemon', text: 'Lemon' },
      { key: 'lilac', text: 'Lilac' },
      { key: 'mint', text: 'Mint' },
      { key: 'h-water', text: 'Water', type: H },
      { key: 'aqua', text: 'Aqua' },
      { key: 'seafoam', text: 'Seafoam' },
      { key: 'ocean', text: 'Ocean' },
      { key: 'h-dark', text: 'Dark (light text)', type: H },
      { key: 'slate', text: 'Slate' },
      { key: 'indigo', text: 'Indigo' },
      { key: 'forest', text: 'Forest' },
      { key: 'maroon', text: 'Maroon' },
      { key: 'plum', text: 'Plum' },
      { key: 'charcoal', text: 'Charcoal' }
    ];
    const tileColorFields = [
      PropertyPaneDropdown('tileColorMode', {
        label: 'Tile colours',
        options: [
          { key: 'none', text: 'None (default)' },
          { key: 'h-auto', text: 'Auto palettes (unique per tile)', type: PropertyPaneDropdownOptionType.Header },
          { key: 'palette', text: 'Soft palette' },
          { key: 'fluro', text: 'Fluorescent palette' },
          { key: 'water', text: 'Water palette' },
          { key: 'dark', text: 'Dark palette (light text)' },
          { key: 'h-other', text: 'Other', type: PropertyPaneDropdownOptionType.Header },
          { key: 'trend', text: 'By trend (up green / down rose)' },
          { key: 'custom', text: 'Custom (choose each tile)' }
        ]
      })
    ];
    if ((this.properties.tileColorMode || 'none') === 'custom') {
      const tileCount = Math.min(MAX_TILE_COLORS, Math.max((this._items && this._items.length) || 0, 4));
      for (let i = 1; i <= tileCount; i++) {
        tileColorFields.push(PropertyPaneDropdown('tileColor' + i, { label: 'Tile ' + i, options: colorOptions }));
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
