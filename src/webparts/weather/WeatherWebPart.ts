import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WeatherWebPartStrings';
import Weather from './components/Weather';
import { IWeatherProps } from './components/IWeatherProps';
import { IWeather, IWeatherDay } from './models/IWeather';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IWeatherWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  location: string;
  units: string;
}

export default class WeatherWebPart extends BaseClientSideWebPart<IWeatherWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _weather: IWeather | undefined = undefined;
  private _loading: boolean = true;
  private _error: string | undefined = undefined;

  public render(): void {
    const isDemo = !(this.properties.location && this.properties.location.trim());
    const element: React.ReactElement<IWeatherProps> = React.createElement(
      Weather,
      {
        title: this.properties.title || 'Weather',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        isDemo: isDemo,
        accent: this._accent,
        weather: this._weather,
        loading: this._loading,
        error: this._error
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => { await this._loadWeather(); });
  }

  private _demo(): IWeather {
    const base = new Date().getTime();
    const day = (n: number): string => new Date(base + n * 86400000).toISOString().substring(0, 10);
    return {
      location: 'Contoso HQ',
      currentTemp: 21,
      currentCode: 2,
      unit: this.properties.units === 'F' ? 'F' : 'C',
      days: [
        { date: day(1), min: 14, max: 22, code: 1 },
        { date: day(2), min: 13, max: 20, code: 3 },
        { date: day(3), min: 15, max: 24, code: 0 }
      ]
    };
  }

  private async _loadWeather(): Promise<void> {
    this._loading = true;
    this._error = undefined;

    const loc = (this.properties.location || '').trim();
    if (!loc) { this._weather = this._demo(); this._loading = false; return; }

    try {
      const geoRes = await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(loc)}&count=1&language=en&format=json`);
      if (!geoRes.ok) { throw new Error('geo'); }
      const geo = await geoRes.json();
      if (!geo.results || geo.results.length === 0) {
        this._error = `Location "${loc}" was not found.`;
        this._weather = undefined;
        this._loading = false;
        return;
      }
      const g = geo.results[0];
      const unitParam = this.properties.units === 'F' ? 'fahrenheit' : 'celsius';
      const fRes = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${g.latitude}&longitude=${g.longitude}`
        + `&current=temperature_2m,weather_code&daily=temperature_2m_max,temperature_2m_min,weather_code`
        + `&timezone=auto&forecast_days=3&temperature_unit=${unitParam}`);
      if (!fRes.ok) { throw new Error('forecast'); }
      const f = await fRes.json();

      const daily = f.daily || {};
      const times = daily.time || [];
      const mins = daily.temperature_2m_min || [];
      const maxs = daily.temperature_2m_max || [];
      const codes = daily.weather_code || [];
      const days: IWeatherDay[] = [];
      for (let i = 0; i < times.length; i++) {
        const mn = Math.round(Number(mins[i]));
        const mx = Math.round(Number(maxs[i]));
        if (isNaN(mn) || isNaN(mx)) { continue; } // skip a partial/missing day
        const cd = Number(codes[i]);
        days.push({ date: times[i], min: mn, max: mx, code: isNaN(cd) ? 0 : cd });
      }
      const current = f.current || {};
      const curTemp = Math.round(Number(current.temperature_2m));
      if (isNaN(curTemp)) { throw new Error('bad-data'); } // degrade to the friendly error
      const curCode = Number(current.weather_code);
      this._weather = {
        location: g.name + (g.country_code ? `, ${g.country_code}` : ''),
        currentTemp: curTemp,
        currentCode: isNaN(curCode) ? 0 : curCode,
        unit: this.properties.units === 'F' ? 'F' : 'C',
        days: days
      };
    } catch {
      this._error = 'Could not load the weather right now.';
      this._weather = undefined;
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
    if (path === 'location' || path === 'units') {
      this._loading = true;
      this._safeRender();
      this._loadWeather().then(() => this._safeRender()).catch(() => this._safeRender());
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
    this.properties.title = 'Weather';
    this.properties.layout = 'card';
    this.properties.location = '';
    this.properties.units = 'C';
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.context.propertyPane.refresh();
    this._loading = true;
    this._safeRender();
    this._loadWeather().then(() => this._safeRender()).catch(() => this._safeRender());
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
                })
              ]
            },
            {
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('location', { label: strings.LocationFieldLabel, placeholder: 'e.g. London' }),
                PropertyPaneDropdown('units', {
                  label: strings.UnitsFieldLabel,
                  options: [
                    { key: 'C', text: 'Celsius (°C)' },
                    { key: 'F', text: 'Fahrenheit (°F)' }
                  ]
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
