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

import * as strings from 'EventCountdownWebPartStrings';
import EventCountdown from './components/EventCountdown';
import { IEventCountdownProps } from './components/IEventCountdownProps';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';
import { resolveTileColors, collectCustomKeys, tileColorModeOptions, tileColorPickerOptions } from '../shared/tileColors';

export interface IEventCountdownWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  eventName: string;
  targetDate: string;
  tileColorMode: string;   // none | palette | fluro | water | dark | custom
  tileColor1: string;
  tileColor2: string;
  tileColor3: string;
  tileColor4: string;
}

const UNIT_LABELS = ['Days', 'Hours', 'Mins', 'Secs'];

export default class EventCountdownWebPart extends BaseClientSideWebPart<IEventCountdownWebPartProps> {

  private _accent: string = '#0f6cbd';

  public render(): void {
    const hasTarget = !!(this.properties.targetDate && this.properties.targetDate.trim());
    // With no target set, count down to ten days out so the web part is alive on drop.
    const targetIso = hasTarget
      ? this.properties.targetDate
      : new Date(new Date().getTime() + 10 * 86400000).toISOString();

    const element: React.ReactElement<IEventCountdownProps> = React.createElement(
      EventCountdown,
      {
        title: this.properties.title || 'Counting down to',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        eventName: this.properties.eventName || 'Company kickoff',
        targetIso: targetIso,
        isDemo: !hasTarget,
        accent: this._accent,
        unitColors: resolveTileColors(this.properties.tileColorMode || 'none', 4, collectCustomKeys(this.properties, 4))
      }
    );
    ReactDom.render(element, this.domElement);
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh();
    }
    if (path === 'tileColorMode') {
      this.context.propertyPane.refresh(); // show or hide the per-unit pickers
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
    this.properties.title = 'Counting down to';
    this.properties.layout = 'card';
    this.properties.eventName = 'Company kickoff';
    this.properties.targetDate = '';
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.tileColorMode = 'none';
    const p = this.properties as unknown as { [k: string]: string };
    for (let i = 1; i <= 4; i++) { p['tileColor' + i] = ''; }
    this.context.propertyPane.refresh();
    this._safeRender();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Unit colours group: a mode dropdown, plus a picker per unit box in Custom mode.
    const unitColorFields = [
      PropertyPaneDropdown('tileColorMode', { label: 'Unit colours', options: tileColorModeOptions(false) })
    ];
    if ((this.properties.tileColorMode || 'none') === 'custom') {
      const pickerOptions = tileColorPickerOptions();
      for (let i = 1; i <= 4; i++) {
        unitColorFields.push(PropertyPaneDropdown('tileColor' + i, { label: UNIT_LABELS[i - 1], options: pickerOptions }));
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
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('eventName', { label: strings.EventNameFieldLabel }),
                PropertyPaneTextField('targetDate', { label: strings.TargetDateFieldLabel }),
                PropertyPaneLabel('fieldHint', { text: strings.FieldHint })
              ]
            },
            {
              groupName: 'Unit colours',
              groupFields: unitColorFields
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
