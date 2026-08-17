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

import * as strings from 'GreetingWebPartStrings';
import Greeting from './components/Greeting';
import { IGreetingProps, IChip } from './components/IGreetingProps';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IGreetingWebPartProps extends IStandardWebPartProps {
  layout: string;
  chipsText: string;
}

export default class GreetingWebPart extends BaseClientSideWebPart<IGreetingWebPartProps> {

  private _accent: string = '#0f6cbd';

  private _firstName(): string {
    const dn = this.context.pageContext.user ? this.context.pageContext.user.displayName : '';
    return (dn || '').split(' ')[0] || 'there';
  }

  private _chips(): IChip[] {
    const raw = this.properties.chipsText || '';
    const parts = raw.split(',');
    const chips: IChip[] = [];
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i].trim();
      if (!p) { continue; }
      const bar = p.indexOf('|');
      if (bar >= 0) {
        const count = p.substring(0, bar).trim();
        const label = p.substring(bar + 1).trim();
        chips.push({ count: count || undefined, label: label });
      } else {
        chips.push({ label: p });
      }
    }
    return chips;
  }

  public render(): void {
    const element: React.ReactElement<IGreetingProps> = React.createElement(
      Greeting,
      {
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        name: this._firstName(),
        chips: this._chips(),
        accent: this._accent
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
    this.properties.layout = 'card';
    this.properties.chipsText = '3|Tasks due, 2|Approvals, |Next: Standup 10:30';
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
                PropertyPaneTextField('chipsText', { label: strings.ChipsFieldLabel, multiline: true }),
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
