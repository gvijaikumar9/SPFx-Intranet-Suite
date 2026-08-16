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

import * as strings from 'PollSurveyWebPartStrings';
import PollSurvey from './components/PollSurvey';
import { IPollSurveyProps, IPollOption } from './components/IPollSurveyProps';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IPollSurveyWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  question: string;
  optionsText: string;
  listTitle: string;
  useDemoData: boolean;
}

export default class PollSurveyWebPart extends BaseClientSideWebPart<IPollSurveyWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _results: IPollOption[] = [];
  private _total: number = 0;
  private _voted: boolean = false;
  private _loading: boolean = true;
  private _error: string | undefined = undefined;
  private _submitting: boolean = false;
  private _submitError: string | undefined = undefined;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const options = this._parseOptions();
    const element: React.ReactElement<IPollSurveyProps> = React.createElement(
      PollSurvey,
      {
        title: this.properties.title || 'Poll of the week',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        question: this.properties.question || 'What should we improve first?',
        options: options,
        results: this._results,
        total: this._total,
        voted: this._voted,
        configured: isDemo || !!this.properties.listTitle,
        isDemo: isDemo,
        accent: this._accent,
        submitting: this._submitting,
        submitError: this._submitError,
        loading: this._loading,
        error: this._error,
        onVote: (option: string) => this._vote(option)
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
      if (!(this.properties.useDemoData || !this.properties.listTitle)) {
        this._voted = this._readVoted();
      }
      await this._loadListOptions().catch(() => { /* dropdown stays empty; not fatal */ });
      await this._loadResults();
    });
  }

  private _parseOptions(): string[] {
    const raw = this.properties.optionsText || 'Cafeteria, Parking, Meeting rooms, Remote tools';
    const list = raw.split(',').map((o) => o.trim()).filter((o) => o.length > 0);
    return list.length > 0 ? list : ['Cafeteria', 'Parking', 'Meeting rooms', 'Remote tools'];
  }

  private _votedKey(): string {
    return `intranetsuite-poll-${this.context.instanceId}-${this.properties.listTitle || 'demo'}`;
  }

  private _readVoted(): boolean {
    try { return window.localStorage.getItem(this._votedKey()) === '1'; }
    catch { return false; }
  }

  private _writeVoted(): void {
    try { window.localStorage.setItem(this._votedKey(), '1'); }
    catch { /* private mode or blocked; voted stays in-memory */ }
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

  private async _loadResults(): Promise<void> {
    this._loading = true;
    this._error = undefined;
    const options = this._parseOptions();

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._results = options.map((o, i) => ({ option: o, count: Math.max(1, 30 - i * 8) }));
      this._total = this._results.reduce((s, r) => s + r.count, 0);
      this._loading = false;
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items?$select=Title&$top=5000`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load votes (HTTP ${res.status}).`;
        this._results = [];
        this._total = 0;
      } else {
        const json = await res.json();
        const counts: Record<string, number> = {};
        (json.value || []).forEach((row: Record<string, unknown>) => {
          const t = (row.Title ?? '').toString();
          if (t) { counts[t] = (counts[t] || 0) + 1; }
        });
        this._results = options.map((o) => ({ option: o, count: counts[o] || 0 }));
        this._total = this._results.reduce((s, r) => s + r.count, 0);
      }
    } catch {
      this._error = 'Could not load votes.';
      this._results = [];
      this._total = 0;
    }

    this._loading = false;
  }

  private _vote(option: string): void {
    if (this._voted || this._submitting) { return; }
    this._submitting = true;
    this._submitError = undefined;
    this._safeRender();

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._results = this._results.map((r) => r.option === option ? { option: r.option, count: r.count + 1 } : r);
      this._total += 1;
      this._voted = true;
      this._submitting = false;
      this._safeRender();
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`;
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({ Title: option })
    }).then(async (res: SPHttpClientResponse) => {
      if (!res.ok) {
        this._submitError = `Could not save your vote (HTTP ${res.status}).`;
        this._submitting = false;
        this._safeRender();
        return;
      }
      this._voted = true;
      this._writeVoted();
      this._submitting = false;
      await this._loadResults();
      this._safeRender();
    }).catch(() => {
      this._submitError = 'Could not save your vote. Please try again.';
      this._submitting = false;
      this._safeRender();
    });
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh();
    }
    const dataProps = ['listTitle', 'optionsText', 'useDemoData'];
    if (dataProps.indexOf(path) >= 0) {
      if (path === 'useDemoData') {
        this.context.propertyPane.refresh();
        this._voted = (path === 'useDemoData' && !this.properties.useDemoData) ? this._readVoted() : false;
      }
      this._loading = true;
      this._safeRender();
      this._loadResults().then(() => this._safeRender()).catch(() => this._safeRender());
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
    this.properties.title = 'Poll of the week';
    this.properties.layout = 'card';
    this.properties.question = 'What should we improve first?';
    this.properties.optionsText = 'Cafeteria, Parking, Meeting rooms, Remote tools';
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.useDemoData = true;
    this._voted = false;
    this._submitError = undefined;
    this.context.propertyPane.refresh();
    this._loading = true;
    this._safeRender();
    this._loadResults().then(() => this._safeRender()).catch(() => this._safeRender());
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
                })
              ]
            },
            {
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('question', { label: strings.QuestionFieldLabel }),
                PropertyPaneTextField('optionsText', { label: strings.OptionsFieldLabel })
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
