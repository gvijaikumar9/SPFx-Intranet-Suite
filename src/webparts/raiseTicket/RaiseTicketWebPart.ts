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

import * as strings from 'RaiseTicketWebPartStrings';
import RaiseTicket from './components/RaiseTicket';
import { IRaiseTicketProps } from './components/IRaiseTicketProps';
import { ITicket } from './models/ITicket';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IRaiseTicketWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  listTitle: string;
  categoriesText: string;
  useDemoData: boolean;
}

const DEMO_RECENT: ITicket[] = [
  { id: 4821, title: 'Laptop will not charge', category: 'IT', status: 'In progress' },
  { id: 4790, title: 'New monitor request', category: 'Facilities', status: 'Resolved' }
];

export default class RaiseTicketWebPart extends BaseClientSideWebPart<IRaiseTicketWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _recent: ITicket[] = [];
  private _loadingRecent: boolean = true;
  private _recentError: string | undefined = undefined;
  private _submitting: boolean = false;
  private _submitError: string | undefined = undefined;
  private _submittedId: number | undefined = undefined;
  private _demoCounter: number = 4900;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const categories = this._parseCategories();
    const element: React.ReactElement<IRaiseTicketProps> = React.createElement(
      RaiseTicket,
      {
        title: this.properties.title || 'Raise a ticket',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        categories: categories,
        isDemo: isDemo,
        accent: this._accent,
        configured: isDemo || !!this.properties.listTitle,
        submitting: this._submitting,
        submitError: this._submitError,
        submittedId: this._submittedId,
        recent: this._recent,
        loadingRecent: this._loadingRecent,
        recentError: this._recentError,
        onSubmit: (subject: string, category: string, description: string) =>
          this._submitTicket(subject, category, description)
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
      await this._loadListOptions().catch(() => { /* dropdown stays empty; not fatal */ });
      await this._loadRecent();
    });
  }

  private _parseCategories(): string[] {
    const raw = this.properties.categoriesText || 'IT, HR, Facilities, Other';
    const list = raw.split(',').map((c) => c.trim()).filter((c) => c.length > 0);
    return list.length > 0 ? list : ['IT', 'HR', 'Facilities', 'Other'];
  }

  /** Populate the property-pane list dropdown with the site's non-hidden custom lists. */
  private async _loadListOptions(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100`
      + `&$select=Title&$orderby=Title`;
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return; }
    const json = await res.json();
    this._listOptions = (json.value || []).map((l: { Title: string }) => ({ key: l.Title, text: l.Title }));
  }

  /** Load the current user's most recent tickets, or demo rows. */
  private async _loadRecent(): Promise<void> {
    this._loadingRecent = true;
    this._recentError = undefined;

    if (this.properties.useDemoData || !this.properties.listTitle) {
      this._recent = DEMO_RECENT;
      this._loadingRecent = false;
      return;
    }

    const userId = this.context.pageContext.legacyPageContext
      ? this.context.pageContext.legacyPageContext.userId
      : undefined;
    if (!userId) {
      // Without the current user's id we cannot scope to "my" tickets, so show none
      // rather than leaking everyone's tickets under a "My recent tickets" heading.
      this._recent = [];
      this._loadingRecent = false;
      return;
    }
    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=Id,Title,Category,TicketStatus,Created&$filter=AuthorId eq ${userId}&$orderby=Created desc&$top=5`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._recentError = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load recent tickets (HTTP ${res.status}).`;
        this._recent = [];
      } else {
        const json = await res.json();
        this._recent = (json.value || []).map((row: Record<string, unknown>): ITicket => ({
          id: Number(row.Id ?? 0),
          title: (row.Title ?? '(no subject)').toString(),
          category: row.Category ? row.Category.toString() : undefined,
          status: row.TicketStatus ? row.TicketStatus.toString() : 'New'
        }));
      }
    } catch {
      this._recentError = 'Could not load recent tickets.';
      this._recent = [];
    }

    this._loadingRecent = false;
  }

  /** Create a ticket. In demo mode this only updates the on-screen list. */
  private _submitTicket(subject: string, category: string, description: string): void {
    this._submitting = true;
    this._submitError = undefined;
    this._safeRender();

    if (this.properties.useDemoData || !this.properties.listTitle) {
      const id = ++this._demoCounter;
      this._recent = [{ id: id, title: subject, category: category, status: 'New' }, ...this._recent].slice(0, 5);
      this._submitting = false;
      this._submittedId = id;
      this._safeRender();
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`;
    const body = JSON.stringify({
      Title: subject,
      Category: category,
      TicketDesc: description,
      TicketStatus: 'New'
    });

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: body
    }).then(async (res: SPHttpClientResponse) => {
      if (!res.ok) {
        let detail = `HTTP ${res.status}`;
        try {
          const err = await res.json();
          if (err && err['odata.error'] && err['odata.error'].message) {
            detail = err['odata.error'].message.value || detail;
          }
        } catch { /* keep the status-code message */ }
        this._submitError = `Could not submit your ticket (${detail}).`;
        this._submitting = false;
        this._safeRender();
        return;
      }
      const created = await res.json();
      this._submittedId = created && created.Id ? Number(created.Id) : 0;
      this._submitting = false;
      await this._loadRecent();
      this._safeRender();
    }).catch(() => {
      this._submitError = 'Could not submit your ticket. Please try again.';
      this._submitting = false;
      this._safeRender();
    });
  }

  /**
   * Render only when BOTH the DOM element and the properties bag exist. onThemeChanged
   * fires during _internalInitialize, before this.properties is set.
   */
  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh(); // enable/disable the custom-colour field
    }
    const dataProps = ['listTitle', 'useDemoData'];
    if (dataProps.indexOf(path) >= 0) {
      if (path === 'useDemoData') {
        this.context.propertyPane.refresh();
      }
      this._submittedId = undefined;
      this._submitError = undefined;
      this._loadingRecent = true;
      this._safeRender();
      this._loadRecent().then(() => this._safeRender()).catch(() => this._safeRender());
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

  /** Restore every property to its shipped default and reload. */
  private _resetToDefault(): void {
    this.properties.title = 'Raise a ticket';
    this.properties.layout = 'card';
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.categoriesText = 'IT, HR, Facilities, Other';
    this.properties.useDemoData = true;
    this._submitError = undefined;
    this._submittedId = undefined;
    this.context.propertyPane.refresh();
    this._loadingRecent = true;
    this._safeRender();
    this._loadRecent().then(() => this._safeRender()).catch(() => this._safeRender());
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
                }),
                PropertyPaneTextField('categoriesText', { label: strings.CategoriesFieldLabel })
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
