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

import * as strings from 'KudosWebPartStrings';
import Kudos from './components/Kudos';
import { IKudosProps } from './components/IKudosProps';
import { IKudos, IPeopleResult } from './models/IKudos';
import { IStandardWebPartProps, computeFrameStyle, shouldShowTitle, standardPaneFields } from '../shared/standardProps';

export interface IKudosWebPartProps extends IStandardWebPartProps {
  title: string;
  layout: string;
  showPhoto: boolean;
  listTitle: string;
  maxItems: number;
  useDemoData: boolean;
}

const DEMO_KUDOS: IKudos[] = [
  { id: 1, headline: 'Saved the demo', message: 'Jumped in and fixed the build an hour before the client call.', toName: 'Priya Nair', fromName: 'Sam Okoro' },
  { id: 2, headline: 'Best onboarding', message: 'Made the new starter feel welcome from day one.', toName: 'Liam Chen', fromName: 'Maria Gomez' },
  { id: 3, headline: 'Went the extra mile', message: 'Stayed late to help QA reproduce a tricky bug.', toName: 'Ahmed Ali', fromName: 'Priya Nair' }
];

export default class KudosWebPart extends BaseClientSideWebPart<IKudosWebPartProps> {

  private _accent: string = '#0f6cbd';
  private _listOptions: { key: string; text: string }[] = [];
  private _items: IKudos[] = [];
  private _loading: boolean = true;
  private _error: string | undefined = undefined;
  private _submitting: boolean = false;
  private _submitError: string | undefined = undefined;
  private _submittedKey: number = 0;
  private _demoId: number = 1000;

  public render(): void {
    const isDemo = this.properties.useDemoData || !this.properties.listTitle;
    const element: React.ReactElement<IKudosProps> = React.createElement(
      Kudos,
      {
        title: this.properties.title || 'Kudos',
        layout: this.properties.layout || 'card',
        showTitle: shouldShowTitle(this.properties),
        frameStyle: computeFrameStyle(this.properties, this._accent),
        showPhoto: this.properties.showPhoto !== false,
        webUrl: this.context.pageContext.web.absoluteUrl,
        isDemo: isDemo,
        accent: this._accent,
        configured: isDemo || !!this.properties.listTitle,
        canGive: isDemo || !!this.properties.listTitle,
        items: this._items,
        loading: this._loading,
        error: this._error,
        submitting: this._submitting,
        submitError: this._submitError,
        submittedKey: this._submittedKey,
        onSearchPeople: (query: string) => this._searchPeople(query),
        onSubmit: (login: string | undefined, name: string, message: string) =>
          this._submitKudos(login, name, message)
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
      this._items = DEMO_KUDOS.slice(0, this.properties.maxItems || 6);
      this._loading = false;
      return;
    }

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const top = this.properties.maxItems || 6;
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`
      + `?$select=Id,Title,KudosMsg,ToPerson/Title,ToPerson/EMail,FromPerson/Title`
      + `&$expand=ToPerson,FromPerson&$top=${top}&$orderby=Created desc`;

    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._error = res.status === 404
          ? `List "${this.properties.listTitle}" was not found.`
          : `Could not load kudos (HTTP ${res.status}).`;
        this._items = [];
      } else {
        const json = await res.json();
        this._items = (json.value || []).map((row: Record<string, unknown>): IKudos => {
          const to = row.ToPerson as { Title?: string; EMail?: string } | undefined;
          const from = row.FromPerson as { Title?: string } | undefined;
          return {
            id: Number(row.Id ?? 0),
            headline: (row.Title ?? '').toString(),
            message: (row.KudosMsg ?? '').toString(),
            toName: to && to.Title ? to.Title : undefined,
            toEmail: to && to.EMail ? to.EMail : undefined,
            fromName: from && from.Title ? from.Title : undefined
          };
        });
      }
    } catch {
      this._error = 'Could not load kudos.';
      this._items = [];
    }

    this._loading = false;
  }

  /** Search the tenant's users for the people picker. */
  private async _searchPeople(query: string): Promise<IPeopleResult[]> {
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const body = JSON.stringify({
      queryParams: {
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        MaximumEntitySuggestions: 8,
        PrincipalSource: 15,
        PrincipalType: 1,          // User only
        QueryString: query
      }
    });
    try {
      const res: SPHttpClientResponse = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata', 'Content-type': 'application/json;odata=nometadata' },
        body: body
      });
      if (!res.ok) { return []; }
      const json = await res.json();
      const payload = json.value ?? (json.d && json.d.ClientPeoplePickerSearchUser);
      if (!payload) { return []; }
      const entities: Array<{ Key: string; DisplayText: string; EntityData?: { Email?: string } }> = JSON.parse(payload);
      return entities.map((e) => ({
        key: e.Key,
        text: e.DisplayText,
        secondaryText: e.EntityData && e.EntityData.Email ? e.EntityData.Email : undefined
      }));
    } catch {
      return [];
    }
  }

  /** Resolve a login name to a site user id via ensureuser. */
  private async _ensureUserId(login: string): Promise<number | undefined> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/ensureuser`;
    const res: SPHttpClientResponse = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata', 'Content-type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ logonName: login })
    });
    if (!res.ok) { return undefined; }
    const json = await res.json();
    return json && json.Id ? Number(json.Id) : undefined;
  }

  private _submitKudos(login: string | undefined, name: string, message: string): void {
    this._submitting = true;
    this._submitError = undefined;
    this._safeRender();

    if (this.properties.useDemoData || !this.properties.listTitle) {
      const id = ++this._demoId;
      const me = this.context.pageContext.user ? this.context.pageContext.user.displayName : 'You';
      this._items = [{ id: id, headline: `Kudos to ${name}`, message: message, toName: name, fromName: me }, ...this._items]
        .slice(0, this.properties.maxItems || 6);
      this._submitting = false;
      this._submittedKey += 1;
      this._safeRender();
      return;
    }

    this._createKudos(login, name, message).catch(() => {
      this._submitError = 'Could not send kudos. Please try again.';
      this._submitting = false;
      this._safeRender();
    });
  }

  private async _createKudos(login: string | undefined, name: string, message: string): Promise<void> {
    let toId: number | undefined;
    if (login) {
      toId = await this._ensureUserId(login);
      if (!toId) {
        this._submitError = 'We could not find that person. Please pick them again.';
        this._submitting = false;
        this._safeRender();
        return;
      }
    }
    const fromId = this.context.pageContext.legacyPageContext
      ? this.context.pageContext.legacyPageContext.userId
      : undefined;

    const listTitle = this.properties.listTitle.replace(/'/g, "''");
    const url = `${this.context.pageContext.web.absoluteUrl}`
      + `/_api/web/lists/getByTitle('${listTitle}')/items`;
    const payload: Record<string, unknown> = {
      Title: `Kudos to ${name}`,
      KudosMsg: message
    };
    if (toId) { payload.ToPersonId = toId; }
    if (fromId) { payload.FromPersonId = fromId; }

    const res: SPHttpClientResponse = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(payload)
    });
    if (!res.ok) {
      let detail = `HTTP ${res.status}`;
      try {
        const err = await res.json();
        if (err && err['odata.error'] && err['odata.error'].message) {
          detail = err['odata.error'].message.value || detail;
        }
      } catch { /* keep the status-code message */ }
      this._submitError = `Could not send kudos (${detail}).`;
      this._submitting = false;
      this._safeRender();
      return;
    }
    this._submitting = false;
    this._submittedKey += 1;
    await this._loadItems();
    this._safeRender();
  }

  private _safeRender(): void {
    if (this.domElement && this.properties) { this.render(); }
  }

  protected onPropertyPaneFieldChanged(path: string): void {
    if (path === 'backgroundMode') {
      this.context.propertyPane.refresh(); // enable/disable the custom-colour field
    }
    const dataProps = ['listTitle', 'maxItems', 'useDemoData'];
    if (dataProps.indexOf(path) >= 0) {
      if (path === 'useDemoData') {
        this.context.propertyPane.refresh();
      }
      this._submitError = undefined;
      this._loading = true;
      this._safeRender();
      this._loadItems().then(() => this._safeRender()).catch(() => this._safeRender());
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
    this.properties.title = 'Kudos';
    this.properties.layout = 'card';
    this.properties.showPhoto = true;
    this.properties.showTitle = true;
    this.properties.showBorder = false;
    this.properties.backgroundMode = 'transparent';
    this.properties.backgroundColor = '#eef3f8';
    this.properties.listTitle = '';
    this.properties.maxItems = 6;
    this.properties.useDemoData = true;
    this._submitError = undefined;
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
                PropertyPaneToggle('showPhoto', { label: strings.ShowPhotoLabel, onText: 'Profile photo', offText: 'Initials' }),
                PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 3, max: 15, step: 1 })
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
