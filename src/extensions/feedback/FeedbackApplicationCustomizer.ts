import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import FeedbackWidget, { IFeedbackWidgetProps } from './components/FeedbackWidget';
import { submitFeedback } from './services/feedbackStore';

export interface IFeedbackProperties {
  listTitle?: string;   // Feedback list; default 'Feedback'
  accent?: string;      // hex; matches the site theme accent
}

const HOST_ID = 'intranet-suite-feedback';

// Site-wide feedback widget: a fixed, always-visible bubble in the bottom-right corner of
// every page. Click it to open a panel (rating + comment + optional follow-up); submits to
// a Feedback list, capturing the page URL. Simple - it is just a fixed overlay mounted on
// <body>, no scroll logic. Never breaks the page.
export default class FeedbackApplicationCustomizer extends BaseApplicationCustomizer<IFeedbackProperties> {

  private _host: HTMLDivElement | undefined = undefined;

  public onInit(): Promise<void> {
    let host = document.getElementById(HOST_ID) as HTMLDivElement | null;
    if (!host) {
      host = document.createElement('div');
      host.id = HOST_ID;
      document.body.appendChild(host);
    }
    this._host = host;

    const list = this.properties.listTitle || 'Feedback';
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const user = this.context.pageContext.user;
    const userName = user ? (user.displayName || '') : '';
    // Only a real email is actionable. loginName is a claims token (i:0#.f|membership|...),
    // so never use it as a contact address.
    const email = user && user.email ? user.email : '';
    // stored on opt-in: "Name <email>", or whichever of the two is available
    const userContact = email ? (userName ? `${userName} <${email}>` : email) : userName;

    const element: React.ReactElement<IFeedbackWidgetProps> = React.createElement(FeedbackWidget, {
      accent: this.properties.accent || '#0f6cbd',
      userName: userName,
      userContact: userContact,
      onSubmit: (rating: number, comment: string, contact: string): Promise<boolean> =>
        submitFeedback(this.context.spHttpClient, webUrl, list, {
          rating: rating,
          comment: comment,
          pageUrl: window.location.href,
          contact: contact
        })
    });
    ReactDom.render(element, host);
    return Promise.resolve();
  }

  protected onDispose(): void {
    if (this._host) {
      ReactDom.unmountComponentAtNode(this._host);
      if (this._host.parentElement) { this._host.parentElement.removeChild(this._host); }
      this._host = undefined;
    }
  }
}
