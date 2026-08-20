import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { SPPermission } from '@microsoft/sp-page-context';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import Header, { IHeaderProps } from './components/Header';
import { IHeaderLink, IHeaderSettings, IHeaderConfigPatch, parseUrl } from './models/IHeader';
import { getHeaderLinks } from './services/headerStore';
import { getHeaderConfig, saveHeaderConfig, IHeaderConfigRow } from './services/headerConfigStore';

export interface IIntranetHeaderProperties {
  brandText?: string;
  logoUrl?: string;       // optional image URL; empty = text brand
  homeUrl?: string;       // logo/brand link target; default = site home
  listTitle?: string;     // HeaderLinks list; default 'HeaderLinks'
  configList?: string;    // HeaderConfig list (owner-editable settings); default 'HeaderConfig'
  ctaText?: string;       // optional right-hand button label
  ctaUrl?: string;
  accent?: string;            // hex; matches the site theme accent
  hideNativeHeader?: boolean; // hide the built-in site header (title row + nav) so this band is the only site chrome (default true)
}

// The custom header provides the site title and navigation, so by default we hide SharePoint's
// built-in site header - its title row AND its horizontal nav - to avoid duplicates. This is
// DOM-based (Microsoft does not expose the site header as a placeholder), scoped to the site
// header region only. It does NOT touch the suite bar, the page command bar (New/Edit/Share),
// the left nav, or hub navigation. Reversible: removed on dispose.
const HEADER_HIDE_ID = 'intranet-header-hide-native-header';
// #spSiteHeader is the long-stable id of the modern site header (title row + nav). It is a
// sibling of the page command bar (#spCommandBar), so New/Edit/Share are untouched. The
// .ms-HorizontalNav fallback covers layouts where the nav renders outside the header id.
const HEADER_HIDE_CSS =
  '#spSiteHeader,' +
  '.ms-HorizontalNav{display:none !important;}';

// Site-wide intranet header: a full-width band in the Top placeholder, shown on every page.
// Brand (logo or text), an inline nav from the HeaderLinks list, and an optional CTA button.
// Deploy-script properties are the defaults; the HeaderConfig list (owner-editable in-browser,
// via the gear) overrides them. It never throws, so it can never break the page.
export default class IntranetHeaderApplicationCustomizer extends BaseApplicationCustomizer<IIntranetHeaderProperties> {

  private _top: PlaceholderContent | undefined = undefined;
  private _links: IHeaderLink[] = [];
  private _config: IHeaderConfigRow = { id: 0 };
  private _settings: IHeaderSettings = this._defaults();
  private _canManage: boolean = false;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._render);
    // re-render on client-side navigation so the active-link highlight tracks the current page
    this.context.application.navigatedEvent.add(this, this._render);

    // Hide the native header up front (based on the default) so it never flashes before the
    // async config loads. _applyNativeHide() reconciles once the config resolves (e.g. removes
    // the hide if an owner turned the header off).
    if (this.properties.hideNativeHeader !== false) { this._hideNativeHeader(); }

    this._canManage = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const linksList = this.properties.listTitle || 'HeaderLinks';
    const cfgList = this.properties.configList || 'HeaderConfig';
    const client = this.context.spHttpClient;

    return Promise.all([
      getHeaderLinks(client, webUrl, linksList),
      getHeaderConfig(client, webUrl, cfgList)
    ]).then((results) => {
      this._links = results[0];
      this._config = results[1];
      this._settings = this._resolve(this._config);
      this._applyNativeHide();
      this._render();
    });
  }

  private _defaults(): IHeaderSettings {
    return { enabled: true, hideNativeHeader: true, brandText: 'Contoso', logoUrl: '', homeUrl: '', showCta: false, ctaText: '', ctaUrl: '', accent: '#0f6cbd' };
  }

  // Layer the config row (wins when set) over the deploy-script properties (defaults).
  private _resolve(cfg: IHeaderConfigRow): IHeaderSettings {
    const p = this.properties;
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const pick = (a: string | undefined, b: string | undefined): string => (a !== undefined ? a : (b || ''));
    const ctaText = pick(cfg.ctaText, p.ctaText);
    const ctaUrl = parseUrl(cfg.ctaUrl !== undefined ? cfg.ctaUrl : p.ctaUrl);
    // Button visibility is an explicit flag. Default (no flag set yet): show it if a label and
    // link exist - so sites configured before this toggle existed keep their button.
    const showCta = cfg.showCta !== undefined ? cfg.showCta : (ctaText.length > 0 && ctaUrl.length > 0);
    return {
      enabled: cfg.enabled !== undefined ? cfg.enabled : true,
      hideNativeHeader: cfg.hideNativeHeader !== undefined ? cfg.hideNativeHeader : (p.hideNativeHeader !== false),
      brandText: pick(cfg.brandText, p.brandText) || 'Contoso',
      logoUrl: parseUrl(p.logoUrl),
      homeUrl: parseUrl(p.homeUrl) || webUrl,
      showCta: showCta,
      ctaText: ctaText,
      ctaUrl: ctaUrl,
      accent: pick(cfg.accent, p.accent) || '#0f6cbd'
    };
  }

  private _applyNativeHide(): void {
    if (this._settings.enabled && this._settings.hideNativeHeader) { this._hideNativeHeader(); }
    else { this._showNativeHeader(); }
  }

  private _hideNativeHeader(): void {
    if (document.getElementById(HEADER_HIDE_ID)) { return; }
    const style = document.createElement('style');
    style.id = HEADER_HIDE_ID;
    style.textContent = HEADER_HIDE_CSS;
    document.head.appendChild(style);
  }

  private _showNativeHeader(): void {
    const el = document.getElementById(HEADER_HIDE_ID);
    if (el && el.parentElement) { el.parentElement.removeChild(el); }
  }

  // Merge a saved patch onto the in-memory row, keeping the existing id. Used as a fallback so
  // the UI still reflects a save even if the confirming re-read fails - and, crucially, so the
  // row id never regresses to 0 (which would make the next save create a duplicate row).
  private _mergeConfig(cfg: IHeaderConfigRow, patch: IHeaderConfigPatch): IHeaderConfigRow {
    return {
      id: cfg.id,
      enabled: patch.enabled !== undefined ? patch.enabled : cfg.enabled,
      hideNativeHeader: patch.hideNativeHeader !== undefined ? patch.hideNativeHeader : cfg.hideNativeHeader,
      brandText: patch.brandText !== undefined ? patch.brandText : cfg.brandText,
      showCta: patch.showCta !== undefined ? patch.showCta : cfg.showCta,
      ctaText: patch.ctaText !== undefined ? patch.ctaText : cfg.ctaText,
      ctaUrl: patch.ctaUrl !== undefined ? patch.ctaUrl : cfg.ctaUrl,
      accent: patch.accent !== undefined ? patch.accent : cfg.accent
    };
  }

  // Persist an owner's change to the HeaderConfig list, then re-resolve + re-apply live.
  private _save = (patch: IHeaderConfigPatch): Promise<boolean> => {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const cfgList = this.properties.configList || 'HeaderConfig';
    const client = this.context.spHttpClient;
    return saveHeaderConfig(client, webUrl, cfgList, this._config.id, patch).then((ok) => {
      if (!ok) { return false; }
      // re-read to pick up the new row id (on first create) and the saved values; if the read
      // hiccups, fall back to a local merge so we keep a valid id and the change still shows.
      return getHeaderConfig(client, webUrl, cfgList).then((cfg) => {
        this._config = cfg.id > 0 ? cfg : this._mergeConfig(this._config, patch);
        this._settings = this._resolve(this._config);
        this._applyNativeHide();
        this._render();
        return true;
      });
    });
  };

  private _render = (): void => {
    if (!this._top) {
      this._top = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
    }
    if (!this._top || !this._top.domElement) { return; } // placeholder not ready yet

    const element: React.ReactElement<IHeaderProps> = React.createElement(Header, {
      settings: this._settings,
      links: this._links,   // already sorted by HeaderOrder
      currentPath: window.location.pathname,
      canManage: this._canManage,
      onSave: this._save
    });
    ReactDom.render(element, this._top.domElement);
  };

  private _onDispose = (): void => {
    if (this._top && this._top.domElement) {
      ReactDom.unmountComponentAtNode(this._top.domElement);
    }
  };

  protected onDispose(): void {
    this.context.placeholderProvider.changedEvent.remove(this, this._render);
    this.context.application.navigatedEvent.remove(this, this._render);
    this._showNativeHeader();
    this._onDispose();
  }
}
