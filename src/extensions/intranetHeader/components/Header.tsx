import * as React from 'react';
import { useState } from 'react';
import styles from './Header.module.scss';
import { IHeaderLink, IHeaderSettings, IHeaderConfigPatch, isActive } from '../models/IHeader';
import HeaderSettings from './HeaderSettings';

export interface IHeaderProps {
  settings: IHeaderSettings;   // resolved effective settings (config over defaults)
  links: IHeaderLink[];        // flat, already in HeaderOrder
  currentPath: string;         // window.location.pathname, for active-link highlight
  canManage: boolean;          // true = site manager, shows the gear + settings panel
  onSave: (patch: IHeaderConfigPatch) => Promise<boolean>;
}

// A slim, full-width header band shown at the top of every page. Brand (logo or text) on the
// left, an inline nav in the middle, and an optional call-to-action button on the right. Site
// managers also get a gear that opens the settings panel. On narrow screens the nav collapses.
const Header: React.FC<IHeaderProps> = (props) => {
  const { settings, links, currentPath, canManage, onSave } = props;
  const { brandText, logoUrl, homeUrl, showCta, ctaText, ctaUrl, accent, enabled } = settings;
  const rootStyle = { ['--accent' as string]: accent } as React.CSSProperties;
  const [menuOpen, setMenuOpen] = useState<boolean>(false);
  const [settingsOpen, setSettingsOpen] = useState<boolean>(false);
  const hasNav = enabled && links.length > 0;

  const gear = (
    <button className={styles.gear} type="button" aria-label="Header settings" onClick={() => setSettingsOpen(true)}>
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
        <circle cx="12" cy="12" r="3" />
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
      </svg>
    </button>
  );

  const panel = settingsOpen && (
    <HeaderSettings settings={settings} accent={accent} onSave={onSave} onClose={() => setSettingsOpen(false)} />
  );

  // Disabled: the header is off. Regular users see nothing; managers see a slim strip with a
  // gear so they can turn it back on from the browser (no PowerShell needed).
  if (!enabled) {
    if (!canManage) { return null; }
    return (
      <div className={`${styles.header} ${styles.off}`} style={rootStyle} role="banner">
        <div className={styles.inner}>
          <span className={styles.offNote}>Custom header is off</span>
          <div className={styles.right}>{gear}</div>
        </div>
        {panel}
      </div>
    );
  }

  const nav = (extraClass: string): React.ReactElement => (
    <nav className={`${styles.links} ${extraClass}`} aria-label="Primary">
      {links.map((l, i) => {
        const active = isActive(l.url, currentPath);
        return (
          <a
            key={i}
            className={`${styles.link} ${active ? styles.active : ''}`}
            href={l.url}
            aria-current={active ? 'page' : undefined}
          >{l.title}</a>
        );
      })}
    </nav>
  );

  return (
    <div className={styles.header} style={rootStyle} role="banner">
      <div className={styles.inner}>
        <a className={styles.brand} href={homeUrl || '/'}>
          {logoUrl
            ? <img className={styles.logo} src={logoUrl} alt={brandText || 'Home'} />
            : <span className={styles.brandText}>{brandText || 'Contoso'}</span>}
          {logoUrl && brandText && <span className={styles.brandText}>{brandText}</span>}
        </a>

        {hasNav && nav(styles.desktop)}

        <div className={styles.right}>
          {showCta && ctaText && ctaUrl && (
            <a className={styles.cta} href={ctaUrl}>{ctaText}</a>
          )}
          {canManage && gear}
          {hasNav && (
            <button
              className={styles.burger}
              type="button"
              aria-label="Toggle menu"
              aria-expanded={menuOpen}
              onClick={() => setMenuOpen(!menuOpen)}
            >
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
                <line x1="3" y1="6" x2="21" y2="6" />
                <line x1="3" y1="12" x2="21" y2="12" />
                <line x1="3" y1="18" x2="21" y2="18" />
              </svg>
            </button>
          )}
        </div>
      </div>

      {hasNav && menuOpen && (
        <div className={styles.drawer}>{nav(styles.stack)}</div>
      )}

      {panel}
    </div>
  );
};

export default Header;
