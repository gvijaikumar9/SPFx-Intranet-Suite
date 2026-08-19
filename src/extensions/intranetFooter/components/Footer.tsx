import * as React from 'react';
import styles from './Footer.module.scss';
import { IFooterGroup, ISocialLink } from '../models/IFooter';

export interface IFooterProps {
  brandText: string;
  blurb: string;
  copyright: string;
  groups: IFooterGroup[];
  social: ISocialLink[];
  accent: string;
}

// Minimal inline icon set (stroke icons scale with currentColor).
const ICONS: { [k: string]: React.ReactElement } = {
  linkedin: <path d="M4.98 3.5a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5zM3 9h4v12H3zM9 9h3.8v1.7h.1c.5-1 1.8-2 3.7-2 4 0 4.7 2.6 4.7 6V21h-4v-5.3c0-1.3 0-3-1.8-3s-2 1.4-2 2.9V21H9z" />,
  twitter: <path d="M18.9 3H22l-7 8 8.2 10h-6.4l-5-6.1L6 21H3l7.5-8.6L2.6 3H9l4.5 5.6zM17.8 19h1.7L7.3 4.8H5.5z" />,
  x: <path d="M18.9 3H22l-7 8 8.2 10h-6.4l-5-6.1L6 21H3l7.5-8.6L2.6 3H9l4.5 5.6zM17.8 19h1.7L7.3 4.8H5.5z" />,
  facebook: <path d="M22 12a10 10 0 1 0-11.6 9.9v-7H7.9V12h2.5V9.8c0-2.5 1.5-3.9 3.8-3.9 1.1 0 2.2.2 2.2.2v2.5h-1.2c-1.2 0-1.6.8-1.6 1.6V12h2.7l-.4 2.9h-2.3v7A10 10 0 0 0 22 12z" />,
  youtube: <path d="M23 12s0-3-.4-4.4a2.6 2.6 0 0 0-1.8-1.8C19.3 5.4 12 5.4 12 5.4s-7.3 0-8.8.4a2.6 2.6 0 0 0-1.8 1.8C1 9 1 12 1 12s0 3 .4 4.4a2.6 2.6 0 0 0 1.8 1.8c1.5.4 8.8.4 8.8.4s7.3 0 8.8-.4a2.6 2.6 0 0 0 1.8-1.8C23 15 23 12 23 12zM9.8 15V9l5.2 3z" />,
  instagram: <path d="M12 7.5A4.5 4.5 0 1 0 16.5 12 4.5 4.5 0 0 0 12 7.5zm0 7.4A2.9 2.9 0 1 1 14.9 12 2.9 2.9 0 0 1 12 14.9zM16.8 6a1.1 1.1 0 1 0 1.1 1.1A1.1 1.1 0 0 0 16.8 6zM12 4.6c2.4 0 2.7 0 3.6.1a5 5 0 0 1 1.7.3 3 3 0 0 1 1.7 1.7 5 5 0 0 1 .3 1.7c0 .9.1 1.2.1 3.6s0 2.7-.1 3.6a5 5 0 0 1-.3 1.7 3 3 0 0 1-1.7 1.7 5 5 0 0 1-1.7.3c-.9 0-1.2.1-3.6.1s-2.7 0-3.6-.1a5 5 0 0 1-1.7-.3 3 3 0 0 1-1.7-1.7 5 5 0 0 1-.3-1.7c0-.9-.1-1.2-.1-3.6s0-2.7.1-3.6a5 5 0 0 1 .3-1.7A3 3 0 0 1 6.7 5a5 5 0 0 1 1.7-.3c.9 0 1.2-.1 3.6-.1M12 3c-2.4 0-2.8 0-3.7.1a6.6 6.6 0 0 0-2.2.4 4.6 4.6 0 0 0-2.6 2.6 6.6 6.6 0 0 0-.4 2.2C3 9.2 3 9.6 3 12s0 2.8.1 3.7a6.6 6.6 0 0 0 .4 2.2 4.6 4.6 0 0 0 2.6 2.6 6.6 6.6 0 0 0 2.2.4c.9.1 1.3.1 3.7.1s2.8 0 3.7-.1a6.6 6.6 0 0 0 2.2-.4 4.6 4.6 0 0 0 2.6-2.6 6.6 6.6 0 0 0 .4-2.2c.1-.9.1-1.3.1-3.7s0-2.8-.1-3.7a6.6 6.6 0 0 0-.4-2.2 4.6 4.6 0 0 0-2.6-2.6 6.6 6.6 0 0 0-2.2-.4C14.8 3 14.4 3 12 3z" />,
  github: <path d="M12 2a10 10 0 0 0-3.2 19.5c.5.1.7-.2.7-.5v-1.7c-2.8.6-3.4-1.3-3.4-1.3-.5-1.2-1.1-1.5-1.1-1.5-.9-.6.1-.6.1-.6 1 .1 1.5 1 1.5 1 .9 1.5 2.3 1.1 2.9.8.1-.7.3-1.1.6-1.4-2.2-.3-4.6-1.1-4.6-4.9 0-1.1.4-2 1-2.7-.1-.3-.4-1.3.1-2.7 0 0 .8-.3 2.7 1a9.4 9.4 0 0 1 5 0c1.9-1.3 2.7-1 2.7-1 .5 1.4.2 2.4.1 2.7.6.7 1 1.6 1 2.7 0 3.8-2.3 4.6-4.6 4.9.4.3.7.9.7 1.9v2.7c0 .3.2.6.7.5A10 10 0 0 0 12 2z" />,
  email: <path d="M3 5h18a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1zm9 7L4.5 7h15z" />,
  web: <path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm6.9 6h-2.6a15.6 15.6 0 0 0-1.2-3.3A8 8 0 0 1 18.9 8zM12 4c.8 1 1.4 2.3 1.8 4h-3.6C10.6 6.3 11.2 5 12 4zM4.3 14a7.8 7.8 0 0 1 0-4h2.9a17.6 17.6 0 0 0 0 4zm.8 2h2.6a15.6 15.6 0 0 0 1.2 3.3A8 8 0 0 1 5.1 16zm2.6-8H5.1a8 8 0 0 1 3.8-3.3A15.6 15.6 0 0 0 7.7 8zM12 20c-.8-1-1.4-2.3-1.8-4h3.6c-.4 1.7-1 3-1.8 4zm2.2-6H9.8a15.7 15.7 0 0 1 0-4h4.4a15.7 15.7 0 0 1 0 4zm.6 5.3a15.6 15.6 0 0 0 1.2-3.3h2.6a8 8 0 0 1-3.8 3.3zm2-5.3a17.6 17.6 0 0 0 0-4h2.9a7.8 7.8 0 0 1 0 4z" />
};

function icon(network: string): React.ReactElement {
  return ICONS[network] || ICONS.web;
}

const Footer: React.FC<IFooterProps> = (props) => {
  const { brandText, blurb, copyright, groups, social, accent } = props;
  const rootStyle = { ['--accent' as string]: accent } as React.CSSProperties;

  return (
    <footer className={styles.footer} style={rootStyle} role="contentinfo">
      <div className={styles.inner}>
        <div className={styles.brandCol}>
          <div className={styles.brand}>{brandText || 'Contoso'}</div>
          {blurb && <p className={styles.blurb}>{blurb}</p>}
          {social.length > 0 && (
            <div className={styles.social}>
              {social.map((s, i) => (
                <a key={i} className={styles.socialLink} href={s.url} target="_blank" rel="noreferrer" aria-label={s.network}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">{icon(s.network)}</svg>
                </a>
              ))}
            </div>
          )}
        </div>

        <div className={styles.cols}>
          {groups.map((g, gi) => (
            <div key={gi} className={styles.col}>
              <div className={styles.colHead}>{g.name}</div>
              <ul className={styles.colList}>
                {g.links.map((l, li) => (
                  <li key={li}>
                    <a className={styles.link} href={l.url} target="_blank" rel="noreferrer">{l.title}</a>
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </div>

      <div className={styles.bar}>
        <span>{copyright || `© ${brandText || 'Contoso'}`}</span>
      </div>
    </footer>
  );
};

export default Footer;
