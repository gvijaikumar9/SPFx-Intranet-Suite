import * as React from 'react';
import { useState, useMemo } from 'react';
import styles from './EmployeeSpotlight.module.scss';
import { IEmployeeSpotlightProps } from './IEmployeeSpotlightProps';
import { IEmployee } from '../models/IEmployee';

function initials(name: string | undefined): string {
  if (!name) { return '?'; }
  const parts = name.trim().split(/\s+/);
  const first = parts[0] ? parts[0][0] : '';
  const last = parts.length > 1 ? parts[parts.length - 1][0] : '';
  return (first + last).toUpperCase() || '?';
}

/**
 * Avatar circle. When showPhoto is on it tries the person's profile photo and falls
 * back to initials only if that photo will not load. When showPhoto is off it always
 * shows initials.
 */
const Avatar: React.FC<{ name: string; email?: string; webUrl: string; showPhoto: boolean; big?: boolean }> = (props) => {
  const { name, email, webUrl, showPhoto, big } = props;
  const [failed, setFailed] = useState(false);
  const photoUrl = email
    ? `${webUrl}/_layouts/15/userphoto.aspx?size=${big ? 'L' : 'M'}&accountname=${encodeURIComponent(email)}`
    : '';
  const cls = big ? `${styles.avatar} ${styles.avatarBig}` : styles.avatar;

  if (showPhoto && photoUrl && !failed) {
    return <img className={cls} src={photoUrl} alt="" aria-hidden="true" onError={() => setFailed(true)} />;
  }
  return <span className={cls} aria-hidden="true">{initials(name)}</span>;
};

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const EmployeeSpotlight: React.FC<IEmployeeSpotlightProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, webUrl, showPhoto, items, loading, error, emptyNote } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );

  const latest: IEmployee | undefined = items[0];
  const past = items.slice(1);
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.spotlight} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Employee of the month'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && !latest && (
        <p className={styles.muted}>{emptyNote || 'No one has been recognised yet.'}</p>
      )}

      {!loading && !error && latest && (
        <>
          <div className={styles.card}>
            <Avatar name={latest.name} email={latest.email} webUrl={webUrl} showPhoto={showPhoto} big />
            <div className={styles.cardBody}>
              <p className={styles.eyebrow}>
                Employee of the month{latest.month ? ` · ${latest.month}` : ''}
              </p>
              <p className={styles.name}>{latest.name}</p>
              {latest.citation && <p className={styles.citation}>&ldquo;{latest.citation}&rdquo;</p>}
            </div>
          </div>

          {past.length > 0 && (
            <div className={styles.past}>
              <p className={styles.pastLabel}>Previously recognised</p>
              <ul className={styles.pastList}>
                {past.map((p) => (
                  <li key={p.id} className={styles.chip}>
                    <Avatar name={p.name} email={p.email} webUrl={webUrl} showPhoto={showPhoto} />
                    <span className={styles.chipText}>
                      <span className={styles.chipName}>{p.name}</span>
                      {p.month && <span className={styles.chipMonth}>{p.month}</span>}
                    </span>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </>
      )}
    </section>
  );
};

export default EmployeeSpotlight;
