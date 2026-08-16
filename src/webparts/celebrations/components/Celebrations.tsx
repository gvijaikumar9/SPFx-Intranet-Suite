import * as React from 'react';
import { useMemo, useState } from 'react';
import styles from './Celebrations.module.scss';
import { ICelebrationsProps } from './ICelebrationsProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function initials(name: string | undefined): string {
  if (!name) { return '?'; }
  const parts = name.trim().split(/\s+/);
  const a = parts[0] ? parts[0][0] : '';
  const b = parts.length > 1 ? parts[parts.length - 1][0] : '';
  return (a + b).toUpperCase() || '?';
}

function isAnniversary(type: string): boolean {
  return type.toLowerCase().indexOf('anniv') >= 0 || type.toLowerCase().indexOf('work') >= 0;
}

const Avatar: React.FC<{ name: string; email?: string; webUrl: string; showPhoto: boolean }> = (props) => {
  const { name, email, webUrl, showPhoto } = props;
  const [failed, setFailed] = useState(false);
  const url = email ? `${webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}` : '';
  if (showPhoto && url && !failed) {
    return <img className={styles.avatar} src={url} alt="" aria-hidden="true" onError={() => setFailed(true)} />;
  }
  return <span className={styles.avatar} aria-hidden="true">{initials(name)}</span>;
};

const Celebrations: React.FC<ICelebrationsProps> = (props) => {
  const { title, layout, showTitle, frameStyle, showPhoto, webUrl, isDemo, accent, items, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.celebrations} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Celebrations'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className={styles.muted}>No celebrations coming up.</p>
      )}
      {!loading && !error && items.length > 0 && (
        <ul className={styles.list}>
          {items.map((c) => {
            const d = new Date(c.date);
            const valid = !isNaN(d.getTime());
            const anniv = isAnniversary(c.type);
            return (
              <li key={c.id} className={styles.item}>
                <Avatar name={c.name} email={c.email} webUrl={webUrl} showPhoto={showPhoto} />
                <div className={styles.body}>
                  <p className={styles.name}>
                    <span className={styles.emoji} aria-hidden="true">{anniv ? '🎉' : '🎂'}</span>
                    {c.name}
                  </p>
                  <p className={styles.meta}>
                    <span>{c.type || (anniv ? 'Anniversary' : 'Birthday')}</span>
                    {valid && <span className={styles.date}>{d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}</span>}
                  </p>
                </div>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

export default Celebrations;
