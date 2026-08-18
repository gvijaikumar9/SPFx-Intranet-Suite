import * as React from 'react';
import { useMemo } from 'react';
import { Icon } from '@fluentui/react';
import styles from './QuickLinks.module.scss';
import { IQuickLinksProps } from './IQuickLinksProps';
import { readableTileText } from '../../shared/tileColors';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const QuickLinks: React.FC<IQuickLinksProps> = (props) => {
  const { title, layout, showTitle, frameStyle, columns, isDemo, accent, items, tileColors, loading, error } = props;
  const style = useMemo(
    () => ({
      ...frameStyle,
      ['--accent' as string]: accent,
      ['--cols' as string]: String(Math.max(1, Math.min(6, columns || 4)))
    } as React.CSSProperties),
    [accent, frameStyle, columns]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.quickLinks} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Quick links'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className={styles.muted}>No links yet. Add items to the list.</p>
      )}
      {!loading && !error && items.length > 0 && (
        <ul className={styles.grid}>
          {items.map((l, i) => {
            const bg = tileColors[i];
            const fg = readableTileText(bg);
            return (
            <li key={l.id} className={styles.cell}>
              <a
                className={`${styles.tile} ${fg ? styles.onDark : ''}`}
                style={bg ? { background: bg, borderColor: 'transparent', color: fg } : undefined}
                href={l.url || '#'}
                target="_blank"
                rel="noreferrer"
                aria-label={l.label}
              >
                <span className={styles.iconWrap} aria-hidden="true">
                  <Icon iconName={l.icon || 'Link'} />
                </span>
                <span className={styles.label}>{l.label}</span>
              </a>
            </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

export default QuickLinks;
