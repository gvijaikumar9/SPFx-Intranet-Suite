import * as React from 'react';
import { useMemo } from 'react';
import styles from './KpiTiles.module.scss';
import { IKpiTilesProps } from './IKpiTilesProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

// Returns a light text colour for dark fills, or undefined to keep the theme's text.
function readableText(bg: string | undefined): string | undefined {
  if (!bg) { return undefined; }
  const h = bg.replace('#', '');
  if (h.length !== 6) { return undefined; }
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255; // perceived brightness
  return lum < 0.55 ? '#ffffff' : undefined;
}

const TREND_ARROW: Record<string, string> = { up: '▲', down: '▼', flat: '–' };
const TREND_CLASS: Record<string, string | undefined> = {
  up: styles.up,
  down: styles.down,
  flat: styles.flat
};

const KpiTiles: React.FC<IKpiTilesProps> = (props) => {
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
    <section className={`${styles.kpiTiles} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Key numbers'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className={styles.muted}>No numbers yet. Add items to the list.</p>
      )}
      {!loading && !error && items.length > 0 && (
        <ul className={styles.grid}>
          {items.map((k, i) => {
            const bg = tileColors[i];
            const fg = readableText(bg);
            return (
            <li
              key={k.id}
              className={`${styles.tile} ${fg ? styles.onDark : ''}`}
              style={bg ? { background: bg, borderColor: 'transparent', color: fg } : undefined}
            >
              <span className={styles.value}>{k.value}</span>
              <span className={styles.label}>{k.label}</span>
              {(k.delta || k.trend !== 'flat') && (
                <span className={`${styles.trend} ${TREND_CLASS[k.trend] || ''}`}>
                  <span aria-hidden="true">{TREND_ARROW[k.trend]}</span>
                  {k.delta ? ` ${k.delta}` : ''}
                </span>
              )}
            </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

export default KpiTiles;
