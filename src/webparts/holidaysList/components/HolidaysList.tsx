import * as React from 'react';
import { useMemo } from 'react';
import styles from './HolidaysList.module.scss';
import { IHolidaysListProps } from './IHolidaysListProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function parse(iso: string): Date | undefined {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? undefined : d;
}

const HolidaysList: React.FC<IHolidaysListProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, items, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.holidays} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Upcoming holidays'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className={styles.muted}>No holidays coming up.</p>
      )}
      {!loading && !error && items.length > 0 && (
        <ul className={styles.list}>
          {items.map((h) => {
            const d = parse(h.date);
            return (
              <li key={h.id} className={styles.item}>
                <div className={styles.date} aria-hidden="true">
                  {d
                    ? <>
                        <span className={styles.day}>{d.getDate()}</span>
                        <span className={styles.month}>{d.toLocaleDateString(undefined, { month: 'short' })}</span>
                      </>
                    : <span className={styles.month}>TBD</span>}
                </div>
                <div className={styles.body}>
                  <p className={styles.name}>{h.name}</p>
                  {(d || h.region) && (
                    <p className={styles.meta}>
                      {d && <span>{d.toLocaleDateString(undefined, { weekday: 'long' })}</span>}
                      {h.region && <span className={styles.region}>{h.region}</span>}
                    </p>
                  )}
                </div>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

export default HolidaysList;
