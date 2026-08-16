import * as React from 'react';
import { useMemo } from 'react';
import styles from './EventsList.module.scss';
import { IEventsListProps } from './IEventsListProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function parse(iso: string | undefined): Date | undefined {
  if (!iso) { return undefined; }
  const d = new Date(iso);
  return isNaN(d.getTime()) ? undefined : d;
}

function day(d: Date): string {
  return String(d.getDate());
}

function month(d: Date): string {
  return d.toLocaleDateString(undefined, { month: 'short' });
}

function time(d: Date): string {
  return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
}

const EventsList: React.FC<IEventsListProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, items, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.events} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Upcoming events'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className={styles.muted}>Nothing coming up right now.</p>
      )}
      {!loading && !error && items.length > 0 && (
        <ul className={styles.list}>
          {items.map((e) => {
            const start = parse(e.start);
            const end = parse(e.end);
            return (
              <li key={e.id} className={styles.item}>
                <div className={styles.date} aria-hidden="true">
                  {start ? <><span className={styles.dateDay}>{day(start)}</span><span className={styles.dateMonth}>{month(start)}</span></> : <span className={styles.dateMonth}>TBD</span>}
                </div>
                <div className={styles.body}>
                  <p className={styles.eventTitle}>{e.title}</p>
                  <p className={styles.meta}>
                    {start && <span>{time(start)}{end ? ` to ${time(end)}` : ''}</span>}
                    {e.location && <span className={styles.location}>{e.location}</span>}
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

export default EventsList;
