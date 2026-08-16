import * as React from 'react';
import styles from './AnnouncementsTicker.module.scss';
import type { IAnnouncementsTickerProps } from './IAnnouncementsTickerProps';
import type { Severity } from '../models/IAnnouncement';

const { useState, useEffect, useCallback } = React;

const SEVERITY_PILL: { [K in Severity]?: string } = {
  warning: styles?.pillWarning,
  critical: styles?.pillCritical
};

const SEVERITY_LABEL: { [K in Severity]: string } = {
  info: 'Info',
  warning: 'Warning',
  critical: 'Critical'
};

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles?.layoutCard,
  minimal: styles?.layoutMinimal,
  bold: styles?.layoutBold,
  compact: styles?.layoutCompact
};

const AnnouncementsTicker: React.FC<IAnnouncementsTickerProps> = (props) => {
  const { label, items, rotateSeconds, showControls, loading, error, isDemo, accent, layout, showTitle, frameStyle } = props;
  const layoutClass = LAYOUT_CLASS[layout] || styles?.layoutCard;
  const [index, setIndex] = useState<number>(0);
  const [paused, setPaused] = useState<boolean>(false);
  const count = items.length;

  // Keep the active index valid if the item set shrinks (e.g. property change).
  useEffect(() => {
    if (index > count - 1) {
      setIndex(0);
    }
  }, [count, index]);

  const go = useCallback((dir: number) => {
    setIndex(prev => (count === 0 ? 0 : (prev + dir + count) % count));
  }, [count]);

  // Auto-rotation. Respects reduced-motion, pause-on-hover/focus, and single-item sets.
  useEffect(() => {
    if (paused || count <= 1) {
      return undefined;
    }
    const reduce = typeof window !== 'undefined'
      && typeof window.matchMedia === 'function'
      && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduce) {
      return undefined;
    }
    const ms = Math.max(3, rotateSeconds) * 1000;
    const timer = window.setInterval(() => {
      setIndex(prev => (prev + 1) % count);
    }, ms);
    return () => window.clearInterval(timer);
  }, [paused, count, rotateSeconds]);

  const rootStyle = { ['--ticker-accent' as string]: accent } as React.CSSProperties;

  const renderBody = (): React.ReactNode => {
    if (loading) {
      return <div className={styles.state}>Loading announcements&hellip;</div>;
    }
    if (error) {
      return <div className={styles.state}>{error}</div>;
    }
    if (count === 0) {
      return <div className={styles.state}>No announcements right now.</div>;
    }
    const current = items[Math.min(index, count - 1)];
    const pillClass = SEVERITY_PILL[current.severity];
    return (
      <div className={styles.stage} aria-live="polite" aria-atomic="true">
        <div key={current.id} className={styles.msg}>
          {pillClass && (
            <span className={`${styles.pill} ${pillClass}`}>{SEVERITY_LABEL[current.severity]}</span>
          )}
          {current.href
            ? <a className={styles.link} href={current.href} target="_blank" rel="noreferrer">{current.message}</a>
            : <span className={styles.text}>{current.message}</span>}
        </div>
      </div>
    );
  };

  return (
    <div style={frameStyle}>
    <section
      className={`${styles.ticker} ${layoutClass || ''}`}
      style={rootStyle}
      role="region"
      aria-label={label}
      onMouseEnter={() => setPaused(true)}
      onMouseLeave={() => setPaused(false)}
      onFocus={() => setPaused(true)}
      onBlur={() => setPaused(false)}
    >
      {showTitle && (
        <span className={styles.badge}>
          <span className={styles.megaphone} aria-hidden="true">&#128227;</span>
          <span className={styles.labelText}>{label}{isDemo ? ' · sample' : ''}</span>
        </span>
      )}

      {renderBody()}

      {showControls && count > 1 && !loading && !error && (
        <div className={styles.controls}>
          <button type="button" className={styles.arrow} aria-label="Previous announcement" onClick={() => go(-1)}>&#8249;</button>
          <div className={styles.dots} role="tablist" aria-label="Announcements">
            {items.map((it, i) => (
              <button
                key={it.id}
                type="button"
                role="tab"
                aria-selected={i === index}
                aria-label={`Announcement ${i + 1} of ${count}`}
                className={`${styles.dot} ${i === index ? styles.dotOn : ''}`}
                onClick={() => setIndex(i)}
              />
            ))}
          </div>
          <button type="button" className={styles.arrow} aria-label="Next announcement" onClick={() => go(1)}>&#8250;</button>
        </div>
      )}
    </section>
    </div>
  );
};

export default AnnouncementsTicker;
