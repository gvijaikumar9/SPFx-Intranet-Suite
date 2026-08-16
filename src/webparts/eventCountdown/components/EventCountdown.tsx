import * as React from 'react';
import { useMemo, useState, useEffect } from 'react';
import styles from './EventCountdown.module.scss';
import { IEventCountdownProps } from './IEventCountdownProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

interface IParts { d: number; h: number; m: number; s: number; }

function split(ms: number): IParts {
  const total = Math.max(0, Math.floor(ms / 1000));
  return {
    d: Math.floor(total / 86400),
    h: Math.floor((total % 86400) / 3600),
    m: Math.floor((total % 3600) / 60),
    s: total % 60
  };
}

const EventCountdown: React.FC<IEventCountdownProps> = (props) => {
  const { title, layout, showTitle, frameStyle, eventName, targetIso, isDemo, accent } = props;
  const [nowMs, setNowMs] = useState<number>(() => new Date().getTime());

  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const targetMs = useMemo(() => {
    const d = new Date(targetIso);
    return isNaN(d.getTime()) ? undefined : d.getTime();
  }, [targetIso]);

  useEffect(() => {
    const timer = window.setInterval(() => setNowMs(new Date().getTime()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  const remaining = targetMs === undefined ? undefined : targetMs - nowMs;
  const parts = remaining !== undefined ? split(remaining) : undefined;
  const passed = remaining !== undefined && remaining <= 0;

  const cell = (value: number, label: string): React.ReactElement => (
    <div className={styles.cell}>
      <span className={styles.num}>{value < 10 ? `0${value}` : value}</span>
      <span className={styles.lbl}>{label}</span>
    </div>
  );

  return (
    <section className={`${styles.countdown} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Counting down to'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Set an event name and date in the property pane.</p>}

      <p className={styles.eventName}>{eventName || 'Company kickoff'}</p>

      {targetMs === undefined && (
        <p className={styles.muted}>Set a valid target date in the property pane.</p>
      )}
      {parts && !passed && (
        <div className={styles.grid}>
          {cell(parts.d, parts.d === 1 ? 'day' : 'days')}
          {cell(parts.h, 'hours')}
          {cell(parts.m, 'mins')}
          {cell(parts.s, 'secs')}
        </div>
      )}
      {passed && <p className={styles.passed}>This event is happening now.</p>}
    </section>
  );
};

export default EventCountdown;
