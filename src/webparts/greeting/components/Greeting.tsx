import * as React from 'react';
import { useMemo } from 'react';
import styles from './Greeting.module.scss';
import { IGreetingProps } from './IGreetingProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function timeGreeting(): string {
  const h = new Date().getHours();
  if (h < 12) { return 'Good morning'; }
  if (h < 18) { return 'Good afternoon'; }
  return 'Good evening';
}

const Greeting: React.FC<IGreetingProps> = (props) => {
  const { layout, frameStyle, name, chips, accent } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.greeting} ${layoutClass}`} style={style}>
      <h1 className={styles.hi}>
        {timeGreeting()}, <span className={styles.name}>{name || 'there'}</span>
      </h1>
      {chips.length > 0 && (
        <div className={styles.chips}>
          {chips.map((c, i) => (
            <span key={i} className={styles.chip}>
              {c.count ? <span className={styles.cnt}>{c.count}</span> : undefined}
              <span className={styles.chipLbl}>{c.label}</span>
            </span>
          ))}
        </div>
      )}
    </section>
  );
};

export default Greeting;
