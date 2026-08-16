import * as React from 'react';
import { useState, useMemo } from 'react';
import styles from './FaqAccordion.module.scss';
import { IFaqAccordionProps } from './IFaqAccordionProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const FaqAccordion: React.FC<IFaqAccordionProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, items, loading, error } = props;
  const [openIds, setOpenIds] = useState<number[]>([]);
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const toggle = (id: number): void => {
    setOpenIds((prev) => (prev.indexOf(id) >= 0 ? prev.filter((x) => x !== id) : prev.concat(id)));
  };

  return (
    <section className={`${styles.faq} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Frequently asked questions'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className={styles.muted}>No questions yet. Add items to the list.</p>
      )}
      {!loading && !error && items.length > 0 && (
        <ul className={styles.list}>
          {items.map((f) => {
            const open = openIds.indexOf(f.id) >= 0;
            const panelId = `faq-panel-${f.id}`;
            return (
              <li key={f.id} className={styles.item}>
                <button
                  type="button"
                  className={styles.question}
                  aria-expanded={open}
                  aria-controls={panelId}
                  onClick={() => toggle(f.id)}
                >
                  <span className={styles.qText}>
                    {f.question}
                    {f.category && <span className={styles.tag}>{f.category}</span>}
                  </span>
                  <span className={`${styles.chevron} ${open ? styles.chevronOpen : ''}`} aria-hidden="true">›</span>
                </button>
                {open && (
                  <div id={panelId} className={styles.answer}>
                    {f.answer}
                  </div>
                )}
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

export default FaqAccordion;
