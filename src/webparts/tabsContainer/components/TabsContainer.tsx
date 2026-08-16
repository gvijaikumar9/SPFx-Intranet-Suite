import * as React from 'react';
import { useMemo, useState, useEffect } from 'react';
import styles from './TabsContainer.module.scss';
import { ITabsContainerProps } from './ITabsContainerProps';
import { tabsOf } from '../models/ITabItem';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

let tabsUidCounter = 0;

const TabsContainer: React.FC<ITabsContainerProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, items, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const tabs = useMemo(() => tabsOf(items), [items]);
  const [active, setActive] = useState(0);
  const uid = useMemo(() => `tabs-${++tabsUidCounter}`, []);

  // Keep the active tab valid if the tab set changes.
  useEffect(() => {
    if (active > tabs.length - 1) { setActive(0); }
  }, [tabs.length, active]);

  // Roving arrow-key focus across the tabs (WAI-ARIA tab pattern).
  const onTabKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    let next = -1;
    if (e.key === 'ArrowRight') { next = (active + 1) % tabs.length; }
    else if (e.key === 'ArrowLeft') { next = (active - 1 + tabs.length) % tabs.length; }
    if (next >= 0) {
      e.preventDefault();
      setActive(next);
      const el = document.getElementById(`${uid}-tab-${next}`);
      if (el) { el.focus(); }
    }
  };

  const activeTab = tabs[Math.min(active, tabs.length - 1)];
  const rows = items.filter((it) => (it.tab || 'General') === activeTab);

  return (
    <section className={`${styles.tabs} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || "What's happening"}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && tabs.length === 0 && (
        <p className={styles.muted}>Nothing to show yet.</p>
      )}
      {!loading && !error && tabs.length > 0 && (
        <>
          <div className={styles.tabbar} role="tablist" aria-label={title || 'Tabs'} onKeyDown={onTabKey}>
            {tabs.map((t, i) => (
              <button
                key={t}
                id={`${uid}-tab-${i}`}
                type="button"
                role="tab"
                aria-selected={i === active}
                aria-controls={`${uid}-panel`}
                tabIndex={i === active ? 0 : -1}
                className={`${styles.tab} ${i === active ? styles.tabOn : ''}`}
                onClick={() => setActive(i)}
              >{t}</button>
            ))}
          </div>
          <ul className={styles.list} role="tabpanel" id={`${uid}-panel`} aria-labelledby={`${uid}-tab-${active}`}>
            {rows.map((it) => (
              <li key={it.id} className={styles.item}>
                <div className={styles.itemBody}>
                  {it.href
                    ? <a className={styles.itemTitle} href={it.href} target="_blank" rel="noreferrer">{it.title}</a>
                    : <span className={styles.itemTitle}>{it.title}</span>}
                  {it.subtitle && <span className={styles.itemSub}>{it.subtitle}</span>}
                </div>
              </li>
            ))}
            {rows.length === 0 && <li className={styles.muted}>Nothing in this tab.</li>}
          </ul>
        </>
      )}
    </section>
  );
};

export default TabsContainer;
