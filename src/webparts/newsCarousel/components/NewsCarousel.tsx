import * as React from 'react';
import { useMemo, useState, useEffect, useCallback } from 'react';
import styles from './NewsCarousel.module.scss';
import { INewsCarouselProps } from './INewsCarouselProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const GRADIENTS = [
  'linear-gradient(135deg, #0f6cbd, #123a6b)',
  'linear-gradient(135deg, #107c41, #0b5132)',
  'linear-gradient(135deg, #8764b8, #4b2e77)',
  'linear-gradient(135deg, #ca5010, #7a2f0a)'
];

const NewsCarousel: React.FC<INewsCarouselProps> = (props) => {
  const { title, layout, showTitle, frameStyle, rotateSeconds, isDemo, accent, items, loading, error } = props;
  const [index, setIndex] = useState(0);
  const [paused, setPaused] = useState(false);
  const count = items.length;

  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  useEffect(() => {
    if (index > count - 1) { setIndex(0); }
  }, [count, index]);

  const go = useCallback((dir: number) => {
    setIndex((prev) => (count === 0 ? 0 : (prev + dir + count) % count));
  }, [count]);

  useEffect(() => {
    if (paused || count <= 1) { return undefined; }
    const reduce = typeof window !== 'undefined'
      && typeof window.matchMedia === 'function'
      && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduce) { return undefined; }
    const ms = Math.max(3, rotateSeconds || 6) * 1000;
    const timer = window.setInterval(() => setIndex((prev) => (prev + 1) % count), ms);
    return () => window.clearInterval(timer);
  }, [paused, count, rotateSeconds]);

  const current = count > 0 ? items[Math.min(index, count - 1)] : undefined;
  // Strip quotes, parens and backslashes so a list value cannot break out of the CSS url().
  const safeImg = current && current.imageUrl ? current.imageUrl.replace(/["')(\\]/g, '') : '';
  const bg = safeImg
    ? `center/cover no-repeat url('${safeImg}')`
    : GRADIENTS[index % GRADIENTS.length];

  return (
    <section className={`${styles.news} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'News'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && !current && (
        <p className={styles.muted}>No news yet. Add items to the list.</p>
      )}

      {!loading && !error && current && (
        <div
          className={styles.stage}
          onMouseEnter={() => setPaused(true)}
          onMouseLeave={() => setPaused(false)}
          onFocus={() => setPaused(true)}
          onBlur={() => setPaused(false)}
        >
          <div className={styles.slide} style={{ background: bg }}>
            <div className={styles.scrim} />
            <div className={styles.cap}>
              {current.category && <span className={styles.kicker}>{current.category}</span>}
              <h3 className={styles.slideTitle}>
                {current.linkUrl
                  ? <a className={styles.slideLink} href={current.linkUrl} target="_blank" rel="noreferrer">{current.title}</a>
                  : current.title}
              </h3>
              {current.summary && <p className={styles.summary}>{current.summary}</p>}
            </div>
          </div>

          {count > 1 && (
            <>
              <button type="button" className={`${styles.arrow} ${styles.prev}`} aria-label="Previous story" onClick={() => go(-1)}>‹</button>
              <button type="button" className={`${styles.arrow} ${styles.next}`} aria-label="Next story" onClick={() => go(1)}>›</button>
              <div className={styles.dots} role="tablist" aria-label="Stories">
                {items.map((it, i) => (
                  <button
                    key={it.id}
                    type="button"
                    role="tab"
                    aria-selected={i === index}
                    aria-label={`Story ${i + 1} of ${count}`}
                    className={`${styles.dot} ${i === index ? styles.dotOn : ''}`}
                    onClick={() => setIndex(i)}
                  />
                ))}
              </div>
            </>
          )}
        </div>
      )}
    </section>
  );
};

export default NewsCarousel;
