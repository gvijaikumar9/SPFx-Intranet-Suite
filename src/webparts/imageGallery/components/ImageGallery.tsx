import * as React from 'react';
import * as ReactDom from 'react-dom';
import { useMemo, useState, useEffect } from 'react';
import styles from './ImageGallery.module.scss';
import { IImageGalleryProps } from './IImageGalleryProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

// Gradient placeholders for demo mode (no real images are loaded).
const PLACEHOLDERS = [
  'linear-gradient(135deg, #0f6cbd, #4aa3df)',
  'linear-gradient(135deg, #107c41, #4cc38a)',
  'linear-gradient(135deg, #8764b8, #c39bd3)',
  'linear-gradient(135deg, #ca5010, #f0a170)',
  'linear-gradient(135deg, #038387, #4bc0c4)',
  'linear-gradient(135deg, #5c2e91, #9a7bc8)'
];

const ImageGallery: React.FC<IImageGalleryProps> = (props) => {
  const { title, layout, showTitle, frameStyle, columns, isDemo, accent, items, loading, error } = props;
  const [openIndex, setOpenIndex] = useState<number | undefined>(undefined);

  const style = useMemo(
    () => ({
      ...frameStyle,
      ['--accent' as string]: accent,
      ['--cols' as string]: String(Math.max(2, Math.min(6, columns || 4)))
    } as React.CSSProperties),
    [accent, frameStyle, columns]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const count = items.length;
  const step = (dir: number): void => {
    setOpenIndex((prev) => (prev === undefined ? undefined : (prev + dir + count) % count));
  };

  // Keyboard: Escape closes, arrows navigate, while the lightbox is open.
  useEffect(() => {
    if (openIndex === undefined) { return undefined; }
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') { setOpenIndex(undefined); }
      else if (e.key === 'ArrowRight') { step(1); }
      else if (e.key === 'ArrowLeft') { step(-1); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [openIndex, count]);

  const current = openIndex !== undefined ? items[openIndex] : undefined;

  return (
    <section className={`${styles.gallery} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Gallery'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Upload photos to the library to show them here.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && count === 0 && (
        <p className={styles.muted}>No photos yet. Upload images to the library.</p>
      )}
      {!loading && !error && count > 0 && (
        <ul className={styles.grid}>
          {items.map((img, i) => (
            <li key={img.id} className={styles.cell}>
              {img.url ? (
                <button type="button" className={styles.tile} aria-label={`Open ${img.name}`} onClick={() => setOpenIndex(i)}>
                  <img className={styles.img} src={img.url} alt={img.name} loading="lazy" />
                </button>
              ) : (
                <div className={styles.tile} style={{ background: PLACEHOLDERS[i % PLACEHOLDERS.length] }}>
                  <span className={styles.placeholder}>{img.name}</span>
                </div>
              )}
            </li>
          ))}
        </ul>
      )}

      {current && current.url && ReactDom.createPortal(
        <div
          className={styles.lightbox}
          role="dialog"
          aria-modal="true"
          aria-label={current.name}
          onClick={() => setOpenIndex(undefined)}
        >
          <button type="button" className={styles.lbClose} aria-label="Close" autoFocus onClick={() => setOpenIndex(undefined)}>×</button>
          {count > 1 && (
            <button
              type="button"
              className={`${styles.lbNav} ${styles.lbPrev}`}
              aria-label="Previous image"
              onClick={(e) => { e.stopPropagation(); step(-1); }}
            >‹</button>
          )}
          <figure className={styles.lbFigure} onClick={(e) => e.stopPropagation()}>
            <img className={styles.lbImg} src={current.url} alt={current.name} />
            <figcaption className={styles.lbCaption}>{current.name}</figcaption>
          </figure>
          {count > 1 && (
            <button
              type="button"
              className={`${styles.lbNav} ${styles.lbNext}`}
              aria-label="Next image"
              onClick={(e) => { e.stopPropagation(); step(1); }}
            >›</button>
          )}
        </div>,
        document.body
      )}
    </section>
  );
};

export default ImageGallery;
