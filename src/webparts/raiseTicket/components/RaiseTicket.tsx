import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import styles from './RaiseTicket.module.scss';
import { IRaiseTicketProps } from './IRaiseTicketProps';
import { classifyStatus } from '../models/ITicket';

const STATUS_PILL: Record<string, string | undefined> = {
  'new': styles?.pillNew,
  'progress': styles?.pillProgress,
  'resolved': styles?.pillResolved
};

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const RaiseTicket: React.FC<IRaiseTicketProps> = (props) => {
  const {
    title, layout, showTitle, frameStyle, categories, isDemo, accent, configured,
    submitting, submitError, submittedId, recent, loadingRecent, recentError, onSubmit
  } = props;
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const [subject, setSubject] = useState('');
  const [category, setCategory] = useState(categories[0] || '');
  const [description, setDescription] = useState('');
  const [touched, setTouched] = useState(false);
  const [lastSeenId, setLastSeenId] = useState<number | undefined>(undefined);

  // When a submit succeeds (submittedId changes), clear the form and show the thank-you.
  useEffect(() => {
    if (submittedId !== undefined && submittedId !== lastSeenId) {
      setLastSeenId(submittedId);
      setSubject('');
      setDescription('');
      setCategory(categories[0] || '');
      setTouched(false);
    }
  }, [submittedId, lastSeenId, categories]);

  const justSubmitted = submittedId !== undefined && submittedId === lastSeenId;
  const subjectMissing = touched && subject.trim().length === 0;

  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );

  const handleSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    setTouched(true);
    if (subject.trim().length === 0) { return; }
    onSubmit(subject.trim(), category, description.trim());
  };

  return (
    <section className={`${styles.raiseTicket} ${layoutClass}`} style={style}>
      <div className={styles.grid}>
        {/* Form */}
        <form className={styles.card} onSubmit={handleSubmit} noValidate>
          {showTitle && <h2 className={styles.heading}>{title || 'Raise a ticket'}</h2>}
          {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

          <label className={styles.label} htmlFor="rt-subject">What do you need help with?</label>
          <input
            id="rt-subject"
            className={styles.input}
            type="text"
            value={subject}
            maxLength={255}
            placeholder="Short summary"
            onChange={(e) => setSubject(e.target.value)}
            aria-invalid={subjectMissing}
            aria-describedby={subjectMissing ? 'rt-subject-err' : undefined}
          />
          {subjectMissing && <span id="rt-subject-err" className={styles.fieldError}>Please add a short summary.</span>}

          <label className={styles.label} htmlFor="rt-category">Category</label>
          <select
            id="rt-category"
            className={styles.input}
            value={category}
            onChange={(e) => setCategory(e.target.value)}
          >
            {categories.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>

          <label className={styles.label} htmlFor="rt-desc">Details (optional)</label>
          <textarea
            id="rt-desc"
            className={styles.textarea}
            value={description}
            rows={3}
            maxLength={2000}
            placeholder="Anything that helps us fix it faster"
            onChange={(e) => setDescription(e.target.value)}
          />

          <div className={styles.actions}>
            <button type="submit" className={styles.button} disabled={submitting || !configured}>
              {submitting ? 'Sending...' : 'Submit ticket'}
            </button>
            {!configured && <span className={styles.fieldError}>No list configured yet.</span>}
          </div>

          {justSubmitted && !submitError && (
            <p className={styles.success} role="status">
              Thanks. Your ticket {submittedId && submittedId > 0 ? `#${submittedId}` : ''} was submitted.
            </p>
          )}
          {submitError && <p className={styles.error} role="alert">{submitError}</p>}
        </form>

        {/* Recent tickets */}
        <aside className={styles.card} aria-label="My recent tickets">
          <h3 className={styles.subheading}>My recent tickets</h3>
          {loadingRecent && <p className={styles.muted}>Loading...</p>}
          {recentError && !loadingRecent && <p className={styles.error}>{recentError}</p>}
          {!loadingRecent && !recentError && recent.length === 0 && (
            <p className={styles.muted}>Nothing yet. Your submitted tickets will show here.</p>
          )}
          {!loadingRecent && !recentError && recent.length > 0 && (
            <ul className={styles.list}>
              {recent.map((t) => {
                const cls = classifyStatus(t.status);
                return (
                  <li key={t.id} className={styles.item}>
                    <div className={styles.itemMain}>
                      <span className={styles.itemTitle}>{t.title}</span>
                      {t.category && <span className={styles.itemMeta}>{t.category}</span>}
                    </div>
                    <span className={`${styles.pill} ${STATUS_PILL[cls] || ''}`}>{t.status || 'New'}</span>
                  </li>
                );
              })}
            </ul>
          )}
        </aside>
      </div>
    </section>
  );
};

export default RaiseTicket;
