import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import styles from './FeedbackWidget.module.scss';

export interface IFeedbackWidgetProps {
  accent: string;
  userName: string;      // display name, shown on the opt-in line
  userContact: string;   // actionable contact stored when opted in (e.g. "Name <email>")
  onSubmit: (rating: number, comment: string, contact: string) => Promise<boolean>;
}

type SendState = 'idle' | 'sending' | 'done' | 'error';

const MAX_COMMENT = 1000; // generous cap; the Comment list column (Note) holds far more

const FeedbackWidget: React.FC<IFeedbackWidgetProps> = (props) => {
  const { accent, userName, userContact, onSubmit } = props;
  const [open, setOpen] = useState<boolean>(false);
  const [rating, setRating] = useState<number>(0);
  const [comment, setComment] = useState<string>('');
  const [contactOk, setContactOk] = useState<boolean>(false);
  const [state, setState] = useState<SendState>('idle');

  const sendingRef = useRef<boolean>(false);          // synchronous double-submit guard
  const closedRef = useRef<boolean>(false);           // ignore a submit that resolves after close
  const doneTimerRef = useRef<number | undefined>(undefined);

  const rootStyle = { ['--accent' as string]: accent } as React.CSSProperties;
  const contactLabel = userName || userContact;       // opt-in shows a name, or the email if no name

  const clearTimer = (): void => {
    if (doneTimerRef.current !== undefined) { window.clearTimeout(doneTimerRef.current); doneTimerRef.current = undefined; }
  };

  // clean up the auto-close timer if SPFx navigates away / unmounts
  useEffect(() => (): void => { clearTimer(); }, []);

  const reset = (): void => { setRating(0); setComment(''); setContactOk(false); setState('idle'); };

  const openPanel = (): void => { closedRef.current = false; setOpen(true); };

  const close = (): void => {
    closedRef.current = true;   // any in-flight submit result is now ignored
    sendingRef.current = false;
    clearTimer();
    reset();
    setOpen(false);
  };

  const send = (): void => {
    if (sendingRef.current) { return; } // hard guard against double-submit
    sendingRef.current = true;
    setState('sending');
    onSubmit(rating, comment, contactOk ? userContact : '')
      .then((ok) => {
        sendingRef.current = false;
        if (closedRef.current) { return; } // panel was closed meanwhile
        setState(ok ? 'done' : 'error');
        if (ok) {
          clearTimer();
          doneTimerRef.current = window.setTimeout(() => { setOpen(false); reset(); }, 1800);
        }
      })
      .catch(() => {
        sendingRef.current = false;
        if (closedRef.current) { return; }
        setState('error');
      });
  };

  const canSend = (rating > 0 || comment.trim().length > 0) && state !== 'sending';

  return (
    <div className={styles.root} style={rootStyle}>
      {!open && (
        <button className={styles.fab} type="button" onClick={openPanel} aria-label="Give feedback">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
          </svg>
          <span className={styles.fabText}>Feedback</span>
        </button>
      )}

      {open && (
        <div className={styles.panel} role="dialog" aria-label="Share feedback">
          <div className={styles.head}>
            <span className={styles.title}>Share your feedback</span>
            <button className={styles.close} type="button" onClick={close} aria-label="Close">×</button>
          </div>

          {state !== 'done' && (
            <div className={styles.body}>
              <div className={styles.stars} role="radiogroup" aria-label="Rating">
                {[1, 2, 3, 4, 5].map((n) => (
                  <button
                    key={n}
                    type="button"
                    role="radio"
                    className={`${styles.star} ${n <= rating ? styles.on : ''}`}
                    onClick={() => setRating(n)}
                    aria-label={`${n} star${n > 1 ? 's' : ''}`}
                    aria-checked={n === rating}
                  >★</button>
                ))}
              </div>

              <textarea
                className={styles.comment}
                placeholder="What's working well, and what could be better?"
                value={comment}
                maxLength={MAX_COMMENT}
                onChange={(e) => setComment(e.target.value.substring(0, MAX_COMMENT))}
                rows={3}
              />
              <div className={`${styles.counter} ${comment.length >= MAX_COMMENT ? styles.counterMax : ''}`}>
                {comment.length} / {MAX_COMMENT}
              </div>

              {contactLabel && (
                <label className={styles.contact}>
                  <input type="checkbox" checked={contactOk} onChange={(e) => setContactOk(e.target.checked)} />
                  <span>OK to follow up with me ({contactLabel})</span>
                </label>
              )}

              {state === 'error' && <p className={styles.error}>Could not send right now. Please try again.</p>}

              <button className={styles.send} type="button" disabled={!canSend} onClick={send}>
                {state === 'sending' ? 'Sending…' : 'Send feedback'}
              </button>
            </div>
          )}

          {state === 'done' && (
            <div className={styles.done}>
              <span className={styles.tick} aria-hidden="true">✓</span>
              Thanks for your feedback!
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default FeedbackWidget;
