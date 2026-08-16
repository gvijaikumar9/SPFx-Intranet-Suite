import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import { NormalPeoplePicker, IPersonaProps } from '@fluentui/react';
import styles from './Kudos.module.scss';
import { IKudosProps } from './IKudosProps';

/** Two-letter initials for the avatar bubble. */
function initials(name: string | undefined): string {
  if (!name) { return '?'; }
  const parts = name.trim().split(/\s+/);
  const first = parts[0] ? parts[0][0] : '';
  const last = parts.length > 1 ? parts[parts.length - 1][0] : '';
  return (first + last).toUpperCase() || '?';
}

/**
 * Avatar circle. With showPhoto on it tries the person's profile photo and falls back
 * to the text (initials) if the photo will not load; with it off it always shows text.
 */
const Avatar: React.FC<{ text: string; email?: string; webUrl: string; showPhoto: boolean }> = (p) => {
  const [failed, setFailed] = useState(false);
  const photoUrl = p.email
    ? `${p.webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(p.email)}`
    : '';
  if (p.showPhoto && photoUrl && !failed) {
    return <img className={styles.avatar} src={photoUrl} alt="" aria-hidden="true" onError={() => setFailed(true)} />;
  }
  return <span className={styles.avatar} aria-hidden="true">{p.text}</span>;
};

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const Kudos: React.FC<IKudosProps> = (props) => {
  const {
    title, layout, showTitle, frameStyle, showPhoto, webUrl, isDemo, accent, canGive, items, loading, error,
    submitting, submitError, submittedKey, onSearchPeople, onSubmit
  } = props;
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const [selected, setSelected] = useState<IPersonaProps[]>([]);
  const [message, setMessage] = useState('');
  const [touched, setTouched] = useState(false);
  const [lastKey, setLastKey] = useState(0);

  // Clear the form when a submit succeeds.
  useEffect(() => {
    if (submittedKey !== lastKey) {
      setLastKey(submittedKey);
      if (submittedKey > 0) {
        setSelected([]);
        setMessage('');
        setTouched(false);
      }
    }
  }, [submittedKey, lastKey]);

  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const justSubmitted = submittedKey > 0 && submittedKey === lastKey;

  const resolveSuggestions = useCallback(
    async (filter: string): Promise<IPersonaProps[]> => {
      if (!filter || filter.length < 2) { return []; }
      const people = await onSearchPeople(filter);
      return people.map((p) => ({ key: p.key, text: p.text, secondaryText: p.secondaryText }));
    },
    [onSearchPeople]
  );

  const recipient = selected[0];
  const recipientMissing = touched && !recipient;
  const messageMissing = touched && message.trim().length === 0;

  const handleSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    setTouched(true);
    if (!recipient || message.trim().length === 0) { return; }
    const login = recipient.key ? recipient.key.toString() : undefined;
    onSubmit(login, recipient.text || '', message.trim());
  };

  return (
    <section className={`${styles.kudos} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Kudos'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      <div className={styles.grid}>
        {/* Give kudos */}
        <form className={styles.giveCard} onSubmit={handleSubmit} noValidate>
          <h3 className={styles.subheading}>Give kudos</h3>

          <label className={styles.label} id="kd-to-label">To</label>
          <NormalPeoplePicker
            onResolveSuggestions={resolveSuggestions}
            selectedItems={selected}
            onChange={(picked) => setSelected(picked ? picked.slice(0, 1) : [])}
            itemLimit={1}
            disabled={!canGive}
            resolveDelay={300}
            inputProps={{ 'aria-labelledby': 'kd-to-label', placeholder: 'Search for a colleague' }}
            pickerSuggestionsProps={{ noResultsFoundText: 'No people found', loadingText: 'Searching...' }}
          />
          {recipientMissing && <span className={styles.fieldError}>Please choose who this is for.</span>}

          <label className={styles.label} htmlFor="kd-msg">Message</label>
          <textarea
            id="kd-msg"
            className={styles.textarea}
            value={message}
            rows={3}
            maxLength={2000}
            placeholder="Say what they did well"
            disabled={!canGive}
            onChange={(e) => setMessage(e.target.value)}
            aria-invalid={messageMissing}
          />
          {messageMissing && <span className={styles.fieldError}>Please add a short message.</span>}

          <div className={styles.actions}>
            <button type="submit" className={styles.button} disabled={submitting || !canGive}>
              {submitting ? 'Sending...' : 'Send kudos'}
            </button>
            {!canGive && <span className={styles.fieldError}>No list configured yet.</span>}
          </div>

          {justSubmitted && !submitError && <p className={styles.success} role="status">Kudos sent. Thank you.</p>}
          {submitError && <p className={styles.error} role="alert">{submitError}</p>}
        </form>

        {/* Recognition wall */}
        <div className={styles.wall} aria-label="Recent kudos">
          <h3 className={styles.subheading}>Recognition wall</h3>
          {loading && <p className={styles.muted}>Loading...</p>}
          {error && !loading && <p className={styles.error}>{error}</p>}
          {!loading && !error && items.length === 0 && (
            <p className={styles.muted}>No kudos yet. Be the first to recognise someone.</p>
          )}
          {!loading && !error && items.length > 0 && (
            <ul className={styles.list}>
              {items.map((k) => {
                const hasTo = !!k.toName;
                const avatarText = hasTo
                  ? initials(k.toName)
                  : (k.headline ? k.headline.trim().charAt(0).toUpperCase() : '★');
                return (
                  <li key={k.id} className={styles.item}>
                    <Avatar
                      text={avatarText}
                      email={hasTo ? k.toEmail : undefined}
                      webUrl={webUrl}
                      showPhoto={showPhoto}
                    />
                    <div className={styles.itemBody}>
                      <p className={styles.itemLine}>
                        {hasTo ? (
                          <>
                            {k.fromName && <><span className={styles.from}>{k.fromName}</span>{' thanked '}</>}
                            {!k.fromName && 'Kudos to '}
                            <span className={styles.to}>{k.toName}</span>
                          </>
                        ) : (
                          <span className={styles.to}>{k.headline || 'Kudos'}</span>
                        )}
                      </p>
                      <p className={styles.msg}>{k.message}</p>
                    </div>
                  </li>
                );
              })}
            </ul>
          )}
        </div>
      </div>
    </section>
  );
};

export default Kudos;
