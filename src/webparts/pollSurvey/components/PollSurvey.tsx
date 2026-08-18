import * as React from 'react';
import { useMemo } from 'react';
import styles from './PollSurvey.module.scss';
import { IPollSurveyProps, IPollOption } from './IPollSurveyProps';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

const PollSurvey: React.FC<IPollSurveyProps> = (props) => {
  const {
    title, layout, showTitle, frameStyle, question, options, results, total,
    voted, configured, isDemo, accent, submitting, submitError, optionColors, loading, error, onVote
  } = props;

  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const countFor = (option: string): number => {
    const found = results.filter((r: IPollOption) => r.option === option)[0];
    return found ? found.count : 0;
  };
  const pct = (n: number): number => (total > 0 ? Math.round((n / total) * 100) : 0);

  return (
    <section className={`${styles.poll} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Poll of the week'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Votes are not saved until you point this at a list.</p>}

      <p className={styles.question}>{question || 'What should we improve first?'}</p>

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}

      {!loading && !error && !voted && (
        <div className={styles.options}>
          {options.map((o) => (
            <button
              key={o}
              type="button"
              className={styles.optionBtn}
              disabled={submitting || !configured}
              onClick={() => onVote(o)}
            >
              {o}
            </button>
          ))}
          {!configured && <span className={styles.error}>No votes list configured yet.</span>}
        </div>
      )}

      {!loading && !error && voted && (
        <div className={styles.results}>
          {options.map((o, i) => {
            const c = countFor(o);
            const p = pct(c);
            const bg = optionColors[i];
            return (
              <div key={o} className={styles.row}>
                <div className={styles.rowHead}>
                  <span className={styles.optName}>{o}</span>
                  <span className={styles.optPct}>{p}%</span>
                </div>
                <div className={styles.bar}>
                  <div className={styles.fill} style={bg ? { width: `${p}%`, background: bg } : { width: `${p}%` }} />
                </div>
              </div>
            );
          })}
          <p className={styles.total}>{total} {total === 1 ? 'vote' : 'votes'}. Thanks for voting.</p>
        </div>
      )}

      {submitError && <p className={styles.error} role="alert">{submitError}</p>}
    </section>
  );
};

export default PollSurvey;
