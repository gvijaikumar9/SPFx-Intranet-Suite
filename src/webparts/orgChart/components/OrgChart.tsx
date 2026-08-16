import * as React from 'react';
import { useMemo, useState } from 'react';
import styles from './OrgChart.module.scss';
import { IOrgChartProps } from './IOrgChartProps';
import { IOrgNode } from '../models/IOrgNode';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

function initials(name: string): string {
  const parts = name.trim().split(/\s+/);
  const first = parts[0] ? parts[0][0] : '';
  const last = parts.length > 1 ? parts[parts.length - 1][0] : '';
  return (first + last).toUpperCase() || '?';
}

const Node: React.FC<{ node: IOrgNode; webUrl: string; showPhoto: boolean; emphasis?: boolean }> = (p) => {
  const [failed, setFailed] = useState(false);
  const photoUrl = p.node.email
    ? `${p.webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(p.node.email)}`
    : '';
  const avatarCls = p.emphasis ? `${styles.avatar} ${styles.avatarBig}` : styles.avatar;
  return (
    <div className={`${styles.node} ${p.emphasis ? styles.nodeMe : ''}`}>
      {p.showPhoto && photoUrl && !failed
        ? <img className={avatarCls} src={photoUrl} alt="" aria-hidden="true" onError={() => setFailed(true)} />
        : <span className={avatarCls} aria-hidden="true">{initials(p.node.name)}</span>}
      <span className={styles.nodeName}>{p.node.name}</span>
      {p.node.title && <span className={styles.nodeTitle}>{p.node.title}</span>}
    </div>
  );
};

const OrgChart: React.FC<IOrgChartProps> = (props) => {
  const { title, layout, showTitle, frameStyle, showPhoto, webUrl, isDemo, accent, manager, me, reports, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  return (
    <section className={`${styles.org} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Team'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Deploy to a tenant to show your real reporting line.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && !me && (
        <p className={styles.muted}>No profile information available.</p>
      )}

      {!loading && !error && me && (
        <div className={styles.tree}>
          {manager && (
            <>
              <div className={styles.row}><Node node={manager} webUrl={webUrl} showPhoto={showPhoto} /></div>
              <div className={styles.connector} aria-hidden="true" />
            </>
          )}
          <div className={styles.row}><Node node={me} webUrl={webUrl} showPhoto={showPhoto} emphasis /></div>
          {reports.length > 0 && (
            <>
              <div className={styles.connector} aria-hidden="true" />
              <div className={styles.reports}>
                {reports.map((r, i) => <Node key={r.email || `r${i}`} node={r} webUrl={webUrl} showPhoto={showPhoto} />)}
              </div>
            </>
          )}
        </div>
      )}
    </section>
  );
};

export default OrgChart;
