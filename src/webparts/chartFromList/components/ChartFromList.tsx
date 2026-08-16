import * as React from 'react';
import { useMemo } from 'react';
import styles from './ChartFromList.module.scss';
import { IChartFromListProps } from './IChartFromListProps';
import { IChartDatum } from '../models/IChartDatum';

const LAYOUT_CLASS: Record<string, string | undefined> = {
  card: styles.layoutCard,
  minimal: styles.layoutMinimal,
  bold: styles.layoutBold,
  compact: styles.layoutCompact
};

// Chart geometry (a viewBox that scales to the container).
const PAD_X = 10;
const PAD_TOP = 22;    // room for value labels
const PAD_BOTTOM = 46; // room for category labels
const PLOT_H = 170;
const BAR_W = 46;
const GAP = 24;
const R = 4;           // rounded top corner radius

function trim(label: string): string {
  return label.length > 12 ? `${label.substring(0, 11)}…` : label;
}

const ChartFromList: React.FC<IChartFromListProps> = (props) => {
  const { title, layout, showTitle, frameStyle, isDemo, accent, items, loading, error } = props;
  const style = useMemo(
    () => ({ ...frameStyle, ['--accent' as string]: accent } as React.CSSProperties),
    [accent, frameStyle]
  );
  const layoutClass = LAYOUT_CLASS[layout] || styles.layoutCard;

  const n = items.length;
  let max = 1;
  for (let i = 0; i < n; i++) { if (items[i].value > max) { max = items[i].value; } }

  const width = PAD_X * 2 + n * BAR_W + (n > 0 ? (n - 1) * GAP : 0);
  const height = PAD_TOP + PLOT_H + PAD_BOTTOM;
  const baseY = PAD_TOP + PLOT_H;

  const bar = (d: IChartDatum, i: number): React.ReactElement => {
    const x = PAD_X + i * (BAR_W + GAP);
    const h = Math.max(2, (d.value / max) * PLOT_H);
    const y = baseY - h;
    // Bar with rounded top corners, anchored to the baseline.
    const path = `M${x},${baseY} L${x},${y + R} Q${x},${y} ${x + R},${y} `
      + `L${x + BAR_W - R},${y} Q${x + BAR_W},${y} ${x + BAR_W},${y + R} L${x + BAR_W},${baseY} Z`;
    return (
      <g key={i} className={styles.bar}>
        <title>{`${d.label}: ${d.value}`}</title>
        <path d={path} className={styles.fill} />
        <text x={x + BAR_W / 2} y={y - 7} className={styles.val}>{d.value}</text>
        <text x={x + BAR_W / 2} y={baseY + 18} className={styles.cat}>{trim(d.label)}</text>
      </g>
    );
  };

  const summary = items.map((d) => `${d.label} ${d.value}`).join(', ');

  return (
    <section className={`${styles.chart} ${layoutClass}`} style={style}>
      {showTitle && <h2 className={styles.heading}>{title || 'Chart'}</h2>}
      {isDemo && <p className={styles.demoNote}>Demo mode. Nothing is saved until you point this at a list.</p>}

      {loading && <p className={styles.muted}>Loading...</p>}
      {error && !loading && <p className={styles.error}>{error}</p>}
      {!loading && !error && n === 0 && (
        <p className={styles.muted}>No data to chart.</p>
      )}
      {!loading && !error && n > 0 && (
        <div className={styles.plot}>
          <svg
            viewBox={`0 0 ${width} ${height}`}
            width="100%"
            preserveAspectRatio="xMidYMax meet"
            role="img"
            aria-label={`${title || 'Chart'}: ${summary}`}
          >
            <line x1={PAD_X} y1={baseY} x2={width - PAD_X} y2={baseY} className={styles.axis} />
            {items.map((d, i) => bar(d, i))}
          </svg>
        </div>
      )}
    </section>
  );
};

export default ChartFromList;
