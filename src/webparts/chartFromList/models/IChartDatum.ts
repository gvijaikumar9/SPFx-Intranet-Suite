export interface IChartDatum {
  label: string;
  value: number;
}

/** Group rows by a category field and either count them or sum a numeric field. */
export function aggregate(
  rows: Record<string, unknown>[],
  categoryField: string,
  valueField: string | undefined,
  maxBars: number
): IChartDatum[] {
  // Null-prototype map so a category named "toString"/"constructor" cannot collide
  // with an inherited member and corrupt the count.
  const map: { [k: string]: number } = Object.create(null);
  const order: string[] = [];
  for (let i = 0; i < rows.length; i++) {
    const raw = rows[i][categoryField];
    const label = (raw === undefined || raw === null || raw === '') ? '(none)' : raw.toString();
    if (map[label] === undefined) { map[label] = 0; order.push(label); }
    if (valueField) {
      const n = Number(rows[i][valueField]);
      map[label] += isNaN(n) ? 0 : n;
    } else {
      map[label] += 1;
    }
  }
  const data: IChartDatum[] = order.map((label) => ({ label: label, value: map[label] }));
  // Largest first, capped to keep the chart readable.
  data.sort((a, b) => b.value - a.value);
  return data.slice(0, Math.max(1, maxBars));
}
