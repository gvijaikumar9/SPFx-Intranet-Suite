export interface IKpi {
  id: number;
  label: string;   // list Title
  value: string;   // KpiValue
  trend: TrendClass;
  delta?: string;  // KpiDelta
}

export type TrendClass = 'up' | 'down' | 'flat';

export function normalizeTrend(value: unknown): TrendClass {
  const v = (value ?? '').toString().trim().toLowerCase();
  if (v.indexOf('up') >= 0 || v.indexOf('rise') >= 0 || v.indexOf('positive') >= 0) { return 'up'; }
  if (v.indexOf('down') >= 0 || v.indexOf('fall') >= 0 || v.indexOf('negative') >= 0) { return 'down'; }
  return 'flat';
}
