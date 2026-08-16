export type Severity = 'info' | 'warning' | 'critical';

export interface IAnnouncement {
  id: string;
  message: string;
  severity: Severity;
  href?: string;
}

/**
 * Normalise an arbitrary Severity/Choice column value to one of our three levels.
 * Accepts common phrasings so authors are not forced into exact strings.
 */
export function normalizeSeverity(value: unknown): Severity {
  const s = (value ?? '').toString().trim().toLowerCase();
  if (s.indexOf('crit') === 0 || s.indexOf('urgent') >= 0 || s === 'high' || s === 'p1') {
    return 'critical';
  }
  if (s.indexOf('warn') === 0 || s === 'medium' || s === 'p2' || s.indexOf('caution') >= 0) {
    return 'warning';
  }
  return 'info';
}
