export interface ITicket {
  id: number;
  title: string;
  category?: string;
  status?: string;
}

/** Map a raw status string to one of three display buckets so the pill colour is predictable. */
export type StatusClass = 'new' | 'progress' | 'resolved';

export function classifyStatus(value: unknown): StatusClass {
  const v = (value ?? '').toString().trim().toLowerCase();
  if (v.indexOf('progress') >= 0 || v.indexOf('open') >= 0 || v.indexOf('active') >= 0) { return 'progress'; }
  if (v.indexOf('resolved') >= 0 || v.indexOf('closed') >= 0 || v.indexOf('done') >= 0) { return 'resolved'; }
  return 'new';
}
