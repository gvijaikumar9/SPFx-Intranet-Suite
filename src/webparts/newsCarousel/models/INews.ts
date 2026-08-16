export interface INews {
  id: number;
  title: string;
  summary?: string;
  category?: string;
  imageUrl?: string;  // empty falls back to a gradient slide
  linkUrl?: string;
}

/** A SharePoint URL field comes back as { Url } or a "url, desc" string. Take the url. */
export function urlValue(raw: unknown): string {
  if (!raw) { return ''; }
  if (typeof raw === 'object' && 'Url' in (raw as object)) {
    return ((raw as { Url?: string }).Url || '').trim();
  }
  const s = raw.toString();
  const comma = s.indexOf(',');
  return (comma >= 0 ? s.substring(0, comma) : s).trim();
}
