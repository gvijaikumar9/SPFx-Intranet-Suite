export interface INews {
  id: number;
  title: string;
  summary?: string;
  category?: string;
  imageUrl?: string;  // empty falls back to a gradient slide
  linkUrl?: string;
}

/**
 * A SharePoint URL field comes back as { Url } or a "url, desc" string. Take the url,
 * and only allow safe schemes (drop javascript:/data:/vbscript:).
 */
export function urlValue(raw: unknown): string {
  if (!raw) { return ''; }
  let url: string;
  if (typeof raw === 'object' && 'Url' in (raw as object)) {
    url = ((raw as { Url?: string }).Url || '').trim();
  } else {
    const s = raw.toString();
    const comma = s.indexOf(',');
    url = (comma >= 0 ? s.substring(0, comma) : s).trim();
  }
  return /^(https?:|mailto:|#|\/)/i.test(url) ? url : '';
}
