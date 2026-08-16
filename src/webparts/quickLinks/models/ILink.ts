export interface ILink {
  id: number;
  label: string;   // list Title
  url: string;     // from the LinkUrl field
  icon?: string;   // Fluent UI icon name (IconName field)
}

/**
 * A SharePoint URL field comes back as { Url, Description } (or a "url, desc" string).
 * Take the url, and only allow safe schemes (drop javascript:/data:/vbscript:).
 */
export function parseUrlField(raw: unknown): string {
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
