export interface ILink {
  id: number;
  label: string;   // list Title
  url: string;     // from the LinkUrl field
  icon?: string;   // Fluent UI icon name (IconName field)
}

/** A SharePoint URL field comes back as { Url, Description } (or a "url, desc" string). Take the url. */
export function parseUrlField(raw: unknown): string {
  if (!raw) { return ''; }
  if (typeof raw === 'object' && 'Url' in (raw as object)) {
    return ((raw as { Url?: string }).Url || '').trim();
  }
  const s = raw.toString();
  const comma = s.indexOf(',');
  return (comma >= 0 ? s.substring(0, comma) : s).trim();
}
