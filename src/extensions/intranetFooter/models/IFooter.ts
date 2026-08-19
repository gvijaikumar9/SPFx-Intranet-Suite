export interface IFooterLink {
  title: string;
  url: string;
  group: string;   // column heading, e.g. Company / Support / Resources / Legal
  order: number;
}

export interface IFooterGroup {
  name: string;
  links: IFooterLink[];
}

export interface ISocialLink {
  network: string;   // linkedin | twitter | facebook | youtube | instagram | github | email | web
  url: string;
}

// A SharePoint URL field comes back as { Url, Description } or a "url, desc" string.
// Take the url and allow only safe schemes.
export function parseUrl(raw: unknown): string {
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

// Group a flat, ordered link list into columns, first-seen group order preserved.
export function groupLinks(links: IFooterLink[]): IFooterGroup[] {
  const byName: { [k: string]: IFooterGroup } = Object.create(null);
  const order: IFooterGroup[] = [];
  for (let i = 0; i < links.length; i++) {
    const g = links[i].group || 'Links';
    if (!byName[g]) { byName[g] = { name: g, links: [] }; order.push(byName[g]); }
    byName[g].links.push(links[i]);
  }
  return order;
}

// Parse a "network|url, network|url" string into social links (safe schemes only).
export function parseSocial(raw: string | undefined): ISocialLink[] {
  const out: ISocialLink[] = [];
  if (!raw) { return out; }
  const parts = raw.split(',');
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i].trim();
    if (!p) { continue; }
    const bar = p.indexOf('|');
    if (bar < 0) { continue; }
    const network = p.substring(0, bar).trim().toLowerCase();
    const url = parseUrl(p.substring(bar + 1).trim());
    if (network && url) { out.push({ network: network, url: url }); }
  }
  return out;
}
