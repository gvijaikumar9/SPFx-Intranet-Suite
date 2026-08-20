export interface IHeaderLink {
  title: string;
  url: string;
  order: number;
}

// The effective, resolved header settings the band renders from. Built by layering the config
// list (owner-editable, wins) over the deploy-script properties (defaults).
export interface IHeaderSettings {
  enabled: boolean;
  hideNativeHeader: boolean;
  brandText: string;
  logoUrl: string;
  homeUrl: string;
  showCta: boolean;
  ctaText: string;
  ctaUrl: string;
  accent: string;
}

// A partial patch owners save from the settings panel; only the fields they change are sent.
export interface IHeaderConfigPatch {
  enabled?: boolean;
  hideNativeHeader?: boolean;
  brandText?: string;
  showCta?: boolean;
  ctaText?: string;
  ctaUrl?: string;
  accent?: string;
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

// Normalise a URL to a path for active-link matching. An absolute URL keeps its pathname
// only; a site-relative url is returned as-is. Trailing slash removed, lower-cased.
export function toPath(url: string): string {
  if (!url) { return ''; }
  let path = url;
  const scheme = url.indexOf('://');
  if (scheme >= 0) {
    const rest = url.substring(scheme + 3);
    const slash = rest.indexOf('/');
    path = slash >= 0 ? rest.substring(slash) : '/';
  }
  const q = path.indexOf('?');
  if (q >= 0) { path = path.substring(0, q); }
  const h = path.indexOf('#');
  if (h >= 0) { path = path.substring(0, h); }
  if (path.length > 1 && path.charAt(path.length - 1) === '/') { path = path.substring(0, path.length - 1); }
  return path.toLowerCase();
}

// Is a nav link the current page? Exact match, or the current path sits under the link's
// path (so a section link stays highlighted on its child pages). The site root ('/','') is
// only ever an exact match, never a prefix, so it does not light up on every page.
export function isActive(linkUrl: string, currentPath: string): boolean {
  const link = toPath(linkUrl);
  const here = toPath(currentPath);
  if (!link || !here) { return false; }
  if (link === here) { return true; }
  if (link === '/' || link.length < 2) { return false; }
  return here.indexOf(link + '/') === 0;
}
