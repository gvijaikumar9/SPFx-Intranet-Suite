import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IHeaderLink, parseUrl } from '../models/IHeader';

// Read header nav links from a SharePoint list. Columns: Title (link text), HeaderUrl (URL),
// HeaderOrder (number). Never throws - a header must not break the page, so on any failure
// it returns an empty list and the band falls back to just the brand.
export async function getHeaderLinks(client: SPHttpClient, webUrl: string, listTitle: string): Promise<IHeaderLink[]> {
  const safe = (listTitle || 'HeaderLinks').replace(/'/g, "''");
  // No $select: the list may lack an optional column, and $select on a missing field is a
  // hard 400. We read the fields defensively from each row instead.
  const url = `${webUrl}/_api/web/lists/getByTitle('${safe}')/items?$top=100`;
  try {
    const res: SPHttpClientResponse = await client.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return []; }
    const json = await res.json();
    const rows = (json.value || []) as Record<string, unknown>[];
    const links: IHeaderLink[] = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const title = (r.Title || '').toString();
      const linkUrl = parseUrl(r.HeaderUrl);
      if (!title || !linkUrl) { continue; }
      links.push({
        title: title,
        url: linkUrl,
        order: typeof r.HeaderOrder === 'number' ? (r.HeaderOrder as number) : 999
      });
    }
    links.sort((a, b) => a.order - b.order);
    return links;
  } catch {
    return [];
  }
}
