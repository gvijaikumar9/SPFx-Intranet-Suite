import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IFooterLink, parseUrl } from '../models/IFooter';

// Read footer links from a SharePoint list. Columns: Title (link text), FooterUrl (URL),
// FooterGroup (text column heading), FooterOrder (number). Never throws - a footer must
// not break the page, so on any failure it returns an empty list.
export async function getFooterLinks(client: SPHttpClient, webUrl: string, listTitle: string): Promise<IFooterLink[]> {
  const safe = (listTitle || 'FooterLinks').replace(/'/g, "''");
  // No $select: the list may lack one of the optional columns, and $select on a missing
  // field is a hard 400. We read the fields defensively from each row instead.
  const url = `${webUrl}/_api/web/lists/getByTitle('${safe}')/items?$top=100`;
  try {
    const res: SPHttpClientResponse = await client.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return []; }
    const json = await res.json();
    const rows = (json.value || []) as Record<string, unknown>[];
    const links: IFooterLink[] = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const title = (r.Title || '').toString();
      const linkUrl = parseUrl(r.FooterUrl);
      if (!title || !linkUrl) { continue; }
      links.push({
        title: title,
        url: linkUrl,
        group: r.FooterGroup ? r.FooterGroup.toString() : 'Links',
        order: typeof r.FooterOrder === 'number' ? (r.FooterOrder as number) : 999
      });
    }
    links.sort((a, b) => a.order - b.order);
    return links;
  } catch {
    return [];
  }
}
