import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IHeaderConfigPatch } from '../models/IHeader';

// The header's owner-editable settings live in a single-item list (default 'HeaderConfig').
// Columns: Title (row key), HdrEnabled (Yes/No), HdrHideNative (Yes/No), HdrBrand (text),
// HdrCtaText (text), HdrCtaUrl (text), HdrAccent (text). All fields are optional - a missing
// column or a missing list just means "no override", and the deploy-script defaults apply.

export interface IHeaderConfigRow {
  id: number;                 // 0 = no row yet
  enabled?: boolean;
  hideNativeHeader?: boolean;
  brandText?: string;
  showCta?: boolean;
  ctaText?: string;
  ctaUrl?: string;
  accent?: string;
}

function asBool(v: unknown): boolean | undefined {
  if (v === true || v === false) { return v; }
  if (v === 1 || v === '1') { return true; }
  if (v === 0 || v === '0') { return false; }
  return undefined;
}

function str(v: unknown): string | undefined {
  if (v === null || v === undefined) { return undefined; }
  const s = v.toString();
  return s.length > 0 ? s : undefined;
}

// Read the single config row. Never throws - on any failure returns an empty row (id 0), so
// the header falls back to its deploy-script defaults and never breaks the page.
export async function getHeaderConfig(client: SPHttpClient, webUrl: string, listTitle: string): Promise<IHeaderConfigRow> {
  const safe = (listTitle || 'HeaderConfig').replace(/'/g, "''");
  // deterministic: always read the lowest-id row, so reads are stable if a stray row ever exists
  const url = `${webUrl}/_api/web/lists/getByTitle('${safe}')/items?$top=1&$orderby=Id`;
  try {
    const res: SPHttpClientResponse = await client.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) { return { id: 0 }; }
    const json = await res.json();
    const rows = (json.value || []) as Record<string, unknown>[];
    if (rows.length === 0) { return { id: 0 }; }
    const r = rows[0];
    return {
      id: typeof r.Id === 'number' ? (r.Id as number) : 0,
      enabled: asBool(r.HdrEnabled),
      hideNativeHeader: asBool(r.HdrHideNative),
      brandText: str(r.HdrBrand),
      showCta: asBool(r.HdrShowCta),
      ctaText: str(r.HdrCtaText),
      ctaUrl: str(r.HdrCtaUrl),
      accent: str(r.HdrAccent)
    };
  } catch {
    return { id: 0 };
  }
}

// Write the owner's changes. Updates the existing row (MERGE) or creates it if none exists.
// Returns false on any failure (e.g. the caller lacks write access), so the panel can show an
// error rather than throw. SPHttpClient adds the request digest automatically.
export async function saveHeaderConfig(
  client: SPHttpClient,
  webUrl: string,
  listTitle: string,
  existingId: number,
  patch: IHeaderConfigPatch
): Promise<boolean> {
  const safe = (listTitle || 'HeaderConfig').replace(/'/g, "''");
  const base = `${webUrl}/_api/web/lists/getByTitle('${safe}')`;

  const body: Record<string, unknown> = {};
  if (patch.enabled !== undefined) { body.HdrEnabled = patch.enabled; }
  if (patch.hideNativeHeader !== undefined) { body.HdrHideNative = patch.hideNativeHeader; }
  if (patch.brandText !== undefined) { body.HdrBrand = patch.brandText; }
  if (patch.showCta !== undefined) { body.HdrShowCta = patch.showCta; }
  if (patch.ctaText !== undefined) { body.HdrCtaText = patch.ctaText; }
  if (patch.ctaUrl !== undefined) { body.HdrCtaUrl = patch.ctaUrl; }
  if (patch.accent !== undefined) { body.HdrAccent = patch.accent; }

  // 'odata-version': '' is required: SPFx otherwise sends odata-version 4.0, which SharePoint's
  // nometadata REST (OData v3) rejects with a 400. Clearing it makes MERGE/POST writes succeed.
  try {
    // If we do not have a row id, check whether one already exists before creating, so a lost id
    // (e.g. a failed re-read on the caller side) can never create a duplicate config row.
    let id = existingId;
    if (id <= 0) {
      const found = await getHeaderConfig(client, webUrl, listTitle);
      if (found.id > 0) { id = found.id; }
    }
    if (id > 0) {
      const res = await client.post(`${base}/items(${id})`, SPHttpClient.configurations.v1, {
        headers: {
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'Content-Type': 'application/json;odata=nometadata',
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(body)
      });
      return res.ok || res.status === 204;
    }
    // no row yet: create one (seed a Title so the row is easy to find in the list)
    body.Title = 'Header';
    const res = await client.post(`${base}/items`, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    });
    return res.ok;
  } catch {
    return false;
  }
}
