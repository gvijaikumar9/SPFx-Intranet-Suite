import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IFeedback } from '../models/IFeedback';

// Write one feedback item to a SharePoint list. Columns: Title (auto summary), Rating
// (number), Comment (note), PageUrl (note), Contact (text). Never throws - returns false
// on any failure so the widget can show a muted retry message.
export async function submitFeedback(client: SPHttpClient, webUrl: string, listTitle: string, fb: IFeedback): Promise<boolean> {
  const safe = (listTitle || 'Feedback').replace(/'/g, "''");
  const url = `${webUrl}/_api/web/lists/getByTitle('${safe}')/items`;
  const summary = (fb.comment || 'Feedback').replace(/\s+/g, ' ').trim();
  const body = JSON.stringify({
    Title: summary.length > 80 ? summary.substring(0, 79) + '…' : (summary || 'Feedback'),
    Rating: fb.rating > 0 ? fb.rating : null,
    Comment: fb.comment || null,
    PageUrl: fb.pageUrl || null,
    Contact: fb.contact || null
  });
  try {
    const res: SPHttpClientResponse = await client.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: body
    });
    return res.ok;
  } catch {
    return false;
  }
}
