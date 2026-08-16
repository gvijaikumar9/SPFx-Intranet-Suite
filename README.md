# SPFx Intranet Suite

Open-source SharePoint Framework (SPFx) web parts that fill the gaps left by the
out-of-the-box modern SharePoint toolbox. These are the small, practical building
blocks most intranets end up rebuilding by hand.

One solution, one package, many web parts. Install once and get the whole set.
SPFx lazy-loads each web part's bundle, so an unused web part costs a page nothing.

## Web parts

| Web part | What it does | List it reads |
| --- | --- | --- |
| **Announcements Ticker** | A rotating band of announcements with Info / Warning / Critical severity. | Announcements |
| **Raise a Ticket** | A help-request form that writes to a list, plus the person's recent tickets. | Tickets |
| **Kudos** | A recognition wall plus a people-picker form to thank a colleague. | Kudos |
| **Employee of the Month** | A spotlight card with the person's profile photo, month and citation. | EmployeeOfMonth |
| **Quick Links** | A grid of shortcut tiles with icons. | QuickLinks |
| **KPI Tiles** | A row of key-number tiles with a trend arrow and delta. | KPIs |
| **FAQ Accordion** | An expandable list of questions and answers. | FAQ |
| **Upcoming Events** | Date-badged cards for events from today onward. | Events (calendar) |
| **Image Gallery** | A photo grid from a document library. | IntranetGallery |

Every web part ships with sample data on by default, so it renders the moment you
drop it on a page. Point it at a list when you are ready.

## Shared options on every web part

- **Layout style**: Card, Minimal, Bold, or Compact, so the same web part can match
  different page designs without forking the code.
- **Container**: show or hide the title, show or hide a border, and set the
  background (transparent, white, light grey, accent tint, or a custom colour).
- **Reset to default**: one button restores every property to its shipped default.
- **Avatar** (where a person shows): choose a profile photo (with initials as a
  fallback) or always show initials.
- **Theme aware**: colours follow the site theme. Messages use a soft, muted tone,
  never a harsh red.

## Requirements

| Tool | Version |
| --- | --- |
| Node.js | v22 LTS (SPFx 1.21.x is not compatible with Node 20) |
| SPFx generator | `@microsoft/generator-sharepoint@1.21.1` |
| SharePoint | SharePoint Online |
| PnP.PowerShell | 3.x (only for the provisioning script) |

## Build and package

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
# -> sharepoint/solution/spfx-intranet-suite.sppkg
```

Upload the `.sppkg` to your tenant App Catalog (or a site-collection App Catalog),
then add the web parts to a page. They appear under the **Intranet Suite** group in
the toolbox.

## Provision the lists

The web parts read from a set of SharePoint lists and one library. The included
script creates them all, with columns and sample data, on any site. It is
idempotent, so existing lists and fields are left alone.

```powershell
.\deploy\Provision-IntranetLists.ps1 -Url "https://contoso.sharepoint.com/sites/intranet"
# schema only, no sample rows:
.\deploy\Provision-IntranetLists.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -NoSampleData
```

## Wire a web part to real data

1. Add the web part to a page. It shows sample data right away.
2. In the property pane, turn **Use sample data** off.
3. Pick the list from the dropdown (for example Announcements, Tickets, KPIs).
4. Where a web part needs field names, they are internal names and are shown in the
   property pane hint. The provisioning script creates matching columns.

Notes per web part:

- **Announcements Ticker**: message field defaults to `Title`; optional Severity
  choice column (`Info` / `Warning` / `Critical`) and a Link column.
- **Raise a Ticket**: writes Title, Category, Description and Status. "My recent
  tickets" is scoped to the current user.
- **Kudos**: the people picker sets the To and From person fields. The wall reads
  those person fields for names and photos.
- **Employee of the Month**: the person comes from the Employee person column. The
  name and photo are read from it; the text Title is not used as a name.
- **Upcoming Events**: reads a calendar list and shows only events from today
  onward.
- **Image Gallery**: point it at a document library and upload images. Only image
  files are shown.

## License

[MIT](LICENSE) 2026 Vijay Gilakattula
