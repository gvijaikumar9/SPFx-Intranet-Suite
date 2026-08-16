<#
.SYNOPSIS
  Builds a template site's Home.aspx from scratch, placing all 16 Intranet Suite
  web parts into a clean multi-section layout. Optionally binds each web part to its
  real list and applies a layout variant (Card / Minimal / Bold / Compact).

  The suite package must already be deployed to the site (Add-PnPApp) and, for real
  data, the lists must exist (run Provision-IntranetLists.ps1 first).

.PARAMETER Variant
  The layout style applied to every web part: card | minimal | bold | compact.
  Give each template site a different variant to make them look distinct.

.PARAMETER RealData
  Bind each web part to its provisioned list and turn sample data off. Omit to leave
  every web part in demo mode.

.EXAMPLE
  .\deploy\Build-HomePages.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" -Variant card -RealData

.EXAMPLE
  # All five template sites, each with its own variant, live data:
  $cid = "<app-id>"
  $map = @{ cascade='card'; momentum='bold'; cockpit='compact'; squad='minimal'; spotlight='bold' }
  foreach ($k in $map.Keys) {
    .\deploy\Build-HomePages.ps1 -Url "https://contoso.sharepoint.com/sites/intranet-$k" -ClientId $cid -Variant $map[$k] -RealData
  }
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [ValidateSet('card', 'minimal', 'bold', 'compact')][string]$Variant = 'card',
  [switch]$RealData
)
$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
Write-Host "Building Home.aspx on $Url (variant: $Variant, realData: $($RealData.IsPresent))" -ForegroundColor Cyan

# The property parameter on Add-PnPPageWebPart is named differently across PnP versions.
$wpParams = (Get-Command Add-PnPPageWebPart).Parameters.Keys
$propParam = $null
if ($wpParams -contains 'PropertiesJson') { $propParam = 'PropertiesJson' }
elseif ($wpParams -contains 'WebPartProperties') { $propParam = 'WebPartProperties' }
if ($propParam) { Write-Host "  using property parameter: -$propParam" -ForegroundColor DarkGray }
else { Write-Host ("  no known property parameter; placement only. Available params: {0}" -f ($wpParams -join ', ')) -ForegroundColor DarkYellow }

# --- section templates, top to bottom ---
# NOTE: OneColumnFullWidth only accepts image/hero web parts, so custom web parts
# use OneColumn (still spans the content width) for the top bands.
$sections = @(
  'OneColumn',           # 1  hero band
  'OneColumn',           # 2  news hero
  'OneColumn',           # 3  KPI band
  'TwoColumn',           # 4  kudos | ticket
  'ThreeColumn',         # 5  links | employee | events
  'TwoColumn',           # 6  faq | poll
  'ThreeColumn',         # 7  countdown | holidays | org
  'TwoColumn',           # 8  people | rollup
  'OneColumn'            # 9  gallery
)

# --- web part placements. List = provisioned list; LiveNoList = has no list but a
#     live service (turn demo off); Extra = extra properties to set on real data. ---
$plan = @(
  @{ Title = 'Announcements Ticker';   Sec = 1; Col = 1; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  @{ Title = 'News Carousel';          Sec = 2; Col = 1; List = 'News' },
  @{ Title = 'KPI Tiles';              Sec = 3; Col = 1; List = 'KPIs' },
  @{ Title = 'Kudos';                  Sec = 4; Col = 1; List = 'Kudos' },
  @{ Title = 'Raise a Ticket';         Sec = 4; Col = 2; List = 'Tickets' },
  @{ Title = 'Quick Links';            Sec = 5; Col = 1; List = 'QuickLinks' },
  @{ Title = 'Employee of the Month';  Sec = 5; Col = 2; List = 'EmployeeOfMonth' },
  @{ Title = 'Upcoming Events';        Sec = 5; Col = 3; List = 'Events' },
  @{ Title = 'FAQ Accordion';          Sec = 6; Col = 1; List = 'FAQ' },
  @{ Title = 'Poll';                   Sec = 6; Col = 2; List = 'PollVotes' },
  @{ Title = 'Event Countdown';        Sec = 7; Col = 1 },
  @{ Title = 'Upcoming Holidays';      Sec = 7; Col = 2; List = 'Holidays' },
  @{ Title = 'Org Chart';              Sec = 7; Col = 3; LiveNoList = $true },
  @{ Title = 'People Directory';       Sec = 8; Col = 1; LiveNoList = $true },
  @{ Title = 'Content Rollup';         Sec = 8; Col = 2; LiveNoList = $true },
  @{ Title = 'Image Gallery';          Sec = 9; Col = 1; List = 'IntranetGallery' }
)

# Delete + recreate the page: a corrupted CanvasContent1 cannot be fixed in place.
try {
  Remove-PnPPage -Identity "Home.aspx" -Force -ErrorAction Stop
  Write-Host "  removed existing Home.aspx"
} catch {
  Write-Host ("  (no page to remove: {0})" -f $_.Exception.Message) -ForegroundColor DarkGray
}

$null = Add-PnPPage -Name "Home.aspx" -LayoutType Article
for ($i = 0; $i -lt $sections.Count; $i++) {
  Add-PnPPageSection -Page "Home.aspx" -SectionTemplate $sections[$i] -Order ($i + 1) | Out-Null
}
Write-Host "  created $($sections.Count) sections"

$added = 0
foreach ($wp in $plan) {
  $props = @{ layout = $Variant; showTitle = $true }
  if ($RealData) {
    if ($wp.List) { $props.useDemoData = $false; $props.listTitle = $wp.List }
    elseif ($wp.LiveNoList) { $props.useDemoData = $false }
    if ($wp.Extra) { foreach ($e in $wp.Extra.GetEnumerator()) { $props[$e.Key] = $e.Value } }
  }
  $json = $props | ConvertTo-Json -Compress
  $args = @{ Page = "Home.aspx"; Component = $wp.Title; Section = $wp.Sec; Column = $wp.Col; ErrorAction = 'Stop' }
  if ($propParam) { $args[$propParam] = $json }
  try {
    Add-PnPPageWebPart @args | Out-Null
    $added++
    Write-Host ("  + {0}" -f $wp.Title) -ForegroundColor Green
  } catch {
    Write-Host ("  ! could not add {0}: {1}" -f $wp.Title, $_.Exception.Message) -ForegroundColor DarkYellow
  }
}

Set-PnPPage -Identity "Home.aspx" -Publish | Out-Null
try { Set-PnPHomePage -RootFolderRelativeUrl "SitePages/Home.aspx" -ErrorAction Stop } catch {}

Write-Host "Done. Placed $added of $($plan.Count) web parts. Reload the site home." -ForegroundColor Cyan
