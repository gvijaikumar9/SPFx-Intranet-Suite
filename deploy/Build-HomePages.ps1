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
  [switch]$RealData,
  [switch]$Card,   # prototype look: neutral section backgrounds + white bordered web-part cards
  [ValidateSet('default', 'spotlight', 'cockpit')][string]$Layout = 'default'   # composition preset (which web parts go where)
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
# Cascade prototype layout: a full-width ticker, then a 2/3 main column + 1/3 sidebar.
# TwoColumnLeft = wide left column (2/3, main) + narrow right column (1/3, sidebar).
# If the widths come out reversed on your tenant, swap to 'TwoColumnRight'.
$sections = @(
  'OneColumn',       # 1  announcements ticker, full width
  'TwoColumnLeft'    # 2  main (col 1, 2/3) + sidebar (col 2, 1/3)
)

# Spotlight composition (matches prototype-polished.html): a personalised greeting
# strip, a hero + latest-news row, an app-tile row, a what's-on + people-spotlight
# row, a pulse-stats strip, then a full-width learning band. Only the ~8 web parts
# the prototype actually shows, in the prototype's order.
$spotlightSections = @(
  'OneColumn',       # 1  greeting strip
  'TwoColumnLeft',   # 2  news hero (2/3) + latest news list (1/3)
  'OneColumn',       # 3  jump back in - app tiles
  'TwoColumnLeft',   # 4  what's on (2/3) + people spotlight (1/3)
  'OneColumn'        # 5  pulse stats
)

# Cockpit composition (matches prototype-compact.html): a dense operational dashboard.
# Full-width alerts, hero and KPI row across the top, then a packed three-column deck
# of small cards. The compact base variant tightens every card.
$cockpitSections = @(
  'OneColumn',       # 1  alerts ticker
  'OneColumn',       # 2  news hero
  'OneColumn',       # 3  KPI stat tiles
  'ThreeColumn'      # 4  the dense card deck
)

# --- web part placements. Multiple web parts in the same Section+Column stack by
#     Order. List = provisioned list; LiveNoList = live service (turn demo off);
#     Extra = extra properties to set on real data. Variant = per-web-part layout
#     override (universal across templates, from the prototype reconciliation);
#     web parts with no Variant use the site-wide -Variant base. ---
$plan = @(
  @{ Title = 'Announcements Ticker';   Sec = 1; Col = 1; Order = 1; Variant = 'bold'; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },

  # main column (2/3)
  @{ Title = 'News Carousel';          Sec = 2; Col = 1; Order = 1;  Variant = 'bold';    List = 'News' },
  @{ Title = 'Tabs';                   Sec = 2; Col = 1; Order = 2 },
  @{ Title = 'KPI Tiles';              Sec = 2; Col = 1; Order = 3;  Variant = 'minimal'; List = 'KPIs' },
  @{ Title = 'Chart from a List';      Sec = 2; Col = 1; Order = 4;  List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Kudos';                  Sec = 2; Col = 1; Order = 5;  List = 'Kudos' },
  @{ Title = 'Raise a Ticket';         Sec = 2; Col = 1; Order = 6;  List = 'Tickets' },
  @{ Title = 'FAQ Accordion';          Sec = 2; Col = 1; Order = 7;  List = 'FAQ' },
  @{ Title = 'Poll';                   Sec = 2; Col = 1; Order = 8;  List = 'PollVotes' },
  @{ Title = 'Content Rollup';         Sec = 2; Col = 1; Order = 9;  List = 'News' },
  @{ Title = 'Image Gallery';          Sec = 2; Col = 1; Order = 10; List = 'IntranetGallery' },

  # sidebar (1/3)
  @{ Title = 'Weather';                Sec = 2; Col = 2; Order = 1;  Extra = @{ location = 'London' } },
  @{ Title = 'Employee of the Month';  Sec = 2; Col = 2; Order = 2;  Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Upcoming Events';        Sec = 2; Col = 2; Order = 3;  List = 'Events' },
  @{ Title = 'Celebrations';           Sec = 2; Col = 2; Order = 4;  List = 'Celebrations' },
  @{ Title = 'Quick Links';            Sec = 2; Col = 2; Order = 5;  List = 'QuickLinks' },
  @{ Title = 'Event Countdown';        Sec = 2; Col = 2; Order = 6 },
  @{ Title = 'Upcoming Holidays';      Sec = 2; Col = 2; Order = 7;  List = 'Holidays' },
  @{ Title = 'Org Chart';              Sec = 2; Col = 2; Order = 8;  LiveNoList = $true },
  @{ Title = 'People Directory';       Sec = 2; Col = 2; Order = 9;  LiveNoList = $true }
)

# Spotlight plan: the prototype-polished composition. Container overrides keep the
# greeting band tinted (not a white card) and let the hero/learning bands sit flush.
$spotlightPlan = @(
  # 1  greeting strip - transparent so the web part's own tinted band shows through
  @{ Title = 'Greeting';               Sec = 1; Col = 1; Order = 1; Variant = 'minimal'; Container = @{ showBorder = $false; backgroundMode = 'transparent' }; Extra = @{ chipsText = '3|Tasks due, 2|Approvals, |Next: Standup 10:30' } },

  # 2  hero + latest news
  @{ Title = 'News Carousel';          Sec = 2; Col = 1; Order = 1; Variant = 'bold'; List = 'News'; Container = @{ showBorder = $false; backgroundMode = 'transparent' } },
  @{ Title = 'Content Rollup';         Sec = 2; Col = 2; Order = 1; List = 'News' },

  # 3  jump back in - app tiles
  @{ Title = 'Quick Links';            Sec = 3; Col = 1; Order = 1; List = 'QuickLinks' },

  # 4  what's on + people spotlight
  @{ Title = 'Upcoming Events';        Sec = 4; Col = 1; Order = 1; List = 'Events' },
  @{ Title = 'Employee of the Month';  Sec = 4; Col = 2; Order = 1; Variant = 'bold'; List = 'EmployeeOfMonth' },

  # 5  pulse stats
  @{ Title = 'KPI Tiles';              Sec = 5; Col = 1; Order = 1; Variant = 'minimal'; List = 'KPIs' }
)

# Cockpit plan: full-width alerts + hero + KPI, then a balanced three-column deck.
# Base variant is compact (dense); Ticker/News=bold, KPI=minimal, Employee=bold are
# the universal per-web-part overrides. Content Rollup reads the News list.
$cockpitPlan = @(
  # full-width top band
  @{ Title = 'Announcements Ticker';   Sec = 1; Col = 1; Order = 1; Variant = 'bold';    List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  @{ Title = 'News Carousel';          Sec = 2; Col = 1; Order = 1; Variant = 'bold';    List = 'News' },
  @{ Title = 'KPI Tiles';              Sec = 3; Col = 1; Order = 1; Variant = 'minimal'; List = 'KPIs' },

  # deck column 1
  @{ Title = 'Tabs';                   Sec = 4; Col = 1; Order = 1 },
  @{ Title = 'Content Rollup';         Sec = 4; Col = 1; Order = 2; List = 'News' },
  @{ Title = 'Kudos';                  Sec = 4; Col = 1; Order = 3; List = 'Kudos' },
  @{ Title = 'Celebrations';           Sec = 4; Col = 1; Order = 4; List = 'Celebrations' },
  @{ Title = 'Raise a Ticket';         Sec = 4; Col = 1; Order = 5; List = 'Tickets' },
  @{ Title = 'Upcoming Holidays';      Sec = 4; Col = 1; Order = 6; List = 'Holidays' },

  # deck column 2
  @{ Title = 'Weather';                Sec = 4; Col = 2; Order = 1; Extra = @{ location = 'Houston' } },
  @{ Title = 'Event Countdown';        Sec = 4; Col = 2; Order = 2 },
  @{ Title = 'Upcoming Events';        Sec = 4; Col = 2; Order = 3; List = 'Events' },
  @{ Title = 'Quick Links';            Sec = 4; Col = 2; Order = 4; List = 'QuickLinks' },
  @{ Title = 'Image Gallery';          Sec = 4; Col = 2; Order = 5; List = 'IntranetGallery' },
  @{ Title = 'FAQ Accordion';          Sec = 4; Col = 2; Order = 6; List = 'FAQ' },

  # deck column 3
  @{ Title = 'Chart from a List';      Sec = 4; Col = 3; Order = 1; List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Employee of the Month';  Sec = 4; Col = 3; Order = 2; Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Poll';                   Sec = 4; Col = 3; Order = 3; List = 'PollVotes' },
  @{ Title = 'Org Chart';              Sec = 4; Col = 3; Order = 4; LiveNoList = $true },
  @{ Title = 'People Directory';       Sec = 4; Col = 3; Order = 5; LiveNoList = $true }
)

# pick the composition preset
if ($Layout -eq 'spotlight') {
  $sections = $spotlightSections
  $plan = $spotlightPlan
  Write-Host "  composition preset: spotlight ($($sections.Count) sections, $($plan.Count) web parts)" -ForegroundColor DarkGray
}
elseif ($Layout -eq 'cockpit') {
  $sections = $cockpitSections
  $plan = $cockpitPlan
  Write-Host "  composition preset: cockpit ($($sections.Count) sections, $($plan.Count) web parts)" -ForegroundColor DarkGray
}

# Delete + recreate the page: a corrupted CanvasContent1 cannot be fixed in place.
try {
  Remove-PnPPage -Identity "Home.aspx" -Force -ErrorAction Stop
  Write-Host "  removed existing Home.aspx"
} catch {
  Write-Host ("  (no page to remove: {0})" -f $_.Exception.Message) -ForegroundColor DarkGray
}

# Home layout is header-less (no big title banner), matching the prototype which
# starts straight into content.
$null = Add-PnPPage -Name "Home.aspx" -LayoutType Home
for ($i = 0; $i -lt $sections.Count; $i++) {
  if ($Card) {
    Add-PnPPageSection -Page "Home.aspx" -SectionTemplate $sections[$i] -Order ($i + 1) -ZoneEmphasis 1 | Out-Null
  } else {
    Add-PnPPageSection -Page "Home.aspx" -SectionTemplate $sections[$i] -Order ($i + 1) | Out-Null
  }
}
Write-Host "  created $($sections.Count) sections"

$added = 0
foreach ($wp in $plan) {
  # per-web-part variant override wins over the site-wide base variant
  $wpVariant = if ($wp.Variant) { $wp.Variant } else { $Variant }
  $props = @{ layout = $wpVariant; showTitle = $true }
  if ($Card) { $props.showBorder = $true; $props.backgroundMode = 'white' }
  # per-web-part container override wins over -Card (e.g. flush hero / tinted greeting band)
  if ($wp.Container) { foreach ($c in $wp.Container.GetEnumerator()) { $props[$c.Key] = $c.Value } }
  if ($RealData) {
    if ($wp.List) { $props.useDemoData = $false; $props.listTitle = $wp.List }
    elseif ($wp.LiveNoList) { $props.useDemoData = $false }
    if ($wp.Extra) { foreach ($e in $wp.Extra.GetEnumerator()) { $props[$e.Key] = $e.Value } }
  }
  $json = $props | ConvertTo-Json -Compress
  $args = @{ Page = "Home.aspx"; Component = $wp.Title; Section = $wp.Sec; Column = $wp.Col; Order = $wp.Order; ErrorAction = 'Stop' }
  if ($propParam) { $args[$propParam] = $json }
  try {
    Add-PnPPageWebPart @args | Out-Null
    $added++
    Write-Host ("  + {0}" -f $wp.Title) -ForegroundColor Green
  } catch {
    Write-Host ("  ! could not add {0}: {1}" -f $wp.Title, $_.Exception.Message) -ForegroundColor DarkYellow
  }
}

# A home page does not need the per-page comments box.
try { Set-PnPPage -Identity "Home.aspx" -CommentsEnabled:$false -ErrorAction Stop | Out-Null } catch {}
Set-PnPPage -Identity "Home.aspx" -Publish | Out-Null
try { Set-PnPHomePage -RootFolderRelativeUrl "SitePages/Home.aspx" -ErrorAction Stop } catch {}

Write-Host "Done. Placed $added of $($plan.Count) web parts. Reload the site home." -ForegroundColor Cyan
