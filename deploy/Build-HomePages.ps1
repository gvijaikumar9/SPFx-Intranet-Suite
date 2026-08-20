<#
.SYNOPSIS
  Builds a template site's Home.aspx from scratch, placing all 21 Intranet Suite web
  parts into a balanced multi-section layout. Each composition preset (spotlight / cockpit
  / momentum / squad / cascade) carries the FULL web-part set but in a distinct structure,
  and -TileColors gives each site its own element-colour palette, so every template looks
  unique while showing everything.

  The suite package must already be deployed to the site (Add-PnPApp) and, for real data,
  the lists must exist (run Provision-IntranetLists.ps1 first).

.PARAMETER Variant
  The base layout style for every web part: card | minimal | bold | compact. Per-web-part
  overrides in the plan win over this.

.PARAMETER TileColors
  Per-element colour palette applied to the colour-aware web parts (KPI Tiles, Quick Links,
  Chart, Event Countdown, Tabs, Poll): none | palette (soft) | fluro | water | dark. Give
  each template a different palette to make the sites distinct.

.PARAMETER RealData
  Bind each web part to its provisioned list and turn sample data off.

.EXAMPLE
  .\deploy\Build-HomePages.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" -Variant card -RealData -Layout cascade -TileColors palette
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [ValidateSet('card', 'minimal', 'bold', 'compact')][string]$Variant = 'card',
  [switch]$RealData,
  [switch]$Card,   # prototype look: neutral section backgrounds + white bordered web-part cards
  [ValidateSet('default', 'spotlight', 'cockpit', 'momentum', 'squad', 'cascade')][string]$Layout = 'default',   # composition preset (which web parts go where)
  [ValidateSet('none', 'palette', 'fluro', 'water', 'dark')][string]$TileColors = 'none'   # per-element colour palette applied to KPI/Quick Links/Chart/Countdown/Tabs/Poll
)
$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
Write-Host "Building Home.aspx on $Url (variant: $Variant, layout: $Layout, tileColors: $TileColors, realData: $($RealData.IsPresent))" -ForegroundColor Cyan

# The property parameter on Add-PnPPageWebPart is named differently across PnP versions.
$wpParams = (Get-Command Add-PnPPageWebPart).Parameters.Keys
$propParam = $null
if ($wpParams -contains 'PropertiesJson') { $propParam = 'PropertiesJson' }
elseif ($wpParams -contains 'WebPartProperties') { $propParam = 'WebPartProperties' }
if ($propParam) { Write-Host "  using property parameter: -$propParam" -ForegroundColor DarkGray }
else { Write-Host ("  no known property parameter; placement only. Available params: {0}" -f ($wpParams -join ', ')) -ForegroundColor DarkYellow }

# web parts that support the shared tile/element colour palette (tileColorMode property)
$tileColorWebParts = @('KPI Tiles', 'Quick Links', 'Chart from a List', 'Event Countdown', 'Tabs', 'Poll')

# Greeting + Announcements are the same full-width top of every preset; define once.
$greeting = @{ Title = 'Greeting'; Sec = 1; Col = 1; Order = 1; Variant = 'minimal'; Container = @{ showBorder = $false; backgroundMode = 'transparent' }; Extra = @{ chipsText = '3|Tasks due, 2|Approvals, |Next: Standup 10:30' } }

# ============================================================================================
# SECTION TEMPLATES + WEB-PART PLANS. Every preset carries all 21 web parts, distributed so
# columns stay balanced (comparable count + tall web parts spread across columns). List =
# provisioned list; LiveNoList = live service; Extra = extra props; Variant/Container = overrides.
# ============================================================================================

# ---- CASCADE: classic comms. Greeting/Alerts/News/KPI stacked, then a balanced 3-col deck. ----
$cascadeSections = @('OneColumn', 'OneColumn', 'OneColumn', 'OneColumn', 'ThreeColumn')
$cascadePlan = @(
  $greeting,
  @{ Title = 'Announcements Ticker'; Sec = 2; Col = 1; Order = 1; Variant = 'bold'; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  @{ Title = 'News Carousel'; Sec = 3; Col = 1; Order = 1; Variant = 'bold'; List = 'News' },
  @{ Title = 'KPI Tiles'; Sec = 4; Col = 1; Order = 1; Variant = 'minimal'; List = 'KPIs' },
  # deck col 1
  @{ Title = 'Chart from a List'; Sec = 5; Col = 1; Order = 1; List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Content Rollup'; Sec = 5; Col = 1; Order = 2; List = 'News' },
  @{ Title = 'Employee of the Month'; Sec = 5; Col = 1; Order = 3; Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Kudos'; Sec = 5; Col = 1; Order = 4; List = 'Kudos' },
  @{ Title = 'Weather'; Sec = 5; Col = 1; Order = 5; Extra = @{ location = 'London' } },
  @{ Title = 'Upcoming Holidays'; Sec = 5; Col = 1; Order = 6; List = 'Holidays' },
  # deck col 2
  @{ Title = 'Tabs'; Sec = 5; Col = 2; Order = 1 },
  @{ Title = 'Image Gallery'; Sec = 5; Col = 2; Order = 2; List = 'IntranetGallery' },
  @{ Title = 'Poll'; Sec = 5; Col = 2; Order = 3; List = 'PollVotes' },
  @{ Title = 'Raise a Ticket'; Sec = 5; Col = 2; Order = 4; List = 'Tickets' },
  @{ Title = 'Quick Links'; Sec = 5; Col = 2; Order = 5; List = 'QuickLinks' },
  @{ Title = 'Celebrations'; Sec = 5; Col = 2; Order = 6; List = 'Celebrations' },
  # deck col 3
  @{ Title = 'Org Chart'; Sec = 5; Col = 3; Order = 1; LiveNoList = $true },
  @{ Title = 'FAQ Accordion'; Sec = 5; Col = 3; Order = 2; List = 'FAQ' },
  @{ Title = 'Upcoming Events'; Sec = 5; Col = 3; Order = 3; List = 'Events' },
  @{ Title = 'People Directory'; Sec = 5; Col = 3; Order = 4; LiveNoList = $true },
  @{ Title = 'Event Countdown'; Sec = 5; Col = 3; Order = 5 }
)

# ---- COCKPIT: dense dashboard. Same stacked top, then a packed 3-col deck (compact base). ----
$cockpitSections = @('OneColumn', 'OneColumn', 'OneColumn', 'OneColumn', 'ThreeColumn')
$cockpitPlan = @(
  $greeting,
  @{ Title = 'Announcements Ticker'; Sec = 2; Col = 1; Order = 1; Variant = 'bold'; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  @{ Title = 'News Carousel'; Sec = 3; Col = 1; Order = 1; Variant = 'bold'; List = 'News' },
  @{ Title = 'KPI Tiles'; Sec = 4; Col = 1; Order = 1; Variant = 'minimal'; List = 'KPIs' },
  # deck col 1
  @{ Title = 'Tabs'; Sec = 5; Col = 1; Order = 1 },
  @{ Title = 'Content Rollup'; Sec = 5; Col = 1; Order = 2; List = 'News' },
  @{ Title = 'Kudos'; Sec = 5; Col = 1; Order = 3; List = 'Kudos' },
  @{ Title = 'Celebrations'; Sec = 5; Col = 1; Order = 4; List = 'Celebrations' },
  @{ Title = 'Raise a Ticket'; Sec = 5; Col = 1; Order = 5; List = 'Tickets' },
  @{ Title = 'Upcoming Holidays'; Sec = 5; Col = 1; Order = 6; List = 'Holidays' },
  # deck col 2
  @{ Title = 'Weather'; Sec = 5; Col = 2; Order = 1; Extra = @{ location = 'Houston' } },
  @{ Title = 'Event Countdown'; Sec = 5; Col = 2; Order = 2 },
  @{ Title = 'Upcoming Events'; Sec = 5; Col = 2; Order = 3; List = 'Events' },
  @{ Title = 'Quick Links'; Sec = 5; Col = 2; Order = 4; List = 'QuickLinks' },
  @{ Title = 'Image Gallery'; Sec = 5; Col = 2; Order = 5; List = 'IntranetGallery' },
  @{ Title = 'FAQ Accordion'; Sec = 5; Col = 2; Order = 6; List = 'FAQ' },
  # deck col 3
  @{ Title = 'Chart from a List'; Sec = 5; Col = 3; Order = 1; List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Employee of the Month'; Sec = 5; Col = 3; Order = 2; Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Poll'; Sec = 5; Col = 3; Order = 3; List = 'PollVotes' },
  @{ Title = 'Org Chart'; Sec = 5; Col = 3; Order = 4; LiveNoList = $true },
  @{ Title = 'People Directory'; Sec = 5; Col = 3; Order = 5; LiveNoList = $true }
)

# ---- SQUAD: team landing. Greeting/Alerts/News, then a 2/3 board + 1/3 roster, then a small deck. ----
$squadSections = @('OneColumn', 'OneColumn', 'OneColumn', 'TwoColumnLeft', 'ThreeColumn')
$squadPlan = @(
  $greeting,
  @{ Title = 'Announcements Ticker'; Sec = 2; Col = 1; Order = 1; Variant = 'bold'; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  @{ Title = 'News Carousel'; Sec = 3; Col = 1; Order = 1; Variant = 'bold'; List = 'News' },
  # main board (2/3)
  @{ Title = 'Content Rollup'; Sec = 4; Col = 1; Order = 1; List = 'News' },
  @{ Title = 'Chart from a List'; Sec = 4; Col = 1; Order = 2; List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Tabs'; Sec = 4; Col = 1; Order = 3 },
  @{ Title = 'Org Chart'; Sec = 4; Col = 1; Order = 4; LiveNoList = $true },
  @{ Title = 'Image Gallery'; Sec = 4; Col = 1; Order = 5; List = 'IntranetGallery' },
  @{ Title = 'FAQ Accordion'; Sec = 4; Col = 1; Order = 6; List = 'FAQ' },
  # roster (1/3)
  @{ Title = 'KPI Tiles'; Sec = 4; Col = 2; Order = 1; Variant = 'minimal'; List = 'KPIs' },
  @{ Title = 'Employee of the Month'; Sec = 4; Col = 2; Order = 2; Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Upcoming Events'; Sec = 4; Col = 2; Order = 3; List = 'Events' },
  @{ Title = 'Kudos'; Sec = 4; Col = 2; Order = 4; List = 'Kudos' },
  @{ Title = 'Quick Links'; Sec = 4; Col = 2; Order = 5; List = 'QuickLinks' },
  @{ Title = 'People Directory'; Sec = 4; Col = 2; Order = 6; LiveNoList = $true },
  # small footer deck (3 columns x 2)
  @{ Title = 'Weather'; Sec = 5; Col = 1; Order = 1; Extra = @{ location = 'Singapore' } },
  @{ Title = 'Raise a Ticket'; Sec = 5; Col = 1; Order = 2; List = 'Tickets' },
  @{ Title = 'Event Countdown'; Sec = 5; Col = 2; Order = 1 },
  @{ Title = 'Poll'; Sec = 5; Col = 2; Order = 2; List = 'PollVotes' },
  @{ Title = 'Upcoming Holidays'; Sec = 5; Col = 3; Order = 1; List = 'Holidays' },
  @{ Title = 'Celebrations'; Sec = 5; Col = 3; Order = 2; List = 'Celebrations' }
)

# ---- MOMENTUM: bold magazine flow. News+KPI band, a 3-col feature strip, then main + side. ----
$momentumSections = @('OneColumn', 'OneColumn', 'TwoColumnLeft', 'ThreeColumn', 'TwoColumnLeft')
$momentumPlan = @(
  $greeting,
  @{ Title = 'Announcements Ticker'; Sec = 2; Col = 1; Order = 1; Variant = 'bold'; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  # news hero + KPI
  @{ Title = 'News Carousel'; Sec = 3; Col = 1; Order = 1; Variant = 'bold'; List = 'News'; Container = @{ showBorder = $false; backgroundMode = 'transparent' } },
  @{ Title = 'KPI Tiles'; Sec = 3; Col = 2; Order = 1; Variant = 'minimal'; List = 'KPIs' },
  # 3-col feature strip
  @{ Title = 'Chart from a List'; Sec = 4; Col = 1; Order = 1; List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Kudos'; Sec = 4; Col = 1; Order = 2; List = 'Kudos' },
  @{ Title = 'Employee of the Month'; Sec = 4; Col = 2; Order = 1; Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Upcoming Events'; Sec = 4; Col = 2; Order = 2; List = 'Events' },
  @{ Title = 'Poll'; Sec = 4; Col = 3; Order = 1; List = 'PollVotes' },
  @{ Title = 'Quick Links'; Sec = 4; Col = 3; Order = 2; List = 'QuickLinks' },
  # main (2/3)
  @{ Title = 'Tabs'; Sec = 5; Col = 1; Order = 1 },
  @{ Title = 'Content Rollup'; Sec = 5; Col = 1; Order = 2; List = 'News' },
  @{ Title = 'Image Gallery'; Sec = 5; Col = 1; Order = 3; List = 'IntranetGallery' },
  @{ Title = 'Org Chart'; Sec = 5; Col = 1; Order = 4; LiveNoList = $true },
  @{ Title = 'FAQ Accordion'; Sec = 5; Col = 1; Order = 5; List = 'FAQ' },
  # side (1/3)
  @{ Title = 'Weather'; Sec = 5; Col = 2; Order = 1; Extra = @{ location = 'Sydney' } },
  @{ Title = 'Event Countdown'; Sec = 5; Col = 2; Order = 2 },
  @{ Title = 'Upcoming Holidays'; Sec = 5; Col = 2; Order = 3; List = 'Holidays' },
  @{ Title = 'Celebrations'; Sec = 5; Col = 2; Order = 4; List = 'Celebrations' },
  @{ Title = 'People Directory'; Sec = 5; Col = 2; Order = 5; LiveNoList = $true },
  @{ Title = 'Raise a Ticket'; Sec = 5; Col = 2; Order = 6; List = 'Tickets' }
)

# ---- SPOTLIGHT: showcase. Greeting/Alerts, news hero + rollup, KPI, a 3-col grid, then main + side. ----
$spotlightSections = @('OneColumn', 'OneColumn', 'TwoColumnLeft', 'OneColumn', 'ThreeColumn', 'TwoColumnLeft')
$spotlightPlan = @(
  $greeting,
  @{ Title = 'Announcements Ticker'; Sec = 2; Col = 1; Order = 1; Variant = 'bold'; List = 'Announcements'; Extra = @{ severityField = 'Severity'; linkField = 'Link'; messageField = 'Title' } },
  # news hero + latest
  @{ Title = 'News Carousel'; Sec = 3; Col = 1; Order = 1; Variant = 'bold'; List = 'News'; Container = @{ showBorder = $false; backgroundMode = 'transparent' } },
  @{ Title = 'Content Rollup'; Sec = 3; Col = 2; Order = 1; List = 'News' },
  # pulse
  @{ Title = 'KPI Tiles'; Sec = 4; Col = 1; Order = 1; Variant = 'minimal'; List = 'KPIs' },
  # 3-col grid
  @{ Title = 'Quick Links'; Sec = 5; Col = 1; Order = 1; List = 'QuickLinks' },
  @{ Title = 'Chart from a List'; Sec = 5; Col = 1; Order = 2; List = 'Tickets'; Extra = @{ categoryField = 'TicketStatus' } },
  @{ Title = 'Weather'; Sec = 5; Col = 1; Order = 3; Extra = @{ location = 'Bangalore' } },
  @{ Title = 'Employee of the Month'; Sec = 5; Col = 2; Order = 1; Variant = 'bold'; List = 'EmployeeOfMonth' },
  @{ Title = 'Poll'; Sec = 5; Col = 2; Order = 2; List = 'PollVotes' },
  @{ Title = 'Event Countdown'; Sec = 5; Col = 2; Order = 3 },
  @{ Title = 'Upcoming Events'; Sec = 5; Col = 3; Order = 1; List = 'Events' },
  @{ Title = 'Kudos'; Sec = 5; Col = 3; Order = 2; List = 'Kudos' },
  @{ Title = 'Celebrations'; Sec = 5; Col = 3; Order = 3; List = 'Celebrations' },
  # main (2/3) + side (1/3)
  @{ Title = 'Tabs'; Sec = 6; Col = 1; Order = 1 },
  @{ Title = 'Image Gallery'; Sec = 6; Col = 1; Order = 2; List = 'IntranetGallery' },
  @{ Title = 'Org Chart'; Sec = 6; Col = 1; Order = 3; LiveNoList = $true },
  @{ Title = 'FAQ Accordion'; Sec = 6; Col = 1; Order = 4; List = 'FAQ' },
  @{ Title = 'Raise a Ticket'; Sec = 6; Col = 2; Order = 1; List = 'Tickets' },
  @{ Title = 'People Directory'; Sec = 6; Col = 2; Order = 2; LiveNoList = $true },
  @{ Title = 'Upcoming Holidays'; Sec = 6; Col = 2; Order = 3; List = 'Holidays' }
)

# pick the composition preset (default = cascade)
$sections = $cascadeSections
$plan = $cascadePlan
if ($Layout -eq 'spotlight') { $sections = $spotlightSections; $plan = $spotlightPlan }
elseif ($Layout -eq 'cockpit') { $sections = $cockpitSections; $plan = $cockpitPlan }
elseif ($Layout -eq 'momentum') { $sections = $momentumSections; $plan = $momentumPlan }
elseif ($Layout -eq 'squad') { $sections = $squadSections; $plan = $squadPlan }
Write-Host "  composition preset: $Layout ($($sections.Count) sections, $($plan.Count) web parts)" -ForegroundColor DarkGray

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
  # apply this site's tile/element colour palette to the colour-aware web parts
  if ($TileColors -ne 'none' -and $tileColorWebParts -contains $wp.Title) { $props.tileColorMode = $TileColors }
  if ($RealData) {
    if ($wp.List) { $props.useDemoData = $false; $props.listTitle = $wp.List }
    elseif ($wp.LiveNoList) { $props.useDemoData = $false }
    if ($wp.Extra) { foreach ($e in $wp.Extra.GetEnumerator()) { $props[$e.Key] = $e.Value } }
  }
  $json = $props | ConvertTo-Json -Compress
  $wpArgs = @{ Page = "Home.aspx"; Component = $wp.Title; Section = $wp.Sec; Column = $wp.Col; Order = $wp.Order; ErrorAction = 'Stop' }
  if ($propParam) { $wpArgs[$propParam] = $json }
  try {
    Add-PnPPageWebPart @wpArgs | Out-Null
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
