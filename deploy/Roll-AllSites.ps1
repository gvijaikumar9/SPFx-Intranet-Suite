<#
.SYNOPSIS
  End-to-end rollout for all five template sites. For each site it: deploys the suite
  package, cleans up the extracted layout tuner, applies the template's accent theme,
  provisions the lists, builds the home page with that template's base variant (the
  per-web-part variant overrides are baked into Build-HomePages), and sets the news
  hero images.

  Per-web-part variants come from reconciling each prototype: Ticker/News=bold,
  KPI=minimal, Employee=bold (universal); everything else follows the -Variant base
  below. Accents are each prototype's --accent.

.EXAMPLE
  .\deploy\Roll-AllSites.ps1 -Tenant "contoso" -ClientId "<app-id>"

  # skip theming (layout + data only):
  .\deploy\Roll-AllSites.ps1 -Tenant "contoso" -ClientId "<app-id>" -SkipTheme
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Tenant,     # e.g. "contoso"
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$ImageFolder = "C:\Users\gvija\source\repos\spfx-intranet-webparts\images",
  [switch]$SkipTheme,
  [switch]$SkipImages,
  [string]$Only   # optional: roll out just one site key (e.g. "spotlight")
)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$pkg = Join-Path $root "..\sharepoint\solution\spfx-intranet-suite.sppkg"
$tunerId = "a2c4e6f8-1b3d-4f5a-9c7e-2d4f6a8b0c11"

# base variant + accent per template (from the prototype reconciliation)
# Each site: base variant, accent, theme name, composition preset, and its own tile-colour
# palette so every template looks distinct (soft / fluorescent / water / dark).
$sites = @(
  @{ Key = 'cascade';   Variant = 'card';    Accent = '#0f6cbd'; Theme = 'IntranetCascade'; Layout = 'cascade'; TileColors = 'palette' },
  @{ Key = 'cockpit';   Variant = 'compact'; Accent = '#0f6cbd'; Theme = 'IntranetCockpit'; Layout = 'cockpit'; TileColors = 'fluro' },
  @{ Key = 'squad';     Variant = 'card';    Accent = '#0f766e'; Theme = 'IntranetSquad'; Layout = 'squad'; TileColors = 'water'; Menu = @('Home', 'Backlog', 'Docs', 'Runbooks', 'People') },
  @{ Key = 'momentum';  Variant = 'bold';    Accent = '#4f46e5'; Theme = 'IntranetMomentum'; Layout = 'momentum'; TileColors = 'dark' },
  @{ Key = 'spotlight'; Variant = 'card';    Accent = '#1f9d86'; Theme = 'IntranetSpotlight'; Layout = 'spotlight'; TileColors = 'palette' }
)
if ($Only) { $sites = $sites | Where-Object { $_.Key -eq $Only }; if (-not $sites) { throw "No site with key '$Only'." } }

foreach ($s in $sites) {
  $url = "https://$Tenant.sharepoint.com/sites/intranet-$($s.Key)"
  Write-Host "`n=== $url  (variant $($s.Variant), accent $($s.Accent)) ===" -ForegroundColor Cyan
  Connect-PnPOnline -Url $url -Interactive -ClientId $ClientId

  # 1. deploy the package
  $app = Add-PnPApp -Path $pkg -Scope Site -Publish -Overwrite
  Update-PnPApp -Identity $app.Id -Scope Site

  # 2. clean up any leftover layout tuner (custom action + orphaned config list)
  Get-PnPCustomAction -Scope Web | Where-Object { $_.ClientSideComponentId -eq $tunerId } | ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Scope Web -Force }
  $l = Get-PnPList -Identity "SuiteLayoutConfig" -ErrorAction SilentlyContinue
  if ($l) { Remove-PnPList -Identity "SuiteLayoutConfig" -Force }

  # 3. apply the template accent theme
  if (-not $SkipTheme) {
    & (Join-Path $root "Set-SiteTheme.ps1") -Url $url -ClientId $ClientId -Accent $s.Accent -ThemeName $s.Theme
  }

  # 4. provision the lists (adds Celebrations etc.; idempotent)
  & (Join-Path $root "Provision-IntranetLists.ps1") -Url $url -ClientId $ClientId

  # 5. build the home page with this template's base variant + composition preset + palette
  $layout = if ($s.Layout) { $s.Layout } else { 'default' }
  $tiles = if ($s.TileColors) { $s.TileColors } else { 'none' }
  & (Join-Path $root "Build-HomePages.ps1") -Url $url -ClientId $ClientId -Variant $s.Variant -Card -RealData -Layout $layout -TileColors $tiles

  # 6. set the news hero images
  if (-not $SkipImages -and (Test-Path $ImageFolder)) {
    & (Join-Path $root "Set-NewsImages.ps1") -Url $url -ClientId $ClientId -ImageFolder $ImageFolder
  }

  # 7. set the top navigation (per-site menu if the template defines one)
  if ($s.Menu) {
    & (Join-Path $root "Set-SiteNav.ps1") -Url $url -ClientId $ClientId -Menu $s.Menu
  } else {
    & (Join-Path $root "Set-SiteNav.ps1") -Url $url -ClientId $ClientId
  }

  # 8. enable the site-wide header band with this template's accent (nav from HeaderLinks)
  & (Join-Path $root "Enable-Header.ps1") -Url $url -ClientId $ClientId -Accent $s.Accent

  # 9. enable the site-wide footer with this template's accent
  & (Join-Path $root "Enable-Footer.ps1") -Url $url -ClientId $ClientId -Accent $s.Accent

  # 10. enable the site-wide feedback bubble with this template's accent
  & (Join-Path $root "Enable-Feedback.ps1") -Url $url -ClientId $ClientId -Accent $s.Accent
}

Write-Host "`nDone. All five template sites deployed, themed and built to match their prototypes." -ForegroundColor Cyan
