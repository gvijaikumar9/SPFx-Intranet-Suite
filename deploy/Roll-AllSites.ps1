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
  [switch]$SkipImages
)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$pkg = Join-Path $root "..\sharepoint\solution\spfx-intranet-suite.sppkg"
$tunerId = "a2c4e6f8-1b3d-4f5a-9c7e-2d4f6a8b0c11"

# base variant + accent per template (from the prototype reconciliation)
$sites = @(
  @{ Key = 'cascade';   Variant = 'card';    Accent = '#0f6cbd'; Theme = 'IntranetCascade' },
  @{ Key = 'cockpit';   Variant = 'compact'; Accent = '#0f6cbd'; Theme = 'IntranetCockpit' },
  @{ Key = 'squad';     Variant = 'card';    Accent = '#0f766e'; Theme = 'IntranetSquad' },
  @{ Key = 'momentum';  Variant = 'card';    Accent = '#4f46e5'; Theme = 'IntranetMomentum' },
  @{ Key = 'spotlight'; Variant = 'card';    Accent = '#1f9d86'; Theme = 'IntranetSpotlight' }
)

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

  # 5. build the home page with this template's base variant
  & (Join-Path $root "Build-HomePages.ps1") -Url $url -ClientId $ClientId -Variant $s.Variant -Card -RealData

  # 6. set the news hero images
  if (-not $SkipImages -and (Test-Path $ImageFolder)) {
    & (Join-Path $root "Set-NewsImages.ps1") -Url $url -ClientId $ClientId -ImageFolder $ImageFolder
  }
}

Write-Host "`nDone. All five template sites deployed, themed and built to match their prototypes." -ForegroundColor Cyan
