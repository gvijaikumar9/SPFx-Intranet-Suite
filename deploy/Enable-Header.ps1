<#
.SYNOPSIS
  Registers (or updates) the site-wide Intranet Header Application Customizer on a site.
  The header is a full-width band shown at the top of every page: brand (logo or text), an
  inline nav from the HeaderLinks list, and an optional call-to-action button.

  Run Provision-IntranetLists.ps1 first so the HeaderLinks list exists. Brand text, logo URL,
  home URL, CTA and accent are passed here. Re-run with the same values to update, or with
  -Remove to unregister.

.EXAMPLE
  .\deploy\Enable-Header.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" `
     -BrandText "Contoso" -LogoUrl "/sites/intranet/SiteAssets/logo.png" -Accent "#0f6cbd" `
     -CtaText "Raise a ticket" -CtaUrl "/sites/intranet/support"

.EXAMPLE
  .\deploy\Enable-Header.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" -Remove
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$BrandText = 'Contoso',
  [string]$LogoUrl   = '',   # image URL (site-relative or absolute); empty = text brand
  [string]$HomeUrl   = '',   # logo/brand link target; empty = the site itself
  [string]$ListTitle = 'HeaderLinks',
  [string]$CtaText   = '',   # optional right-hand button; both CtaText and CtaUrl needed to show
  [string]$CtaUrl    = '',
  [string]$Accent    = '#0f6cbd',
  [switch]$KeepNativeHeader,  # by default the header hides SharePoint's built-in site header (title row + nav); pass this to keep it
  [switch]$Remove
)
$ErrorActionPreference = "Stop"

# Must match the id in IntranetHeaderApplicationCustomizer.manifest.json
$componentId = "d7b2f4a6-1c8e-4f39-b05a-6e2c9a1d3f58"
$actionName  = "IntranetHeader"

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId

# clear any existing registration first (idempotent)
Get-PnPCustomAction -Scope Web | Where-Object { $_.ClientSideComponentId -eq $componentId -or $_.Name -eq $actionName } |
  ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Scope Web -Force }

if ($Remove) {
  Write-Host "Removed the Intranet Header from $Url" -ForegroundColor Cyan
  return
}

$props = @{
  brandText     = $BrandText
  logoUrl       = $LogoUrl
  homeUrl       = $HomeUrl
  listTitle     = $ListTitle
  ctaText       = $CtaText
  ctaUrl           = $CtaUrl
  accent           = $Accent
  hideNativeHeader = (-not $KeepNativeHeader)
} | ConvertTo-Json -Compress

Add-PnPCustomAction -Name $actionName -Title "Intranet Header" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId $componentId `
  -ClientSideComponentProperties $props `
  -Scope Web | Out-Null

Write-Host "Intranet Header enabled on $Url (accent $Accent). Reload any page to see the band at the top." -ForegroundColor Cyan
