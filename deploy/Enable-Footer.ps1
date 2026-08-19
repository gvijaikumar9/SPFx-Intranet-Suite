<#
.SYNOPSIS
  Registers (or updates) the site-wide Intranet Footer Application Customizer on a site.
  The footer renders in the Bottom placeholder on every page. Link columns come from the
  FooterLinks list (run Provision-IntranetLists.ps1 first); the brand line, copyright,
  accent and social icons are passed here as component properties.

  Re-run with the same values to update, or with -Remove to unregister.

.EXAMPLE
  .\deploy\Enable-Footer.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" `
     -BrandText "Contoso" -Copyright "(c) 2026 Contoso. All rights reserved." -Accent "#0f6cbd" `
     -Social "linkedin|https://www.linkedin.com/company/contoso, email|mailto:hello@contoso.com"

.EXAMPLE
  .\deploy\Enable-Footer.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" -Remove
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$BrandText = 'Contoso',
  [string]$Blurb     = 'One place for the news, tools and people that keep us moving.',
  [string]$Copyright = '',
  [string]$ListTitle = 'FooterLinks',
  [string]$Social    = '',
  [string]$Accent    = '#0f6cbd',
  [switch]$Remove
)
$ErrorActionPreference = "Stop"

# Must match the id in IntranetFooterApplicationCustomizer.manifest.json
$componentId = "c3f1a9d7-2e6b-4a58-9d1c-7b3e5f0a2c64"
$actionName  = "IntranetFooter"

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId

# clear any existing registration first (idempotent)
Get-PnPCustomAction -Scope Web | Where-Object { $_.ClientSideComponentId -eq $componentId -or $_.Name -eq $actionName } |
  ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Scope Web -Force }

if ($Remove) {
  Write-Host "Removed the Intranet Footer from $Url" -ForegroundColor Cyan
  return
}

if (-not $Copyright) { $Copyright = ("(c) {0}. All rights reserved." -f $BrandText) }

$props = @{
  brandText = $BrandText
  blurb     = $Blurb
  copyright = $Copyright
  listTitle = $ListTitle
  social    = $Social
  accent    = $Accent
} | ConvertTo-Json -Compress

Add-PnPCustomAction -Name $actionName -Title "Intranet Footer" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId $componentId `
  -ClientSideComponentProperties $props `
  -Scope Web | Out-Null

Write-Host "Intranet Footer enabled on $Url (accent $Accent). Reload any page to see it." -ForegroundColor Cyan
