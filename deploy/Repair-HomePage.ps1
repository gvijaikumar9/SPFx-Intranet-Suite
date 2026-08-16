<#
.SYNOPSIS
  Repairs a template site's Home.aspx when a duplicate/corrupt web part broke it
  ("Something went wrong"). Removes all canvas controls, then re-adds a single
  clean Announcements Ticker in a one-column section, and publishes.

.EXAMPLE
  .\deploy\Repair-HomePage.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/intranet" -ClientId "<your-entra-app-id>"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$SiteUrl,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$WebPartName = "Announcements Ticker"
)
$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
Write-Host "Resetting Home.aspx on $SiteUrl ..." -ForegroundColor Cyan

# Surgical control removal cannot fix a corrupted CanvasContent1 (PnP lists phantom
# controls it then refuses to remove). Delete the whole page and recreate it clean.
try {
  Remove-PnPPage -Identity "Home.aspx" -Force -ErrorAction Stop
  Write-Host "  deleted corrupted Home.aspx"
} catch {
  Write-Host ("  (no page to delete: {0})" -f $_.Exception.Message) -ForegroundColor DarkGray
}

$null = Add-PnPPage -Name "Home.aspx" -LayoutType Article
Add-PnPPageSection -Page "Home.aspx" -SectionTemplate OneColumn -Order 1
Add-PnPPageWebPart -Page "Home.aspx" -Component $WebPartName -Section 1 -Column 1 | Out-Null
Set-PnPPage -Identity "Home.aspx" -Publish | Out-Null

# make sure the site still opens on this page
try { Set-PnPHomePage -RootFolderRelativeUrl "SitePages/Home.aspx" -ErrorAction Stop } catch {}

Write-Host "Done. Reload the page - one clean ticker should render." -ForegroundColor Green
