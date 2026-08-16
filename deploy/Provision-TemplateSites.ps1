<#
.SYNOPSIS
  Provisions the SPFx Intranet Suite template sites on SharePoint Online.

  For each template (Cascade, Momentum, Cockpit, Squad, Spotlight) it:
    1. creates a communication site (if it does not already exist),
    2. enables a site collection app catalog on that site,
    3. deploys + installs spfx-intranet-suite.sppkg into that site's catalog,
    4. drops the Announcements Ticker on the home page as a smoke test.

  The script is idempotent - safe to re-run after you rebuild the package
  (gulp bundle --ship; gulp package-solution --ship). It overwrites the app
  and skips work that is already done.

.PREREQUISITES
  - PnP.PowerShell 3.x on PowerShell 7  (you have 3.3.0 / 7.6.3)
  - Your own Entra app registration Client Id (PnP 3.x no longer ships one)
  - A TENANT app catalog must already exist (site collection catalogs require it).
    If it does not, the script stops and tells you the one-line command to create it.
  - SharePoint Administrator rights on the tenant.

.EXAMPLE
  .\deploy\Provision-TemplateSites.ps1 -ClientId "<your-entra-app-id>" -TenantName "contoso"

  # Preview only, create nothing:
  .\deploy\Provision-TemplateSites.ps1 -ClientId "<your-entra-app-id>" -TenantName "contoso" -PlanOnly
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$ClientId,
  [Parameter(Mandatory = $true)][string]$TenantName,
  [string]$SppkgPath  = "$PSScriptRoot\..\sharepoint\solution\spfx-intranet-suite.sppkg",
  [string]$SitePrefix = "intranet-",
  [switch]$PlanOnly
)

$ErrorActionPreference = "Stop"
$adminUrl = "https://$TenantName-admin.sharepoint.com"
$rootUrl  = "https://$TenantName.sharepoint.com"
$wpName   = "Announcements Ticker"   # matches the manifest preconfiguredEntries title

# template code-name -> site title (URLs become /sites/<prefix><code>)
$templates = [ordered]@{
  "cascade"   = "Cascade"
  "momentum"  = "Momentum"
  "cockpit"   = "Cockpit"
  "squad"     = "Squad"
  "spotlight" = "Spotlight"
}

function Connect-Site([string]$url) {
  Connect-PnPOnline -Url $url -Interactive -ClientId $ClientId
}

# --- resolve package ---------------------------------------------------------
if (-not (Test-Path $SppkgPath)) {
  throw "Package not found at $SppkgPath. Build it first: gulp bundle --ship; gulp package-solution --ship"
}
$sppkg = (Resolve-Path $SppkgPath).Path
Write-Host "Package : $sppkg"    -ForegroundColor Cyan
Write-Host "Tenant  : $rootUrl"  -ForegroundColor Cyan

# --- connect to admin + verify tenant app catalog ----------------------------
Write-Host "`nConnecting to $adminUrl (a browser sign-in opens on first run)..." -ForegroundColor Cyan
Connect-Site $adminUrl

$tenantCatalog = Get-PnPTenantAppCatalogUrl
if (-not $tenantCatalog) {
  throw @"
No TENANT app catalog exists yet. Site collection app catalogs require one.
Create it once, then re-run this script:

  Register-PnPAppCatalogSite -Url $rootUrl/sites/appcatalog -Owner <you@$TenantName.onmicrosoft.com> -TimeZone 4
"@
}
Write-Host "Tenant app catalog: $tenantCatalog" -ForegroundColor Green

# --- per-template loop -------------------------------------------------------
$results = @()
foreach ($code in $templates.Keys) {
  $title   = $templates[$code]
  $siteUrl = "$rootUrl/sites/$SitePrefix$code"
  Write-Host "`n=== $title  ($siteUrl) ===" -ForegroundColor Yellow

  if ($PlanOnly) {
    Write-Host "  [PlanOnly] would: create site -> app catalog -> deploy+install -> add '$wpName'"
    $results += [pscustomobject]@{ Template = $title; Url = $siteUrl; Status = "planned" }
    continue
  }

  try {
    # 1) create the communication site if missing
    Connect-Site $adminUrl
    $exists = Get-PnPTenantSite -Identity $siteUrl -ErrorAction SilentlyContinue
    if (-not $exists) {
      Write-Host "  creating communication site..."
      # NOTE: New-PnPSite -Wait can hang on a final status poll even after the
      # site is fully provisioned. Create without -Wait, then poll ourselves.
      New-PnPSite -Type CommunicationSite -Title $title -Url $siteUrl | Out-Null
      $ready = $false
      for ($i = 0; $i -lt 30; $i++) {
        $st = (Get-PnPTenantSite -Identity $siteUrl -ErrorAction SilentlyContinue).Status
        if ($st -eq 'Active') { $ready = $true; break }
        Write-Host ("    ...provisioning ({0}s)" -f (($i + 1) * 10))
        Start-Sleep -Seconds 10
      }
      if (-not $ready) { Write-Host "    still provisioning after ~5 min - continuing anyway" -ForegroundColor DarkYellow }
    } else {
      Write-Host "  site already exists - skipping create"
    }

    # 2) enable the site collection app catalog
    Write-Host "  enabling site collection app catalog..."
    try { Add-PnPSiteCollectionAppCatalog -Site $siteUrl -ErrorAction Stop }
    catch { Write-Host "    (already enabled or pending: $($_.Exception.Message))" -ForegroundColor DarkGray }

    # 3) deploy + install the package into the site catalog
    Connect-Site $siteUrl
    Write-Host "  deploying package to the site app catalog..."
    # A newly enabled site collection app catalog can take up to a minute to
    # provision. Retry Add-PnPApp until the catalog exists.
    $app = $null
    for ($r = 0; $r -lt 12; $r++) {
      try { $app = Add-PnPApp -Path $sppkg -Scope Site -Publish -Overwrite -ErrorAction Stop; break }
      catch {
        if ($_.Exception.Message -match 'sitecollectionappcatalog|ResourceNotFound|Cannot find resource') {
          Write-Host "    app catalog not ready yet, waiting 15s..." -ForegroundColor DarkGray
          Start-Sleep -Seconds 15
          Connect-Site $siteUrl
        } else { throw }
      }
    }
    if (-not $app) { throw "Site collection app catalog did not become available in time." }
    Write-Host "  installing / updating app on the site..."
    try { Install-PnPApp -Identity $app.Id -Scope Site -ErrorAction Stop }
    catch {
      try { Update-PnPApp -Identity $app.Id -Scope Site -ErrorAction Stop; Write-Host "    upgraded existing install" }
      catch { Write-Host "    app already installed" -ForegroundColor DarkGray }
    }

    # 4) add the Announcements Ticker to Home.aspx (skip if already there)
    Write-Host "  adding '$wpName' to Home.aspx..."
    try {
      $page = Get-PnPPage -Identity "Home.aspx"
      $has  = $page.Controls | Where-Object { $_.Title -eq $wpName }
      if ($has) {
        Write-Host "    already on the page - skipping" -ForegroundColor DarkGray
      } else {
        # A freshly provisioned Home page may have no canvas sections yet.
        if ($page.Sections.Count -eq 0) {
          Add-PnPPageSection -Page "Home.aspx" -SectionTemplate OneColumn -Order 1
        }
        try {
          Add-PnPPageWebPart -Page "Home.aspx" -Component $wpName -Section 1 -Column 1 -ErrorAction Stop
        } catch {
          # fall back to default placement - but only if the first attempt did
          # not already add it (prevents a duplicate web part).
          $recheck = (Get-PnPPage -Identity "Home.aspx").Controls | Where-Object { $_.Title -eq $wpName }
          if (-not $recheck) {
            Add-PnPPageWebPart -Page "Home.aspx" -Component $wpName -ErrorAction Stop
          }
        }
        Set-PnPPage -Identity "Home.aspx" -Publish | Out-Null
        Write-Host "    added + published" -ForegroundColor Green
      }
    } catch {
      Write-Host "    could not add automatically ($($_.Exception.Message))." -ForegroundColor DarkYellow
      Write-Host "    Add it by hand from the page's web part picker to verify - the app IS installed."
    }

    Write-Host "  DONE -> $siteUrl" -ForegroundColor Green
    $results += [pscustomobject]@{ Template = $title; Url = $siteUrl; Status = "ok" }
  }
  catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
    $results += [pscustomobject]@{ Template = $title; Url = $siteUrl; Status = "failed: $($_.Exception.Message)" }
  }
}

Write-Host "`n================ SUMMARY ================" -ForegroundColor Cyan
$results | Format-Table -AutoSize
Write-Host "Open each site above and confirm the Announcements Ticker is rotating (sample data)." -ForegroundColor Cyan
