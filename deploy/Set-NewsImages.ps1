<#
.SYNOPSIS
  Uploads hero images to the site's Site Assets and sets the NewsImage field on the
  News list items, so the News Carousel web part shows image slides instead of the
  gradient fallback. Images are same-origin, so there is no CSP problem.

  Matches each image to a News item by keywords in the item's Title and Category.

.EXAMPLE
  .\deploy\Set-NewsImages.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" `
    -ClientId "<your-entra-app-id>" `
    -ImageFolder "C:\Users\gvija\source\repos\spfx-intranet-webparts\images"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [Parameter(Mandatory = $true)][string]$ImageFolder,
  [string]$ListTitle = "News"
)
$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
$web = Get-PnPWeb
$siteUrl = $web.Url
$webRel = $web.ServerRelativeUrl.TrimEnd('/')

# SiteAssets often blocks uploads for a non-owner app. Use a dedicated library we
# create (the app owns it, so uploads are always allowed).
$libTitle = "NewsAssets"
if (-not (Get-PnPList -Identity $libTitle -ErrorAction SilentlyContinue)) {
  New-PnPList -Title $libTitle -Template DocumentLibrary -OnQuickLaunch:$false | Out-Null
  Write-Host "  created library $libTitle" -ForegroundColor Green
}
$assetsFolder = "$webRel/$libTitle"
Write-Host "Setting News hero images on $Url (into $assetsFolder)" -ForegroundColor Cyan

# Hero image -> keywords that should appear in the item's Title or Category.
$map = @(
  @{ File = 'featured-engineers.jpg'; Keywords = @('uptime', 'platform', 'engineer', 'company', 'reliability') },
  @{ File = 'featured-award.jpg';     Keywords = @('award', 'winner', 'people', 'recognition', 'innovation') },
  @{ File = 'featured-lounge.jpg';    Keywords = @('hq', 'floor', 'collaboration', 'office', 'workplace') }
)

# Upload each hero image to Site Assets (overwrites if it already exists).
foreach ($m in $map) {
  $local = Join-Path $ImageFolder $m.File
  if (Test-Path $local) {
    Add-PnPFile -Path $local -Folder $assetsFolder -ErrorAction Stop | Out-Null
    $m.Url = "$siteUrl/$libTitle/$($m.File)"
    Write-Host "  uploaded $($m.File)" -ForegroundColor Green
  }
  else {
    Write-Host "  missing $local (skipped)" -ForegroundColor DarkYellow
  }
}

$uploaded = @($map | Where-Object { $_.Url })
if ($uploaded.Count -eq 0) { Write-Host "No images uploaded; nothing to set." -ForegroundColor DarkYellow; return }
$fallback = $uploaded[0]

# Assign an image to each News item by keyword match, else the first uploaded image.
$items = Get-PnPListItem -List $ListTitle
$set = 0
foreach ($it in $items) {
  $text = ((("" + $it["Title"]) + " " + ("" + $it["Category"]))).ToLower()
  $chosen = $null
  foreach ($m in $uploaded) {
    foreach ($kw in $m.Keywords) { if ($text.Contains($kw)) { $chosen = $m; break } }
    if ($chosen) { break }
  }
  if (-not $chosen) { $chosen = $fallback }
  Set-PnPListItem -List $ListTitle -Identity $it.Id -Values @{ NewsImage = $chosen.Url } | Out-Null
  Write-Host ("  News #{0} '{1}' -> {2}" -f $it.Id, $it["Title"], $chosen.File) -ForegroundColor Cyan
  $set++
}

Write-Host "Done. Set images on $set news item(s). Reload the page." -ForegroundColor Cyan
