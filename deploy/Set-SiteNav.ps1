<#
.SYNOPSIS
  Sets a template site's horizontal (top) navigation to match its prototype:
  Home / News / Departments / People / Tools / Support. Existing top-nav nodes are
  cleared first so the menu matches exactly.

  Placeholder sections (News, Departments, People, Tools, Support) point at the home
  page by default; repoint them to real pages later with -Targets or in the UI.

.EXAMPLE
  .\deploy\Set-SiteNav.ps1 -Url "https://contoso.sharepoint.com/sites/intranet-spotlight" -ClientId "<app-id>"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  # label -> relative/absolute URL. Defaults below point placeholders at the home page.
  [hashtable]$Targets
)
$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId

$web = Get-PnPWeb
$homeUrl = ($web.ServerRelativeUrl.TrimEnd('/')) + "/SitePages/Home.aspx"

# prototype menu, in order
$nav = @(
  @{ Label = 'Home';        Url = $homeUrl },
  @{ Label = 'News';        Url = $homeUrl },
  @{ Label = 'Departments'; Url = $homeUrl },
  @{ Label = 'People';      Url = $homeUrl },
  @{ Label = 'Tools';       Url = $homeUrl },
  @{ Label = 'Support';     Url = $homeUrl }
)
if ($Targets) {
  foreach ($n in $nav) { if ($Targets.ContainsKey($n.Label)) { $n.Url = $Targets[$n.Label] } }
}

Write-Host "Setting top navigation on $Url" -ForegroundColor Cyan

# clear existing top-nav nodes
Get-PnPNavigationNode -Location TopNavigationBar | ForEach-Object {
  Remove-PnPNavigationNode -Identity $_.Id -Force -ErrorAction SilentlyContinue
}

# add the prototype menu
foreach ($n in $nav) {
  Add-PnPNavigationNode -Location TopNavigationBar -Title $n.Label -Url $n.Url | Out-Null
  Write-Host ("  + {0}" -f $n.Label) -ForegroundColor Green
}

Write-Host "Done. Reload the site to see the new top navigation." -ForegroundColor Cyan
