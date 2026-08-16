<#
.SYNOPSIS
  Generates a full SharePoint/Fluent theme palette from a single primary accent colour
  and applies it to a site, so the suite web parts (which follow the theme accent) take
  on that colour. This is what makes each template site look distinct.

  Creating the tenant theme needs SharePoint-admin rights; applying it to the site needs
  site-owner rights. If the tenant-theme step is blocked, the script warns and continues.

.EXAMPLE
  .\deploy\Set-SiteTheme.ps1 -Url "https://contoso.sharepoint.com/sites/intranet-momentum" `
    -ClientId "<app-id>" -Accent "#4f46e5" -ThemeName "IntranetMomentum"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [Parameter(Mandatory = $true)][string]$Accent,      # e.g. "#4f46e5"
  [Parameter(Mandatory = $true)][string]$ThemeName
)
$ErrorActionPreference = "Stop"

# --- mix a channel toward black (t=0) or white (t=255) by fraction p ---
function Get-Shade([int]$r, [int]$g, [int]$b, [double]$p, [int]$t) {
  $nr = [int][Math]::Round($r * (1 - $p) + $t * $p)
  $ng = [int][Math]::Round($g * (1 - $p) + $t * $p)
  $nb = [int][Math]::Round($b * (1 - $p) + $t * $p)
  return ('#{0:x2}{1:x2}{2:x2}' -f $nr, $ng, $nb)
}

# --- build a Fluent theme palette from one primary colour ---
function New-ThemePalette([string]$hex) {
  $h = $hex.TrimStart('#')
  $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
  return @{
    themePrimary         = "#$h"
    themeLighterAlt      = (Get-Shade $r $g $b 0.95 255)
    themeLighter         = (Get-Shade $r $g $b 0.81 255)
    themeLight           = (Get-Shade $r $g $b 0.63 255)
    themeTertiary        = (Get-Shade $r $g $b 0.35 255)
    themeSecondary       = (Get-Shade $r $g $b 0.10 255)
    themeDarkAlt         = (Get-Shade $r $g $b 0.08 0)
    themeDark            = (Get-Shade $r $g $b 0.21 0)
    themeDarker          = (Get-Shade $r $g $b 0.32 0)
    neutralLighterAlt    = "#faf9f8"; neutralLighter = "#f3f2f1"; neutralLight = "#edebe9"
    neutralQuaternaryAlt = "#e1dfdd"; neutralQuaternary = "#d2d0ce"; neutralTertiaryAlt = "#c8c6c4"
    neutralTertiary      = "#a19f9d"; neutralSecondary = "#605e5c"; neutralPrimaryAlt = "#3b3a39"
    neutralPrimary       = "#323130"; neutralDark = "#201f1e"; black = "#000000"; white = "#ffffff"
    bodyBackground       = "#ffffff"; bodyText = "#323130"
  }
}

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
$palette = New-ThemePalette $Accent
Write-Host "Applying theme '$ThemeName' (accent $Accent) to $Url" -ForegroundColor Cyan

# Register the tenant theme (needs SharePoint admin). Idempotent via -Overwrite.
try {
  Add-PnPTenantTheme -Identity $ThemeName -Palette $palette -IsInverted:$false -Overwrite -ErrorAction Stop
  Write-Host "  registered tenant theme $ThemeName" -ForegroundColor Green
} catch {
  Write-Host ("  ! could not register tenant theme (need SharePoint admin?): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
}

# Apply it to the site.
try {
  Set-PnPWebTheme -Theme $ThemeName -WebUrl $Url -ErrorAction Stop
  Write-Host "  applied to site. Reload to see the new accent." -ForegroundColor Green
} catch {
  Write-Host ("  ! could not apply theme to site: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
}
