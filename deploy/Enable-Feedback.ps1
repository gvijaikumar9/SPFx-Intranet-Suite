<#
.SYNOPSIS
  Registers (or updates) the site-wide Feedback widget Application Customizer on a site.
  It shows a fixed "Feedback" bubble in the bottom-right of every page; submissions go to
  the Feedback list (run Provision-IntranetLists.ps1 first).

  Re-run with the same values to update, or with -Remove to unregister.

.EXAMPLE
  .\deploy\Enable-Feedback.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" -Accent "#0f6cbd"

.EXAMPLE
  .\deploy\Enable-Feedback.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<app-id>" -Remove
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$ListTitle = 'Feedback',
  [string]$Accent    = '#0f6cbd',
  [switch]$Remove
)
$ErrorActionPreference = "Stop"

# Must match the id in FeedbackApplicationCustomizer.manifest.json
$componentId = "e5a3c7d9-4f8b-4e2a-b6d1-9c3e5a7f1b40"
$actionName  = "IntranetFeedback"

Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId

# clear any existing registration first (idempotent)
Get-PnPCustomAction -Scope Web | Where-Object { $_.ClientSideComponentId -eq $componentId -or $_.Name -eq $actionName } |
  ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Scope Web -Force }

if ($Remove) {
  Write-Host "Removed the Feedback widget from $Url" -ForegroundColor Cyan
  return
}

$props = @{ listTitle = $ListTitle; accent = $Accent } | ConvertTo-Json -Compress

Add-PnPCustomAction -Name $actionName -Title "Intranet Feedback" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId $componentId `
  -ClientSideComponentProperties $props `
  -Scope Web | Out-Null

Write-Host "Feedback widget enabled on $Url (accent $Accent). Reload any page to see the bubble." -ForegroundColor Cyan
