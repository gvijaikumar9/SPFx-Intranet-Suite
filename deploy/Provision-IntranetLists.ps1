<#
.SYNOPSIS
  Provisions the SharePoint lists and libraries that the Intranet Suite web parts
  read from, with columns and sample data, on one site.

  Run it against each template site (or your dev site) so the web parts have real
  data to bind to. Idempotent: existing lists/fields are left alone, and sample
  rows are only added when a list is empty.

.PREREQUISITES
  PnP.PowerShell 3.x on PowerShell 7, and your own Entra app client id.

.EXAMPLE
  .\deploy\Provision-IntranetLists.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<your-entra-app-id>"

  # No sample rows, just the schema:
  .\deploy\Provision-IntranetLists.ps1 -Url "https://contoso.sharepoint.com/sites/intranet" -ClientId "<your-entra-app-id>" -NoSampleData
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Url,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [switch]$NoSampleData
)

$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
Write-Host "Provisioning intranet lists on $Url" -ForegroundColor Cyan

# --- helper: create a list (if missing) and its fields (if missing) ---
function New-IntranetList {
    param([string]$Title, [string]$Template = 'GenericList', [object[]]$Fields)

    if (-not (Get-PnPList -Identity $Title -ErrorAction SilentlyContinue)) {
        New-PnPList -Title $Title -Template $Template -OnQuickLaunch:$false | Out-Null
        Write-Host "  created $Template '$Title'" -ForegroundColor Green
    }
    else { Write-Host "  '$Title' exists" -ForegroundColor DarkGray }

    foreach ($f in $Fields) {
        if (Get-PnPField -List $Title -Identity $f.Internal -ErrorAction SilentlyContinue) { continue }
        $p = @{ List = $Title; DisplayName = $f.Display; InternalName = $f.Internal; Type = $f.Type; AddToDefaultView = $true }
        if ($f.Choices) { $p.Choices = $f.Choices }
        Add-PnPField @p | Out-Null
        Write-Host "    + $($f.Display) ($($f.Type))"
    }
}

# --- helper: add sample rows only when the list is empty ---
function Add-Sample {
    param([string]$List, [hashtable[]]$Rows)
    if ($NoSampleData) { return }
    $count = (Get-PnPList -Identity $List).ItemCount
    if ($count -gt 0) { Write-Host "    ($List already has $count items, skipping sample data)" -ForegroundColor DarkGray; return }
    $seeded = 0
    foreach ($r in $Rows) {
        try { Add-PnPListItem -List $List -Values $r -ErrorAction Stop | Out-Null; $seeded++ }
        catch { Write-Host "    (skipped a row in $($List): $($_.Exception.Message))" -ForegroundColor DarkYellow }
    }
    Write-Host "    seeded $seeded sample row(s) into $List" -ForegroundColor Green
}

# ============================ ANNOUNCEMENTS (ticker) ============================
New-IntranetList -Title 'Announcements' -Fields @(
    @{ Display = 'Severity'; Internal = 'Severity'; Type = 'Choice'; Choices = @('Info', 'Warning', 'Critical') },
    @{ Display = 'Link';     Internal = 'Link';     Type = 'URL' }
)
Add-Sample -List 'Announcements' -Rows @(
    @{ Title = 'Payroll cut-off is this Friday at 5pm.'; Severity = 'Warning' },
    @{ Title = 'VPN maintenance Saturday 10pm to 2am.';  Severity = 'Info' },
    @{ Title = 'Security patch is mandatory by month end.'; Severity = 'Critical' }
)

# ============================ TICKETS (raise a ticket) ============================
New-IntranetList -Title 'Tickets' -Fields @(
    @{ Display = 'Category';    Internal = 'Category';    Type = 'Choice'; Choices = @('IT', 'HR', 'Facilities', 'Other') },
    @{ Display = 'Description'; Internal = 'TicketDesc';  Type = 'Note' },
    @{ Display = 'Status';      Internal = 'TicketStatus'; Type = 'Choice'; Choices = @('New', 'In progress', 'Resolved') },
    @{ Display = 'Requestor';   Internal = 'Requestor';   Type = 'User' }
)
Add-Sample -List 'Tickets' -Rows @(
    @{ Title = 'Laptop will not charge'; Category = 'IT'; TicketStatus = 'In progress' },
    @{ Title = 'New monitor request';    Category = 'Facilities'; TicketStatus = 'Resolved' }
)

# ============================ KUDOS (recognition) ============================
New-IntranetList -Title 'Kudos' -Fields @(
    @{ Display = 'To';      Internal = 'ToPerson';   Type = 'User' },
    @{ Display = 'From';    Internal = 'FromPerson'; Type = 'User' },
    @{ Display = 'Message'; Internal = 'KudosMsg';   Type = 'Note' }
)
Add-Sample -List 'Kudos' -Rows @(
    @{ Title = 'Saved the demo'; KudosMsg = 'Jumped in and fixed the build an hour before the client call.' },
    @{ Title = 'Best onboarding'; KudosMsg = 'Made the new starter feel welcome from day one.' }
)

# ============================ EMPLOYEE OF THE MONTH ============================
New-IntranetList -Title 'EmployeeOfMonth' -Fields @(
    @{ Display = 'Employee'; Internal = 'Employee'; Type = 'User' },
    @{ Display = 'Month';    Internal = 'EomMonth'; Type = 'Text' },
    @{ Display = 'Citation'; Internal = 'Citation'; Type = 'Note' }
)
# Employee is a Person field. Seeding person fields from PnP is unreliable
# ("specified user could not be found"), so seed the text fields only and let an admin
# set the Employee with the people picker (the web part reads it for the name and photo).
Add-Sample -List 'EmployeeOfMonth' -Rows @(
    @{ Title = "This month's recognition"; EomMonth = 'This month'; Citation = 'For steady delivery and mentoring across two teams.' }
)

# ============================ QUICK LINKS ============================
New-IntranetList -Title 'QuickLinks' -Fields @(
    @{ Display = 'Url';  Internal = 'LinkUrl'; Type = 'URL' },
    @{ Display = 'Icon'; Internal = 'IconName'; Type = 'Text' }
)
Add-Sample -List 'QuickLinks' -Rows @(
    @{ Title = 'Book leave';   LinkUrl = 'https://contoso.sharepoint.com/sites/hr, Book leave'; IconName = 'Vacation' },
    @{ Title = 'Raise a PO';   LinkUrl = 'https://contoso.sharepoint.com/sites/finance, Raise a PO'; IconName = 'Money' },
    @{ Title = 'IT support';   LinkUrl = 'https://contoso.sharepoint.com/sites/it, IT support'; IconName = 'Headset' },
    @{ Title = 'Directory';    LinkUrl = 'https://contoso.sharepoint.com/sites/people, Directory'; IconName = 'People' }
)

# ============================ KPI TILES ============================
New-IntranetList -Title 'KPIs' -Fields @(
    @{ Display = 'Value'; Internal = 'KpiValue'; Type = 'Text' },
    @{ Display = 'Trend'; Internal = 'KpiTrend'; Type = 'Choice'; Choices = @('Up', 'Down', 'Flat') },
    @{ Display = 'Delta'; Internal = 'KpiDelta'; Type = 'Text' }
)
Add-Sample -List 'KPIs' -Rows @(
    @{ Title = 'Open IT tickets'; KpiValue = '42';     KpiTrend = 'Down'; KpiDelta = '12% vs last week' },
    @{ Title = 'Employees';       KpiValue = '1,204';  KpiTrend = 'Up';   KpiDelta = '18 new hires' },
    @{ Title = 'Platform uptime'; KpiValue = '99.98%'; KpiTrend = 'Up';   KpiDelta = 'above target' },
    @{ Title = 'Engagement';      KpiValue = '72';     KpiTrend = 'Up';   KpiDelta = '+4 pts' }
)

# ============================ FAQ (accordion) ============================
New-IntranetList -Title 'FAQ' -Fields @(
    @{ Display = 'Answer';   Internal = 'FaqAnswer';   Type = 'Note' },
    @{ Display = 'Category'; Internal = 'FaqCategory'; Type = 'Text' }
)
Add-Sample -List 'FAQ' -Rows @(
    @{ Title = 'How do I reset my password?'; FaqAnswer = 'Use the self-service portal at aka.ms/sspr, or call the service desk.'; FaqCategory = 'IT' },
    @{ Title = 'When is payday?'; FaqAnswer = 'The last working day of each month.'; FaqCategory = 'HR' },
    @{ Title = 'How do I book a meeting room?'; FaqAnswer = 'Use the room finder in Outlook when creating the meeting.'; FaqCategory = 'Facilities' }
)

# ============================ EVENTS (team calendar) ============================
New-IntranetList -Title 'Events' -Template 'Events'
Add-Sample -List 'Events' -Rows @(
    @{ Title = 'All-hands town hall'; EventDate = (Get-Date).AddDays(3);  EndDate = (Get-Date).AddDays(3).AddHours(1);  Location = 'Auditorium' },
    @{ Title = 'Lunch and learn';     EventDate = (Get-Date).AddDays(7);  EndDate = (Get-Date).AddDays(7).AddHours(1);  Location = 'Cafe' }
)

# ============================ GALLERY (photo library) ============================
if (-not (Get-PnPList -Identity 'IntranetGallery' -ErrorAction SilentlyContinue)) {
    New-PnPList -Title 'IntranetGallery' -Template 'DocumentLibrary' -OnQuickLaunch:$false | Out-Null
    Write-Host "  created library 'IntranetGallery' (upload 'Life at Contoso' photos here)" -ForegroundColor Green
}
else { Write-Host "  'IntranetGallery' exists" -ForegroundColor DarkGray }

# ============================ NEWS (carousel) ============================
New-IntranetList -Title 'News' -Fields @(
    @{ Display = 'Summary';  Internal = 'Summary';   Type = 'Note' },
    @{ Display = 'Category'; Internal = 'Category';  Type = 'Text' },
    @{ Display = 'Image';    Internal = 'NewsImage'; Type = 'URL' },
    @{ Display = 'Link';     Internal = 'NewsLink';  Type = 'URL' }
)
Add-Sample -List 'News' -Rows @(
    @{ Title = 'Platform uptime hits 99.98% this quarter'; Summary = 'Engineering and operations share how the reliability program paid off.'; Category = 'Company news' },
    @{ Title = 'Meet the 2026 Innovation Award winners'; Summary = 'Five teams recognised for standout customer impact this year.'; Category = 'People' },
    @{ Title = 'The new HQ collaboration floors open Monday'; Summary = 'Book a tour and see what changed on levels 4 through 6.'; Category = 'Workplace' }
)

# ============================ HOLIDAYS ============================
New-IntranetList -Title 'Holidays' -Fields @(
    @{ Display = 'Date';   Internal = 'HolidayDate'; Type = 'DateTime' },
    @{ Display = 'Region'; Internal = 'Region';      Type = 'Text' }
)
Add-Sample -List 'Holidays' -Rows @(
    @{ Title = 'Thanksgiving';  HolidayDate = (Get-Date).AddDays(20); Region = 'US' },
    @{ Title = 'Diwali';        HolidayDate = (Get-Date).AddDays(35); Region = 'IN' },
    @{ Title = 'Christmas Day'; HolidayDate = (Get-Date).AddDays(60); Region = 'Global' }
)

# ============================ POLL VOTES (poll web part) ============================
# Each vote is one item whose Title is the chosen option. Title column is enough.
New-IntranetList -Title 'PollVotes' -Fields @()

# ============================ CELEBRATIONS (birthdays / anniversaries) ============================
New-IntranetList -Title 'Celebrations' -Fields @(
    @{ Display = 'Date';   Internal = 'CelebrationDate'; Type = 'DateTime' },
    @{ Display = 'Type';   Internal = 'CelebrationType'; Type = 'Choice'; Choices = @('Birthday', 'Work anniversary') },
    @{ Display = 'Person'; Internal = 'Person';          Type = 'User' }
)
Add-Sample -List 'Celebrations' -Rows @(
    @{ Title = 'Priya Nair';     CelebrationDate = (Get-Date).AddDays(5);  CelebrationType = 'Birthday' },
    @{ Title = 'Marcus Lee';     CelebrationDate = (Get-Date).AddDays(12); CelebrationType = 'Work anniversary' },
    @{ Title = 'Elena Petrova';  CelebrationDate = (Get-Date).AddDays(24); CelebrationType = 'Birthday' }
)

# ============================ FOOTER LINKS (site-wide footer extension) ============================
New-IntranetList -Title 'FooterLinks' -Fields @(
    @{ Display = 'Url';   Internal = 'FooterUrl';   Type = 'URL' },
    @{ Display = 'Group'; Internal = 'FooterGroup'; Type = 'Text' },
    @{ Display = 'Order'; Internal = 'FooterOrder'; Type = 'Number' }
)
Add-Sample -List 'FooterLinks' -Rows @(
    @{ Title = 'About us';       FooterUrl = 'https://contoso.sharepoint.com/sites/about, About us';   FooterGroup = 'Company';   FooterOrder = 1 },
    @{ Title = 'Newsroom';       FooterUrl = 'https://contoso.sharepoint.com/sites/news, Newsroom';     FooterGroup = 'Company';   FooterOrder = 2 },
    @{ Title = 'Careers';        FooterUrl = 'https://contoso.sharepoint.com/sites/careers, Careers';   FooterGroup = 'Company';   FooterOrder = 3 },
    @{ Title = 'IT support';     FooterUrl = 'https://contoso.sharepoint.com/sites/it, IT support';     FooterGroup = 'Support';   FooterOrder = 1 },
    @{ Title = 'HR portal';      FooterUrl = 'https://contoso.sharepoint.com/sites/hr, HR portal';      FooterGroup = 'Support';   FooterOrder = 2 },
    @{ Title = 'Service desk';   FooterUrl = 'https://contoso.sharepoint.com/sites/help, Service desk'; FooterGroup = 'Support';   FooterOrder = 3 },
    @{ Title = 'Brand assets';   FooterUrl = 'https://contoso.sharepoint.com/sites/brand, Brand';       FooterGroup = 'Resources'; FooterOrder = 1 },
    @{ Title = 'Templates';      FooterUrl = 'https://contoso.sharepoint.com/sites/templates, Templates'; FooterGroup = 'Resources'; FooterOrder = 2 },
    @{ Title = 'Privacy';        FooterUrl = 'https://contoso.sharepoint.com/sites/privacy, Privacy';   FooterGroup = 'Legal';     FooterOrder = 1 },
    @{ Title = 'Terms of use';   FooterUrl = 'https://contoso.sharepoint.com/sites/terms, Terms';       FooterGroup = 'Legal';     FooterOrder = 2 }
)

Write-Host "`nDone. Lists provisioned on $Url" -ForegroundColor Cyan
Get-PnPList | Where-Object { $_.Title -in 'Announcements','Tickets','Kudos','EmployeeOfMonth','QuickLinks','KPIs','FAQ','Events','IntranetGallery','News','Holidays','PollVotes','Celebrations','FooterLinks' } |
    Select-Object Title, ItemCount | Format-Table -AutoSize
