[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$MaxMailboxesToInspect = 150
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Write-Host ""
Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
Write-Host "  [+] ExchangeOnlineManagement" -ForegroundColor Green

Write-SectionHeader "CONNECTING TO EXCHANGE ONLINE"
if ($TenantId) { Write-Host "[*] Requested tenant: $TenantId" } else { Write-Host "[*] Requested tenant: auto-detect" }
Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
$exoParams = @{ ShowBanner = $false; ShowProgress = $false; Device = $true }
if ($TenantId) { $exoParams.Organization = $TenantId }
Connect-ExchangeOnline @exoParams | Out-Null

try {
    $script:ToolboxTenantLabel = if ($TenantId) { $TenantId } else { "ExchangeOnline" }
    Write-Host "[+] Connected to Exchange Online" -ForegroundColor Green

    Write-SectionHeader "COLLECTING CALENDAR SHARING DATA"
    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName | Select-Object -First $MaxMailboxesToInspect)
    $rows = foreach ($mailbox in $mailboxes) {
        $calendarId = "{0}:\Calendar" -f $mailbox.UserPrincipalName
        $permissions = @(Get-MailboxFolderPermission -Identity $calendarId -ErrorAction SilentlyContinue | Where-Object { $_.User -and $_.User -notmatch "Default|Anonymous" })
        $defaultPerm = Get-MailboxFolderPermission -Identity $calendarId -User Default -ErrorAction SilentlyContinue
        $anonymousPerm = Get-MailboxFolderPermission -Identity $calendarId -User Anonymous -ErrorAction SilentlyContinue
        [pscustomobject]@{
            DisplayName       = [string]$mailbox.DisplayName
            UserPrincipalName = [string]$mailbox.UserPrincipalName
            Delegates         = $permissions.Count
            DefaultAccess     = if ($defaultPerm) { [string]$defaultPerm.AccessRights } else { "Unknown" }
            AnonymousAccess   = if ($anonymousPerm) { [string]$anonymousPerm.AccessRights } else { "None" }
            Exposure          = if ($anonymousPerm -and [string]$anonymousPerm.AccessRights -notmatch "None|AvailabilityOnly") { "Anonymous" } elseif ($permissions.Count -gt 0) { "Delegated" } else { "Restricted" }
        }
    }

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "CalendarSharingAudit" -OutputPath $OutputPath
    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Calendar Sharing Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Calendar sharing posture across inspected user mailboxes" -Kpis @(
        @{ label = "Mailboxes"; value = $rows.Count; sub = "Inspected calendars"; cls = "neutral" },
        @{ label = "Delegated"; value = @($rows | Where-Object { $_.Delegates -gt 0 }).Count; sub = "Named delegates"; cls = "warn" },
        @{ label = "Anonymous"; value = @($rows | Where-Object { $_.Exposure -eq "Anonymous" }).Count; sub = "Anonymous exposure"; cls = "crit" },
        @{ label = "Inspection Limit"; value = $MaxMailboxesToInspect; sub = "User mailboxes"; cls = "neutral" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Calendar Sharing"
            badge = "$($rows.Count) calendars"
            columns = @(
                @{ key = "DisplayName"; header = "Mailbox" },
                @{ key = "UserPrincipalName"; header = "UPN" },
                @{ key = "Delegates"; header = "Delegates" },
                @{ key = "DefaultAccess"; header = "Default Access" },
                @{ key = "AnonymousAccess"; header = "Anonymous Access" },
                @{ key = "Exposure"; header = "Exposure"; type = "pill" }
            )
            rows = @($rows | Sort-Object Exposure, Delegates -Descending)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
