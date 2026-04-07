[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
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

    Write-SectionHeader "COLLECTING MAILBOX PERMISSION DATA"
    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress)
    $rows = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $mailboxes) {
        $mailboxPerms = @(Get-ExoMailboxPermission -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object { $_.User -notmatch "NT AUTHORITY|S-1-5|SELF" -and -not $_.IsInherited })
        $recipientPerms = @(Get-RecipientPermission -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object { $_.Trustee -and $_.Trustee -notmatch "NT AUTHORITY|S-1-5|SELF" })

        [void]$rows.Add([pscustomobject]@{
            DisplayName       = [string]$mailbox.DisplayName
            UserPrincipalName = [string]$mailbox.UserPrincipalName
            FullAccess        = $mailboxPerms.Count
            SendAs            = $recipientPerms.Count
            Exposure          = if (($mailboxPerms.Count + $recipientPerms.Count) -gt 3) { "Elevated" } elseif (($mailboxPerms.Count + $recipientPerms.Count) -gt 0) { "Delegated" } else { "Direct Owner Only" }
        })
    }

    $delegated = @($rows | Where-Object { $_.Exposure -ne "Direct Owner Only" })
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailboxPermissions" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mailbox Permission Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Mailbox delegate exposure across user and shared mailboxes" -Kpis @(
        @{ label = "Mailboxes"; value = $rows.Count; sub = "Scoped mailboxes"; cls = "neutral" },
        @{ label = "Delegated"; value = $delegated.Count; sub = "Mailboxes with delegates"; cls = if ($delegated.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Full Access"; value = (($rows | Measure-Object FullAccess -Sum).Sum); sub = "Assignments"; cls = "neutral" },
        @{ label = "Send As"; value = (($rows | Measure-Object SendAs -Sum).Sum); sub = "Assignments"; cls = "neutral" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Mailbox Permission Summary"
            badge = "$($rows.Count) mailboxes"
            columns = @(
                @{ key = "DisplayName"; header = "Mailbox" },
                @{ key = "UserPrincipalName"; header = "UPN" },
                @{ key = "FullAccess"; header = "Full Access" },
                @{ key = "SendAs"; header = "Send As" },
                @{ key = "Exposure"; header = "Exposure"; type = "pill" }
            )
            rows = @($rows | Sort-Object FullAccess -Descending)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
