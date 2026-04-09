[CmdletBinding()]
param(
    [string]$TenantId,
    [int]$MaxMailboxesToInspect = 150,
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
    Write-SectionHeader "COLLECTING SHARED MAILBOX ABUSE SIGNALS"

    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select-Object -First $MaxMailboxesToInspect)
    $rows = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $mailboxes) {
        $delegateCount = @(
            Get-MailboxPermission -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue |
                Where-Object { -not $_.IsInherited -and $_.User -notmatch 'NT AUTHORITY|S-1-5-|SELF' }
        ).Count
        $ruleCount = @(Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue).Count
        $signals = [System.Collections.Generic.List[string]]::new()
        if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) { [void]$signals.Add("Forwarding configured") }
        if ($delegateCount -gt 5) { [void]$signals.Add("High delegate count") }
        if ($ruleCount -gt 0) { [void]$signals.Add("Inbox rules present") }
        if ($signals.Count -eq 0) { [void]$signals.Add("Review shared access posture") }

        [void]$rows.Add([pscustomobject]@{
            SharedMailbox = [string]$mailbox.UserPrincipalName
            DisplayName   = [string]$mailbox.DisplayName
            Delegates     = $delegateCount
            InboxRules    = $ruleCount
            Forwarding    = if ($mailbox.ForwardingSmtpAddress) { [string]$mailbox.ForwardingSmtpAddress } elseif ($mailbox.ForwardingAddress) { [string]$mailbox.ForwardingAddress } else { "None" }
            Signals       = ($signals -join "; ")
        })
    }

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "SharedMailboxAbuseReview" -OutputPath $OutputPath
    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Shared Mailbox Abuse Review" -Tenant $script:ToolboxTenantLabel -Subtitle "Forwarding, delegate, and inbox rule signals across shared mailboxes" -Kpis @(
        @{ label = "Shared Mailboxes"; value = $mailboxes.Count; sub = "Shared mailboxes reviewed"; cls = "neutral" },
        @{ label = "With Forwarding"; value = @($rows | Where-Object { $_.Forwarding -ne 'None' }).Count; sub = "Forwarding configured"; cls = "crit" },
        @{ label = "High Delegates"; value = @($rows | Where-Object { $_.Delegates -gt 5 }).Count; sub = "Shared mailboxes with many delegates"; cls = "warn" },
        @{ label = "Rules Present"; value = @($rows | Where-Object { $_.InboxRules -gt 0 }).Count; sub = "Shared mailboxes with inbox rules"; cls = "warn" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Shared Mailbox Abuse Signals"
            badge = "$($rows.Count) mailbox(es)"
            columns = @(
                @{ key = "SharedMailbox"; header = "Shared Mailbox" },
                @{ key = "DisplayName"; header = "Display Name" },
                @{ key = "Delegates"; header = "Delegates" },
                @{ key = "InboxRules"; header = "Inbox Rules" },
                @{ key = "Forwarding"; header = "Forwarding" },
                @{ key = "Signals"; header = "Signals" }
            )
            rows = @($rows | Sort-Object SharedMailbox)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
