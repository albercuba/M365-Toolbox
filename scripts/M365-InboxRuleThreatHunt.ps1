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
    Write-SectionHeader "HUNTING INBOX RULES"

    $mailboxes = @(Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object -First $MaxMailboxesToInspect)
    $rows = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $mailboxes) {
        $rules = @(Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue)
        foreach ($rule in $rules) {
            $signals = [System.Collections.Generic.List[string]]::new()
            $targets = @($rule.ForwardTo + $rule.ForwardAsAttachmentTo + $rule.RedirectTo | Where-Object { $_ })
            if ($targets.Count -gt 0) { [void]$signals.Add("External or delegated forwarding") }
            if ($rule.DeleteMessage) { [void]$signals.Add("Deletes messages") }
            if ($rule.MarkAsRead) { [void]$signals.Add("Marks as read") }
            if ($rule.MoveToFolder) { [void]$signals.Add("Moves messages") }
            if ($rule.StopProcessingRules) { [void]$signals.Add("Stops further processing") }
            if ($signals.Count -eq 0) { continue }

            [void]$rows.Add([pscustomobject]@{
                Mailbox  = [string]$mailbox.UserPrincipalName
                RuleName = [string]$rule.Name
                Enabled  = if ($rule.Enabled) { "Enabled" } else { "Disabled" }
                Targets  = if ($targets.Count -gt 0) { ($targets -join ", ") } else { "None" }
                Signals  = ($signals -join "; ")
            })
        }
    }

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "InboxRuleThreatHunt" -OutputPath $OutputPath
    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Inbox Rule Threat Hunt" -Tenant $script:ToolboxTenantLabel -Subtitle "Mailbox rules with suspicious forwarding, deletion, or concealment behavior" -Kpis @(
        @{ label = "Mailboxes"; value = $mailboxes.Count; sub = "Inspected user mailboxes"; cls = "neutral" },
        @{ label = "Suspicious Rules"; value = $rows.Count; sub = "Rules with hunt signals"; cls = if ($rows.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Affected Mailboxes"; value = @($rows | Select-Object -ExpandProperty Mailbox -Unique).Count; sub = "Mailboxes with risky rules"; cls = if ($rows.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Forwarding Rules"; value = @($rows | Where-Object { $_.Targets -ne 'None' }).Count; sub = "Rules with redirect or forwarding"; cls = "crit" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Suspicious Inbox Rules"
            badge = "$($rows.Count) rule(s)"
            columns = @(
                @{ key = "Mailbox"; header = "Mailbox" },
                @{ key = "RuleName"; header = "Rule" },
                @{ key = "Enabled"; header = "State"; type = "pill" },
                @{ key = "Targets"; header = "Targets" },
                @{ key = "Signals"; header = "Signals" }
            )
            rows = @($rows | Sort-Object Mailbox, RuleName)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
