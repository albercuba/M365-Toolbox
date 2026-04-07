[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [switch]$IncludeInboxRules
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Write-Host ""
Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
Write-Host "  [+] ExchangeOnlineManagement" -ForegroundColor Green

Write-SectionHeader "CONNECTING TO EXCHANGE ONLINE"
if ($TenantId) {
    Write-Host "[*] Requested tenant: $TenantId"
}
else {
    Write-Host "[*] Requested tenant: auto-detect"
}
Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
$exoParams = @{
    ShowBanner   = $false
    ShowProgress = $false
    Device       = $true
}
if ($TenantId) {
    $exoParams.Organization = $TenantId
}
Connect-ExchangeOnline @exoParams | Out-Null

try {
    Write-Host "[+] Connected to Exchange Online" -ForegroundColor Green
    $script:ToolboxTenantLabel = if ($TenantId) { $TenantId } else { "ExchangeOnline" }

    Write-SectionHeader "COLLECTING MAIL FORWARDING DATA"
    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties ForwardingAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward,DisplayName,UserPrincipalName,PrimarySmtpAddress)

    $forwardRows = [System.Collections.Generic.List[object]]::new()
    $ruleRows = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $mailboxes) {
        $hasMailboxForward = [bool]($mailbox.ForwardingAddress -or $mailbox.ForwardingSmtpAddress)
        if ($hasMailboxForward) {
            [void]$forwardRows.Add([pscustomobject]@{
                DisplayName       = [string]$mailbox.DisplayName
                UserPrincipalName = [string]$mailbox.UserPrincipalName
                ForwardTarget     = if ($mailbox.ForwardingSmtpAddress) { [string]$mailbox.ForwardingSmtpAddress } else { [string]$mailbox.ForwardingAddress }
                DeliverAndForward = if ($mailbox.DeliverToMailboxAndForward) { "Yes" } else { "No" }
                ExternalTarget    = if ([string]$mailbox.ForwardingSmtpAddress -match "@") { "Yes" } else { "No" }
            })
        }

        if ($IncludeInboxRules) {
            try {
                $rules = @(Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop)
                foreach ($rule in $rules) {
                    $targets = @($rule.ForwardTo + $rule.RedirectTo + $rule.ForwardAsAttachmentTo | Where-Object { $_ })
                    if ($targets.Count -eq 0) { continue }

                    [void]$ruleRows.Add([pscustomobject]@{
                        UserPrincipalName = [string]$mailbox.UserPrincipalName
                        RuleName          = [string]$rule.Name
                        Enabled           = if ($rule.Enabled) { "Enabled" } else { "Disabled" }
                        Targets           = ($targets -join ", ")
                    })
                }
            }
            catch {
            }
        }
    }

    $externalMailboxForward = @($forwardRows | Where-Object { $_.ExternalTarget -eq "Yes" }).Count
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailForwarding" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mail Forwarding Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Mailbox forwarding settings and inbox redirect rules" -Kpis @(
        @{ label = "Mailboxes"; value = $mailboxes.Count; sub = "User mailboxes"; cls = "neutral" },
        @{ label = "Forwarding"; value = $forwardRows.Count; sub = "Mailbox-level forwarding"; cls = if ($forwardRows.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "External"; value = $externalMailboxForward; sub = "External mailbox forwarding"; cls = if ($externalMailboxForward -gt 0) { "crit" } else { "ok" } },
        @{ label = "Inbox Rules"; value = $ruleRows.Count; sub = if ($IncludeInboxRules) { "Rules with forwarding" } else { "Not collected" }; cls = if ($ruleRows.Count -gt 0) { "warn" } else { "neutral" } }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Inbox Rules"; value = if ($IncludeInboxRules) { "Included" } else { "Skipped" } },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Mailbox Forwarding"
            badge = "$($forwardRows.Count) mailboxes"
            columns = @(
                @{ key = "DisplayName"; header = "Name" },
                @{ key = "UserPrincipalName"; header = "UPN" },
                @{ key = "ForwardTarget"; header = "Target" },
                @{ key = "DeliverAndForward"; header = "Deliver & Forward"; type = "pill" },
                @{ key = "ExternalTarget"; header = "External"; type = "pill" }
            )
            rows = @($forwardRows | Sort-Object UserPrincipalName)
        },
        @{
            title = "Inbox Rules With Forwarding"
            badge = "$($ruleRows.Count) rules"
            columns = @(
                @{ key = "UserPrincipalName"; header = "UPN" },
                @{ key = "RuleName"; header = "Rule" },
                @{ key = "Enabled"; header = "Enabled"; type = "pill" },
                @{ key = "Targets"; header = "Targets" }
            )
            rows = @($ruleRows | Sort-Object UserPrincipalName)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
