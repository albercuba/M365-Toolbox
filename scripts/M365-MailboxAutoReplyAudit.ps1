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

    Write-SectionHeader "COLLECTING AUTO-REPLY DATA"
    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName)
    $rows = foreach ($mailbox in $mailboxes) {
        $config = Get-MailboxAutoReplyConfiguration -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        if (-not $config) { continue }
        [pscustomobject]@{
            DisplayName       = [string]$mailbox.DisplayName
            UserPrincipalName = [string]$mailbox.UserPrincipalName
            AutoReplyState    = [string]$config.AutoReplyState
            ExternalAudience  = [string]$config.ExternalAudience
            StartTime         = if ($config.StartTime) { (Get-Date $config.StartTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
            EndTime           = if ($config.EndTime) { (Get-Date $config.EndTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
            ExternalEnabled   = if ($config.ExternalMessage) { "Yes" } else { "No" }
        }
    }

    $enabled = @($rows | Where-Object { $_.AutoReplyState -ne "Disabled" })
    $external = @($rows | Where-Object { $_.ExternalEnabled -eq "Yes" })
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailboxAutoReply" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mailbox Auto-Reply Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Automatic reply posture across user and shared mailboxes" -Kpis @(
        @{ label = "Mailboxes"; value = $rows.Count; sub = "With auto-reply config"; cls = "neutral" },
        @{ label = "Enabled"; value = $enabled.Count; sub = "Auto-replies active"; cls = if ($enabled.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "External"; value = $external.Count; sub = "External message present"; cls = if ($external.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Scheduled"; value = @($rows | Where-Object { $_.AutoReplyState -eq "Scheduled" }).Count; sub = "Time-bounded"; cls = "neutral" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Mailbox Auto-Replies"
            badge = "$($rows.Count) mailboxes"
            columns = @(
                @{ key = "DisplayName"; header = "Mailbox" },
                @{ key = "UserPrincipalName"; header = "UPN" },
                @{ key = "AutoReplyState"; header = "State"; type = "pill" },
                @{ key = "ExternalAudience"; header = "External Audience"; type = "pill" },
                @{ key = "ExternalEnabled"; header = "External Message"; type = "pill" },
                @{ key = "StartTime"; header = "Start" },
                @{ key = "EndTime"; header = "End" }
            )
            rows = @($rows | Sort-Object AutoReplyState, DisplayName)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
