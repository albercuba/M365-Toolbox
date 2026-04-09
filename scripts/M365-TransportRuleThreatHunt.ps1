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
    Write-SectionHeader "HUNTING TRANSPORT RULES"

    $rows = foreach ($rule in @(Get-TransportRule)) {
        $signals = [System.Collections.Generic.List[string]]::new()
        if ($rule.RedirectMessageTo -or $rule.BlindCopyTo) { [void]$signals.Add("Redirect or blind copy") }
        if ($rule.DeleteMessage) { [void]$signals.Add("Deletes message") }
        if ($rule.SetSCL -and [int]$rule.SetSCL -le -1) { [void]$signals.Add("Bypasses spam filtering") }
        if ($rule.StopRuleProcessing) { [void]$signals.Add("Stops further processing") }
        if ($rule.RouteMessageOutboundConnector) { [void]$signals.Add("Routes via outbound connector") }
        if ($signals.Count -eq 0) { continue }

        [pscustomobject]@{
            Name     = [string]$rule.Name
            State    = if ($rule.State) { [string]$rule.State } else { "Enabled" }
            Priority = [string]$rule.Priority
            Mode     = [string]$rule.Mode
            Signals  = ($signals -join "; ")
        }
    }

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "TransportRuleThreatHunt" -OutputPath $OutputPath
    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Transport Rule Threat Hunt" -Tenant $script:ToolboxTenantLabel -Subtitle "Transport rules with bypass, redirect, delete, or routing threat signals" -Kpis @(
        @{ label = "Threat Rules"; value = @($rows).Count; sub = "Rules with risky actions"; cls = if (@($rows).Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Enabled"; value = @($rows | Where-Object { $_.State -ne 'Disabled' }).Count; sub = "Enabled risky rules"; cls = "crit" },
        @{ label = "Audit/Test"; value = @($rows | Where-Object { $_.Mode -match 'Audit' }).Count; sub = "Rules in test mode"; cls = "warn" },
        @{ label = "Disabled"; value = @($rows | Where-Object { $_.State -eq 'Disabled' }).Count; sub = "Disabled risky rules"; cls = "neutral" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Transport Rule Threat Signals"
            badge = "$(@($rows).Count) rule(s)"
            columns = @(
                @{ key = "Name"; header = "Rule" },
                @{ key = "State"; header = "State"; type = "pill" },
                @{ key = "Priority"; header = "Priority" },
                @{ key = "Mode"; header = "Mode"; type = "pill" },
                @{ key = "Signals"; header = "Signals" }
            )
            rows = @($rows | Sort-Object Priority)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
