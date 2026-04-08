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
    Write-SectionHeader "COLLECTING MAIL TRANSPORT RULES"

    $rules = @(Get-TransportRule)
    $rows = foreach ($rule in $rules) {
        [pscustomobject]@{
            Name        = [string]$rule.Name
            State       = if ($rule.State) { [string]$rule.State } else { "Enabled" }
            Priority    = [string]$rule.Priority
            Mode        = [string]$rule.Mode
            Description = [string]$rule.Description
        }
    }

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailTransportRules" -OutputPath $OutputPath
    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mail Transport Rules Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Exchange Online transport rules and mail flow logic" -Kpis @(
        @{ label = "Rules"; value = $rows.Count; sub = "Transport rules"; cls = "neutral" },
        @{ label = "Enabled"; value = @($rows | Where-Object { $_.State -ne "Disabled" }).Count; sub = "Active rules"; cls = "ok" },
        @{ label = "Disabled"; value = @($rows | Where-Object { $_.State -eq "Disabled" }).Count; sub = "Inactive rules"; cls = "warn" },
        @{ label = "Test"; value = @($rows | Where-Object { $_.Mode -eq "Audit" -or $_.Mode -eq "AuditAndNotify" }).Count; sub = "Audit/test mode"; cls = "warn" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Transport Rules"
            badge = "$($rows.Count) rules"
            columns = @(
                @{ key = "Name"; header = "Rule" },
                @{ key = "State"; header = "State"; type = "pill" },
                @{ key = "Priority"; header = "Priority" },
                @{ key = "Mode"; header = "Mode"; type = "pill" },
                @{ key = "Description"; header = "Description" }
            )
            rows = @($rows | Sort-Object Priority)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
