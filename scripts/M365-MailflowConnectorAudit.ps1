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

    Write-SectionHeader "COLLECTING MAILFLOW CONNECTORS"
    $inbound = @(Get-InboundConnector -ErrorAction SilentlyContinue)
    $outbound = @(Get-OutboundConnector -ErrorAction SilentlyContinue)

    $rows = @(
        foreach ($connector in $inbound) {
            [pscustomobject]@{
                Name          = [string]$connector.Name
                Direction     = "Inbound"
                Enabled       = if ($connector.Enabled) { "Enabled" } else { "Disabled" }
                ConnectorType = [string]$connector.ConnectorType
                TlsSettings   = [string]$connector.TlsSettings
                SmartHosts    = @($connector.SenderDomains) -join ", "
            }
        }
        foreach ($connector in $outbound) {
            [pscustomobject]@{
                Name          = [string]$connector.Name
                Direction     = "Outbound"
                Enabled       = if ($connector.Enabled) { "Enabled" } else { "Disabled" }
                ConnectorType = [string]$connector.ConnectorType
                TlsSettings   = [string]$connector.TlsSettings
                SmartHosts    = @($connector.SmartHosts) -join ", "
            }
        }
    )

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailflowConnectors" -OutputPath $OutputPath
    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mailflow Connector Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Inbound and outbound Exchange Online connectors, relay paths, and TLS posture" -Kpis @(
        @{ label = "Connectors"; value = $rows.Count; sub = "Inbound + outbound"; cls = "neutral" },
        @{ label = "Inbound"; value = $inbound.Count; sub = "Inbound connectors"; cls = "neutral" },
        @{ label = "Outbound"; value = $outbound.Count; sub = "Outbound connectors"; cls = "neutral" },
        @{ label = "Disabled"; value = @($rows | Where-Object { $_.Enabled -eq "Disabled" }).Count; sub = "Inactive connectors"; cls = "warn" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Mailflow Connectors"
            badge = "$($rows.Count) connectors"
            columns = @(
                @{ key = "Name"; header = "Connector" },
                @{ key = "Direction"; header = "Direction"; type = "pill" },
                @{ key = "Enabled"; header = "Enabled"; type = "pill" },
                @{ key = "ConnectorType"; header = "Type" },
                @{ key = "TlsSettings"; header = "TLS" },
                @{ key = "SmartHosts"; header = "Smart Hosts / Domains" }
            )
            rows = @($rows | Sort-Object Direction, Name)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
