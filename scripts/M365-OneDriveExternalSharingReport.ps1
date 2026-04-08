[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Reports.Read.All", "Sites.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING ONEDRIVE SHARING DATA"

$rows = @(Import-GraphCsvReport -RequestUri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')")
$reportRows = foreach ($row in $rows) {
    [pscustomobject]@{
        Owner        = [string]$row.'Owner Display Name'
        OwnerUPN     = [string]$row.'Owner Principal Name'
        SiteUrl      = [string]$row.'Site URL'
        UsedGB       = if ($row.'Storage Used (Byte)') { [math]::Round(([double]$row.'Storage Used (Byte)') / 1GB, 2) } else { 0 }
        FileCount    = [string]$row.'File Count'
        ActiveFiles  = [string]$row.'Active File Count'
        LastActivity = [string]$row.'Last Activity Date'
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "OneDriveExternalSharing" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 OneDrive External Sharing Report" -Tenant $tenantName -Subtitle "OneDrive usage and high-storage personal sites" -Kpis @(
    @{ label = "Sites"; value = $reportRows.Count; sub = "OneDrive sites"; cls = "neutral" },
    @{ label = "Total Storage"; value = [math]::Round((($reportRows | Measure-Object UsedGB -Sum).Sum), 2); sub = "GB used"; cls = "neutral" },
    @{ label = "Large Sites"; value = @($reportRows | Where-Object { [double]$_.UsedGB -ge 50 }).Count; sub = ">= 50 GB"; cls = "warn" },
    @{ label = "Active Owners"; value = @($reportRows | Where-Object { $_.LastActivity }).Count; sub = "With activity"; cls = "ok" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "OneDrive Sites"
        badge = "$($reportRows.Count) sites"
        columns = @(
            @{ key = "Owner"; header = "Owner" },
            @{ key = "OwnerUPN"; header = "UPN" },
            @{ key = "SiteUrl"; header = "Site URL"; type = "link" },
            @{ key = "UsedGB"; header = "Used GB" },
            @{ key = "FileCount"; header = "Files" },
            @{ key = "ActiveFiles"; header = "Active Files" },
            @{ key = "LastActivity"; header = "Last Activity" }
        )
        rows = @($reportRows | Sort-Object UsedGB -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
