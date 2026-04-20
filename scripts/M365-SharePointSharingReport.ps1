[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [ValidateSet("D7", "D30", "D90", "D180")]
    [string]$ReportPeriod = "D30"
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Sites.Read.All", "Reports.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING SHAREPOINT SHARING DATA"

$settings = $null
try {
    $settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -ErrorAction Stop
}
catch {
}

$siteUsage = @(Import-GraphCsvReport -RequestUri ("https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='{0}')" -f $ReportPeriod))
$siteRows = foreach ($site in $siteUsage) {
    [pscustomobject]@{
        SiteUrl      = [string]$site.'Site URL'
        Owner        = [string]$site.OwnerDisplayName
        Files        = [string]$site.'File Count'
        ActiveFiles  = [string]$site.'Active File Count'
        StorageGB    = if ($site.'Storage Used (Byte)') { [math]::Round(([double]$site.'Storage Used (Byte)') / 1GB, 2) } else { 0 }
        LastActivity = [string]$site.'Last Activity Date'
    }
}

$settingRows = @(
    [pscustomobject]@{ Setting = "Tenant Sharing Capability"; Value = [string]$settings.sharingCapability },
    [pscustomobject]@{ Setting = "Default Link Type"; Value = [string]$settings.defaultSharingLinkType },
    [pscustomobject]@{ Setting = "Default Permission"; Value = [string]$settings.defaultLinkPermission },
    [pscustomobject]@{ Setting = "External Resharing"; Value = if ($settings.isResharingByExternalUsersEnabled) { "Enabled" } else { "Disabled" } },
    [pscustomobject]@{ Setting = "Domain Restriction Mode"; Value = [string]$settings.sharingDomainRestrictionMode }
)

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "SharePointSharing" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 SharePoint Sharing Report" -Tenant $tenantName -Subtitle "Tenant sharing posture and site usage inventory" -Kpis @(
    @{ label = "Sites"; value = $siteRows.Count; sub = "Sites in report"; cls = "neutral" },
    @{ label = "Storage"; value = [math]::Round((($siteRows | Measure-Object StorageGB -Sum).Sum), 2); sub = "GB used"; cls = "neutral" },
    @{ label = "Period"; value = $ReportPeriod; sub = "Usage lookback"; cls = "neutral" },
    @{ label = "Sharing"; value = if ($settings.sharingCapability) { [string]$settings.sharingCapability } else { "Unknown" }; sub = "Tenant capability"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Report Period"; value = $ReportPeriod },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Tenant Sharing Settings"
        badge = "Tenant posture"
        columns = @(
            @{ key = "Setting"; header = "Setting" },
            @{ key = "Value"; header = "Value"; type = "pill" }
        )
        rows = $settingRows
    },
    @{
        title = "SharePoint Site Inventory"
        badge = "$($siteRows.Count) sites"
        columns = @(
            @{ key = "SiteUrl"; header = "Site URL"; type = "link" },
            @{ key = "Owner"; header = "Owner" },
            @{ key = "Files"; header = "Files" },
            @{ key = "ActiveFiles"; header = "Active Files" },
            @{ key = "StorageGB"; header = "Storage GB" },
            @{ key = "LastActivity"; header = "Last Activity" }
        )
        rows = @($siteRows | Sort-Object StorageGB -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
