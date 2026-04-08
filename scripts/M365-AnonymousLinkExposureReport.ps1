[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [string]$ReportPeriod = "D30"
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Reports.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING ANONYMOUS LINK EXPOSURE"

$settings = $null
try { $settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -ErrorAction Stop } catch {}

$sites = @(Import-GraphCsvReport -RequestUri ("https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='{0}')" -f $ReportPeriod))
$rows = foreach ($site in $sites) {
    [pscustomobject]@{
        SiteUrl      = [string]$site.'Site URL'
        Owner        = [string]$site.'Owner Display Name'
        FileCount    = [int]($site.'File Count')
        ActiveFiles  = [int]($site.'Active File Count')
        StorageGB    = if ($site.'Storage Used (Byte)') { [math]::Round(([double]$site.'Storage Used (Byte)') / 1GB, 2) } else { 0 }
        LastActivity = [string]$site.'Last Activity Date'
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "AnonymousLinkExposure" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Anonymous Link Exposure Report" -Tenant $tenantName -Subtitle "Tenant sharing defaults and active SharePoint footprint that may amplify anyone-link risk" -Kpis @(
    @{ label = "Sites"; value = $rows.Count; sub = "Reported sites"; cls = "neutral" },
    @{ label = "Sharing Capability"; value = if ($settings -and $settings.sharingCapability) { [string]$settings.sharingCapability } else { "Unknown" }; sub = "Tenant default"; cls = "warn" },
    @{ label = "Default Link Type"; value = if ($settings -and $settings.defaultSharingLinkType) { [string]$settings.defaultSharingLinkType } else { "Unknown" }; sub = "Link behavior"; cls = "neutral" },
    @{ label = "Large Sites"; value = @($rows | Where-Object { $_.StorageGB -ge 10 }).Count; sub = "10GB+"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Report Period"; value = $ReportPeriod },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Tenant Sharing Settings"
        badge = "SharePoint"
        columns = @(
            @{ key = "Setting"; header = "Setting" },
            @{ key = "Value"; header = "Value"; type = "pill" }
        )
        rows = @(
            [pscustomobject]@{ Setting = "Sharing Capability"; Value = if ($settings -and $settings.sharingCapability) { [string]$settings.sharingCapability } else { "Unknown" } },
            [pscustomobject]@{ Setting = "Default Sharing Link Type"; Value = if ($settings -and $settings.defaultSharingLinkType) { [string]$settings.defaultSharingLinkType } else { "Unknown" } },
            [pscustomobject]@{ Setting = "Default Link Permission"; Value = if ($settings -and $settings.defaultLinkPermission) { [string]$settings.defaultLinkPermission } else { "Unknown" } },
            [pscustomobject]@{ Setting = "External Resharing"; Value = if ($settings -and $settings.isResharingByExternalUsersEnabled) { "Enabled" } else { "Disabled / Unknown" } }
        )
    },
    @{
        title = "Site Footprint"
        badge = "$($rows.Count) sites"
        columns = @(
            @{ key = "SiteUrl"; header = "Site URL" },
            @{ key = "Owner"; header = "Owner" },
            @{ key = "StorageGB"; header = "Storage (GB)" },
            @{ key = "FileCount"; header = "Files" },
            @{ key = "ActiveFiles"; header = "Active Files" },
            @{ key = "LastActivity"; header = "Last Activity" }
        )
        rows = @($rows | Sort-Object StorageGB -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
