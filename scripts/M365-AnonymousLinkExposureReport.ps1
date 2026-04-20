[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [string]$ReportPeriod = "D30"
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports", "Microsoft.Graph.Sites")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Reports.Read.All", "Directory.Read.All", "Sites.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING ANONYMOUS LINK EXPOSURE"

function Get-SiteUsageValue {
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $value = Get-DirectoryObjectValue -DirectoryObject $Row -Name $name
        if ($null -ne $value -and [string]$value -ne '') {
            return $value
        }
    }

    return $null
}

$settings = $null
try { $settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -ErrorAction Stop } catch {}

$sites = @(Import-GraphCsvReport -RequestUri ("https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='{0}')" -f $ReportPeriod))
$siteUrlCache = @{}
$rows = foreach ($site in $sites) {
    $siteId = [string](Get-SiteUsageValue -Row $site -Names @('Site Id', 'SiteId'))
    $siteUrl = [string](Get-SiteUsageValue -Row $site -Names @('Site URL', 'SiteUrl'))
    if (-not $siteUrl -and $siteId) {
        if (-not $siteUrlCache.ContainsKey($siteId)) {
            try {
                $siteLookup = Get-MgSite -SiteId $siteId -Property WebUrl -ErrorAction Stop
                $siteUrlCache[$siteId] = [string]$siteLookup.WebUrl
            }
            catch {
                $siteUrlCache[$siteId] = ''
            }
        }

        $siteUrl = [string]$siteUrlCache[$siteId]
    }

    [pscustomobject]@{
        SiteUrl      = $siteUrl
        Owner        = [string](Get-SiteUsageValue -Row $site -Names @('Owner Display Name', 'OwnerDisplayName', 'Owner Principal Name'))
        FileCount    = [int](Get-SiteUsageValue -Row $site -Names @('File Count', 'FileCount'))
        ActiveFiles  = [int](Get-SiteUsageValue -Row $site -Names @('Active File Count', 'ActiveFileCount'))
        StorageGB    = if ($site.'Storage Used (Byte)') { [math]::Round(([double]$site.'Storage Used (Byte)') / 1GB, 2) } else { 0 }
        LastActivity = [string](Get-SiteUsageValue -Row $site -Names @('Last Activity Date', 'LastActivityDate'))
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
            @{ key = "SiteUrl"; header = "Site URL"; type = "link" },
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
