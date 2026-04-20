[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [ValidateSet("D7", "D30", "D90", "D180")]
    [string]$ReportPeriod = "D30"
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Sites.Read.All", "Reports.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING EXTERNAL SHARING DATA"

$settings = $null
try { $settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -ErrorAction Stop } catch {}
$siteUsage = @(Import-GraphCsvReport -RequestUri ("https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='{0}')" -f $ReportPeriod))

$sharingCapability = [string](Get-DirectoryObjectValue -DirectoryObject $settings -Name 'sharingCapability')
$defaultSharingLinkType = [string](Get-DirectoryObjectValue -DirectoryObject $settings -Name 'defaultSharingLinkType')
$defaultLinkPermission = [string](Get-DirectoryObjectValue -DirectoryObject $settings -Name 'defaultLinkPermission')
$isResharingByExternalUsersEnabled = Get-DirectoryObjectValue -DirectoryObject $settings -Name 'isResharingByExternalUsersEnabled'

$rows = foreach ($site in $siteUsage) {
    $ownerValue = [string](Get-DirectoryObjectValue -DirectoryObject $site -Name 'OwnerDisplayName')
    if (-not $ownerValue) {
        $ownerValue = [string](Get-DirectoryObjectValue -DirectoryObject $site -Name 'Owner Display Name')
    }

    [pscustomobject]@{
        SiteUrl      = [string]$site.'Site URL'
        Owner        = $ownerValue
        StorageGB    = if ($site.'Storage Used (Byte)') { [math]::Round(([double]$site.'Storage Used (Byte)') / 1GB, 2) } else { 0 }
        ActiveFiles  = [string]$site.'Active File Count'
        LastActivity = [string]$site.'Last Activity Date'
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ExternalSharingLinks" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 External Sharing Links Report" -Tenant $tenantName -Subtitle "SharePoint external sharing posture and active sites" -Kpis @(
    @{ label = "Sites"; value = $rows.Count; sub = "Reported sites"; cls = "neutral" },
    @{ label = "Sharing Mode"; value = if ($sharingCapability) { $sharingCapability } else { "Unknown" }; sub = "Tenant capability"; cls = "warn" },
    @{ label = "Default Link"; value = if ($defaultSharingLinkType) { $defaultSharingLinkType } else { "Unknown" }; sub = "Link type"; cls = "neutral" },
    @{ label = "Period"; value = $ReportPeriod; sub = "Usage lookback"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Anonymous Access"; value = if ($sharingCapability -match "externalUserAndGuestSharing|anyone") { "Potentially Allowed" } else { "Restricted" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Tenant Sharing Configuration"
        badge = "SharePoint"
        columns = @(
            @{ key = "Setting"; header = "Setting" },
            @{ key = "Value"; header = "Value"; type = "pill" }
        )
        rows = @(
            [pscustomobject]@{ Setting = "Sharing Capability"; Value = if ($sharingCapability) { $sharingCapability } else { "Unknown" } },
            [pscustomobject]@{ Setting = "Default Sharing Link Type"; Value = if ($defaultSharingLinkType) { $defaultSharingLinkType } else { "Unknown" } },
            [pscustomobject]@{ Setting = "Default Link Permission"; Value = if ($defaultLinkPermission) { $defaultLinkPermission } else { "Unknown" } },
            [pscustomobject]@{ Setting = "External Resharing"; Value = if ($isResharingByExternalUsersEnabled) { "Enabled" } else { "Disabled" } }
        )
    },
    @{
        title = "Active Sites"
        badge = "$($rows.Count) sites"
        columns = @(
            @{ key = "SiteUrl"; header = "Site URL"; type = "link" },
            @{ key = "Owner"; header = "Owner" },
            @{ key = "StorageGB"; header = "Storage GB" },
            @{ key = "ActiveFiles"; header = "Active Files" },
            @{ key = "LastActivity"; header = "Last Activity" }
        )
        rows = @($rows | Sort-Object StorageGB -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
