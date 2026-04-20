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
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Sites.Read.All", "Reports.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING EXTERNAL SHARING DATA"

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

function Get-SafeIntValue {
    param($Value)

    if ($null -eq $Value -or [string]$Value -eq '') {
        return 0
    }

    try {
        return [int]$Value
    }
    catch {
        return 0
    }
}

$settings = $null
try { $settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -ErrorAction Stop } catch {}
$siteUsage = @(Import-GraphCsvReport -RequestUri ("https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='{0}')" -f $ReportPeriod))

$sharingCapability = [string](Get-DirectoryObjectValue -DirectoryObject $settings -Name 'sharingCapability')
$defaultSharingLinkType = [string](Get-DirectoryObjectValue -DirectoryObject $settings -Name 'defaultSharingLinkType')
$defaultLinkPermission = [string](Get-DirectoryObjectValue -DirectoryObject $settings -Name 'defaultLinkPermission')
$isResharingByExternalUsersEnabled = Get-DirectoryObjectValue -DirectoryObject $settings -Name 'isResharingByExternalUsersEnabled'
$siteUrlCache = @{}

$rows = foreach ($site in $siteUsage) {
    $siteId = [string](Get-SiteUsageValue -Row $site -Names @('Site Id', 'SiteId'))
    $siteUrl = [string](Get-SiteUsageValue -Row $site -Names @('Site URL', 'SiteUrl'))
    if (-not $siteUrl -and $siteId) {
        if (-not $siteUrlCache.ContainsKey($siteId)) {
            try {
                $siteLookup = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}?`$select=webUrl" -f $siteId) -ErrorAction Stop
                $siteUrlCache[$siteId] = [string](Get-DirectoryObjectValue -DirectoryObject $siteLookup -Name 'webUrl')
            }
            catch {
                $siteUrlCache[$siteId] = ''
            }
        }

        $siteUrl = [string]$siteUrlCache[$siteId]
    }

    $ownerValue = [string](Get-SiteUsageValue -Row $site -Names @('OwnerDisplayName', 'Owner Display Name'))
    if (-not $ownerValue) {
        $ownerValue = [string](Get-SiteUsageValue -Row $site -Names @('Owner Principal Name', 'OwnerPrincipalName'))
    }

    [pscustomobject]@{
        SiteId                = $siteId
        SiteUrl               = $siteUrl
        Owner                 = $ownerValue
        ExternalSharing       = [string](Get-SiteUsageValue -Row $site -Names @('External Sharing', 'ExternalSharing'))
        AnonymousLinkCount    = Get-SafeIntValue -Value (Get-SiteUsageValue -Row $site -Names @('Anonymous Link Count', 'AnonymousLinkCount'))
        CompanyLinkCount      = Get-SafeIntValue -Value (Get-SiteUsageValue -Row $site -Names @('Company Link Count', 'CompanyLinkCount'))
        GuestSecureLinkCount  = Get-SafeIntValue -Value (Get-SiteUsageValue -Row $site -Names @('Secure Link For Guest Count', 'SecureLinkForGuestCount'))
        MemberSecureLinkCount = Get-SafeIntValue -Value (Get-SiteUsageValue -Row $site -Names @('Secure Link For Member Count', 'SecureLinkForMemberCount'))
        StorageGB             = if ($site.'Storage Used (Byte)') { [math]::Round(([double]$site.'Storage Used (Byte)') / 1GB, 2) } else { 0 }
        ActiveFiles           = [string](Get-SiteUsageValue -Row $site -Names @('Active File Count', 'ActiveFileCount'))
        LastActivity          = [string](Get-SiteUsageValue -Row $site -Names @('Last Activity Date', 'LastActivityDate'))
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ExternalSharingLinks" -OutputPath $OutputPath
$sitesWithSharingLinks = @($rows | Where-Object { ($_.AnonymousLinkCount + $_.CompanyLinkCount + $_.GuestSecureLinkCount + $_.MemberSecureLinkCount) -gt 0 })
$totalAnonymousLinks = (($rows | Measure-Object AnonymousLinkCount -Sum).Sum)
$totalCompanyLinks = (($rows | Measure-Object CompanyLinkCount -Sum).Sum)
$totalSpecificPeopleLinks = (($rows | Measure-Object GuestSecureLinkCount -Sum).Sum) + (($rows | Measure-Object MemberSecureLinkCount -Sum).Sum)

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 External Sharing Links Report" -Tenant $tenantName -Subtitle "SharePoint sharing posture and site-level sharing link counts" -Kpis @(
    @{ label = "Sites"; value = $rows.Count; sub = "Reported sites"; cls = "neutral" },
    @{ label = "Sites With Links"; value = $sitesWithSharingLinks.Count; sub = "Sites with reported sharing links"; cls = if ($sitesWithSharingLinks.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Anyone Links"; value = $totalAnonymousLinks; sub = "Anonymous links"; cls = if ($totalAnonymousLinks -gt 0) { "crit" } else { "ok" } },
    @{ label = "Specific/Guest Links"; value = $totalSpecificPeopleLinks; sub = "Secure guest/member links"; cls = if ($totalSpecificPeopleLinks -gt 0) { "warn" } else { "ok" } },
    @{ label = "Sharing Mode"; value = if ($sharingCapability) { $sharingCapability } else { "Unknown" }; sub = "Tenant capability"; cls = "warn" },
    @{ label = "Default Link"; value = if ($defaultSharingLinkType) { $defaultSharingLinkType } else { "Unknown" }; sub = "Default link type"; cls = "neutral" },
    @{ label = "Period"; value = $ReportPeriod; sub = "Usage lookback"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Anonymous Access"; value = if ($sharingCapability -match "externalUserAndGuestSharing|anyone") { "Potentially Allowed" } else { "Restricted" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Collection Notes"
        badge = "Important"
        text = "This report shows site-level sharing posture and sharing-link counts from the SharePoint usage report. It does not enumerate individual sharing-link URLs."
    },
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
            @{ key = "ExternalSharing"; header = "External Sharing"; type = "pill" },
            @{ key = "AnonymousLinkCount"; header = "Anyone Links" },
            @{ key = "CompanyLinkCount"; header = "Org Links" },
            @{ key = "GuestSecureLinkCount"; header = "Guest Links" },
            @{ key = "MemberSecureLinkCount"; header = "Member Links" },
            @{ key = "StorageGB"; header = "Storage GB" },
            @{ key = "ActiveFiles"; header = "Active Files" },
            @{ key = "LastActivity"; header = "Last Activity" }
        )
        rows = @(
            $rows |
                Sort-Object `
                    @{ Expression = { $_.AnonymousLinkCount + $_.CompanyLinkCount + $_.GuestSecureLinkCount + $_.MemberSecureLinkCount }; Descending = $true },
                    @{ Expression = 'StorageGB'; Descending = $true }
        )
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
