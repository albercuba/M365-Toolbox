[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [switch]$IncludeDisabledUsers,
    [switch]$PaidLicensesOnly,
    [switch]$IncludePaidAddOns
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "Directory.Read.All", "Organization.Read.All")
Resolve-ToolboxTenantLabel

function Test-IsPaidSku {
    param(
        [Parameter(Mandatory)]
        $Sku
    )

    $skuPartNumber = [string]$Sku.skuPartNumber
    if (-not $skuPartNumber) {
        return $false
    }

    $freeOrTrialPatterns = @(
        'FREE',
        'TRIAL',
        'VIRAL',
        'EXPLORATORY',
        'DEVELOPER'
    )

    foreach ($pattern in $freeOrTrialPatterns) {
        if ($skuPartNumber -match $pattern) {
            return $false
        }
    }

    return $true
}

function Test-IsDeniedPaidSku {
    param(
        [Parameter(Mandatory)]
        $Sku
    )

    $skuPartNumber = [string]$Sku.skuPartNumber
    if (-not $skuPartNumber) {
        return $true
    }

    $deniedSkuPartNumbers = @(
        'FLOW_FREE',
        'POWER_BI_STANDARD',
        'TEAMS_EXPLORATORY',
        'WINDOWS_STORE',
        'PHONESYSTEM_VIRTUALUSER'
    )

    return $skuPartNumber -in $deniedSkuPartNumbers
}

function Test-IsPaidAddOnSku {
    param(
        [Parameter(Mandatory)]
        $Sku
    )

    $skuPartNumber = [string]$Sku.skuPartNumber
    if (-not $skuPartNumber) {
        return $false
    }

    $addOnSkuPartNumbers = @(
        'AAD_PREMIUM',
        'AAD_PREMIUM_P2',
        'ATP_ENTERPRISE',
        'ATP_ENTERPRISE_FACULTY',
        'ATP_ENTERPRISE_GOV',
        'EXCHANGESTANDARD',
        'EXCHANGEENTERPRISE',
        'INTUNE_A',
        'MCOEV',
        'MCOMEETADV',
        'POWER_BI_PRO',
        'PROJECTESSENTIALS',
        'PROJECTPREMIUM',
        'PROJECTPROFESSIONAL',
        'RIGHTSMANAGEMENT',
        'SHAREPOINTSTORAGE',
        'VISIO_PLAN1_DEPT',
        'VISIO_PLAN2_DEPT',
        'WACONEDRIVESTANDARD',
        'WIN_DEF_ATP'
    )

    return $skuPartNumber -in $addOnSkuPartNumbers
}

function Get-LicenseScopeLabel {
    param(
        [switch]$PaidOnly,
        [switch]$IncludeAddOns
    )

    if (-not $PaidOnly) {
        return "All licenses"
    }

    if ($IncludeAddOns) {
        return "All paid SKUs including add-ons"
    }

    return "Paid user licenses only"
}

function Get-FriendlySkuName {
    param(
        [string]$SkuPartNumber
    )

    if (-not $SkuPartNumber) {
        return "Unknown SKU"
    }

    $friendlyNames = @{
        "AAD_PREMIUM"                         = "Microsoft Entra ID P1"
        "AAD_PREMIUM_P2"                      = "Microsoft Entra ID P2"
        "ATP_ENTERPRISE"                      = "Microsoft Defender for Office 365 Plan 1"
        "ATP_ENTERPRISE_FACULTY"              = "Microsoft Defender for Office 365 Plan 1 for Faculty"
        "ATP_ENTERPRISE_GOV"                  = "Microsoft Defender for Office 365 Plan 1 for Government"
        "BUSINESS_BASIC"                      = "Microsoft 365 Business Basic"
        "BUSINESS_PREMIUM"                    = "Microsoft 365 Business Premium"
        "DESKLESSPACK"                        = "Office 365 F3"
        "EMS"                                 = "Enterprise Mobility + Security E3"
        "EMSPREMIUM"                          = "Enterprise Mobility + Security E5"
        "ENTERPRISEPACK"                      = "Office 365 E3"
        "ENTERPRISEPREMIUM"                   = "Office 365 E5"
        "EXCHANGESTANDARD"                    = "Exchange Online Plan 1"
        "EXCHANGEENTERPRISE"                  = "Exchange Online Plan 2"
        "FLOW_FREE"                           = "Power Automate Free"
        "INTUNE_A"                            = "Microsoft Intune Plan A"
        "MCOEV"                               = "Microsoft Teams Phone Standard"
        "MCOMEETADV"                          = "Microsoft 365 Audio Conferencing"
        "MCOCAP"                              = "Microsoft Teams Shared Devices"
        "MCOIMP"                              = "Microsoft Teams Phone Resource Account"
        "MCOProfessional"                     = "Skype for Business Online Plan 2"
        "O365_BUSINESS_ESSENTIALS"            = "Office 365 Business Essentials"
        "O365_BUSINESS_PREMIUM"               = "Office 365 Business Premium"
        "POWER_BI_PRO"                        = "Power BI Pro"
        "POWER_BI_STANDARD"                   = "Power BI Free"
        "PROJECTESSENTIALS"                   = "Project Plan 1"
        "PROJECTPREMIUM"                      = "Project Plan 5"
        "PROJECTPROFESSIONAL"                 = "Project Plan 3"
        "PHONESYSTEM_VIRTUALUSER"             = "Microsoft Teams Phone Resource Account"
        "RIGHTSMANAGEMENT"                    = "Azure Information Protection Plan 1"
        "SHAREPOINTSTANDARD"                  = "SharePoint Online Plan 1"
        "SHAREPOINTENTERPRISE"                = "SharePoint Online Plan 2"
        "SHAREPOINTSTORAGE"                   = "SharePoint Online Storage"
        "SPE_E3"                              = "Microsoft 365 E3"
        "SPE_E5"                              = "Microsoft 365 E5"
        "SPB"                                 = "Microsoft 365 Business Premium"
        "STANDARDPACK"                        = "Office 365 E1"
        "STREAM"                              = "Microsoft Stream"
        "TEAMS_EXPLORATORY"                   = "Microsoft Teams Exploratory"
        "VISIO_PLAN1_DEPT"                    = "Visio Plan 1"
        "VISIO_PLAN2_DEPT"                    = "Visio Plan 2"
        "WACONEDRIVESTANDARD"                 = "OneDrive for Business Plan 1"
        "WIN_DEF_ATP"                         = "Microsoft Defender for Endpoint Plan 2"
        "WINDOWS_STORE"                       = "Windows Store for Business"
        "M365EDU_A1"                          = "Microsoft 365 A1"
        "M365EDU_A3_FACULTY"                  = "Microsoft 365 A3 for Faculty"
        "M365EDU_A3_STUDENT"                  = "Microsoft 365 A3 for Students"
        "M365EDU_A5_FACULTY"                  = "Microsoft 365 A5 for Faculty"
        "M365EDU_A5_STUDENT"                  = "Microsoft 365 A5 for Students"
    }

    if ($friendlyNames.ContainsKey($SkuPartNumber)) {
        return $friendlyNames[$SkuPartNumber]
    }

    return $SkuPartNumber
}

Write-SectionHeader "COLLECTING LICENSING DATA"

$users = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,Department,UsageLocation,AssignedLicenses -ErrorAction Stop)
$skus = @(Get-MgSubscribedSku -All -Property SkuId,SkuPartNumber,ConsumedUnits,PrepaidUnits -ErrorAction Stop)

if (-not $IncludeDisabledUsers) {
    $users = @($users | Where-Object { $_.accountEnabled -eq $true })
}

$reportSkus = if ($PaidLicensesOnly) {
    @(
        $skus |
            Where-Object { (Test-IsPaidSku -Sku $_) -and -not (Test-IsDeniedPaidSku -Sku $_) } |
            Where-Object {
                if ($IncludePaidAddOns) {
                    $true
                }
                else {
                    -not (Test-IsPaidAddOnSku -Sku $_)
                }
            }
    )
}
else {
    @($skus)
}

$skuMap = @{}
$skuRows = foreach ($sku in $reportSkus) {
    $friendlySkuName = Get-FriendlySkuName -SkuPartNumber ([string]$sku.skuPartNumber)
    $skuMap[[string]$sku.skuId] = $friendlySkuName
    [pscustomobject]@{
        Product        = $friendlySkuName
        SkuCode        = [string]$sku.skuPartNumber
        Consumed       = [int]$sku.consumedUnits
        Purchased      = [int]($sku.prepaidUnits.enabled)
        Available      = [int](($sku.prepaidUnits.enabled) - $sku.consumedUnits)
        Warning        = [int]($sku.prepaidUnits.warning)
        Suspended      = [int]($sku.prepaidUnits.suspended)
    }
}

$unlicensedUsers = [System.Collections.Generic.List[object]]::new()
$licenseDetailUsers = [System.Collections.Generic.List[object]]::new()
$licensedUsers = 0

foreach ($user in $users) {
    $assignedLicenses = @($user.assignedLicenses | Where-Object { $skuMap.ContainsKey([string]$_.skuId) })
    $licenseNames = @($assignedLicenses | ForEach-Object { $skuMap[[string]$_.skuId] } | Where-Object { $_ })

    if ($assignedLicenses.Count -gt 0) {
        $licensedUsers++
    }
    else {
        [void]$unlicensedUsers.Add([pscustomobject]@{
            DisplayName       = [string]$user.displayName
            UserPrincipalName = [string]$user.userPrincipalName
            Department        = [string]$user.department
            UsageLocation     = [string]$user.usageLocation
            Enabled           = if ($user.accountEnabled) { "Yes" } else { "No" }
        })
    }

    $shouldIncludeInLicenseSection = if ($PaidLicensesOnly) {
        $assignedLicenses.Count -gt 0
    }
    else {
        $assignedLicenses.Count -gt 1
    }

    if ($shouldIncludeInLicenseSection) {
        [void]$licenseDetailUsers.Add([pscustomobject]@{
            DisplayName       = [string]$user.displayName
            UserPrincipalName = [string]$user.userPrincipalName
            LicenseCount      = $assignedLicenses.Count
            Licenses          = ($licenseNames -join ", ")
        })
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$licenseScopeLabel = Get-LicenseScopeLabel -PaidOnly:$PaidLicensesOnly -IncludeAddOns:$IncludePaidAddOns
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "Licensing" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Licensing Report" -Tenant $tenantName -Subtitle "License assignment and SKU consumption" -Kpis @(
    @{ label = "Users"; value = $users.Count; sub = "Scoped users"; cls = "neutral" },
    @{ label = "Licensed"; value = $licensedUsers; sub = if ($PaidLicensesOnly) { "Users with licenses in scope" } else { "Users with licenses" }; cls = "ok" },
    @{ label = "Unlicensed"; value = $unlicensedUsers.Count; sub = if ($PaidLicensesOnly) { "Users without licenses in scope" } else { "Users without licenses" }; cls = if ($unlicensedUsers.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "SKUs"; value = $skuRows.Count; sub = if ($PaidLicensesOnly) { "Paid products in scope" } else { "Subscribed products" }; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "User Scope"; value = if ($IncludeDisabledUsers) { "All users" } else { "Enabled only" } },
    @{ label = "License Scope"; value = $licenseScopeLabel },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = if ($PaidLicensesOnly) { "Subscribed SKUs In Scope" } else { "Subscribed SKUs" }
        badge = "$($skuRows.Count) products"
        columns = @(
            @{ key = "Product"; header = "Product" },
            @{ key = "SkuCode"; header = "SKU Code" },
            @{ key = "Consumed"; header = "Consumed" },
            @{ key = "Purchased"; header = "Purchased" },
            @{ key = "Available"; header = "Available" },
            @{ key = "Warning"; header = "Warning" },
            @{ key = "Suspended"; header = "Suspended" }
        )
        rows = @($skuRows | Sort-Object Product)
    },
    @{
        title = "Unlicensed Users"
        badge = "$($unlicensedUsers.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "Department"; header = "Department" },
            @{ key = "UsageLocation"; header = "Usage Location" },
            @{ key = "Enabled"; header = "Enabled"; type = "pill" }
        )
        rows = @($unlicensedUsers | Sort-Object DisplayName)
    },
    @{
        title = if ($PaidLicensesOnly) { "Users With Licenses In Scope" } else { "Users With Multiple Licenses" }
        badge = "$($licenseDetailUsers.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "LicenseCount"; header = "Count" },
            @{ key = "Licenses"; header = "Licenses" }
        )
        rows = @(
            $licenseDetailUsers |
                Sort-Object `
                    @{ Expression = "LicenseCount"; Descending = $true },
                    @{ Expression = "DisplayName"; Descending = $false }
        )
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
