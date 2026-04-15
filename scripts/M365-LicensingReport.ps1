[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [switch]$IncludeDisabledUsers,
    [switch]$PaidLicensesOnly
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
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
        "RIGHTSMANAGEMENT"                    = "Azure Information Protection Plan 1"
        "SHAREPOINTSTANDARD"                  = "SharePoint Online Plan 1"
        "SHAREPOINTENTERPRISE"                = "SharePoint Online Plan 2"
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

$userUri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,department,usageLocation,assignedLicenses&$top=999'
$skuUri = 'https://graph.microsoft.com/v1.0/subscribedSkus'

$users = @(Invoke-GraphCollection -Uri $userUri)
$skus = @(Invoke-GraphCollection -Uri $skuUri)

if (-not $IncludeDisabledUsers) {
    $users = @($users | Where-Object { $_.accountEnabled -eq $true })
}

$reportSkus = if ($PaidLicensesOnly) {
    @($skus | Where-Object { Test-IsPaidSku -Sku $_ })
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
$multiLicenseUsers = [System.Collections.Generic.List[object]]::new()
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

    if ($assignedLicenses.Count -gt 1) {
        [void]$multiLicenseUsers.Add([pscustomobject]@{
            DisplayName       = [string]$user.displayName
            UserPrincipalName = [string]$user.userPrincipalName
            LicenseCount      = $assignedLicenses.Count
            Licenses          = ($licenseNames -join ", ")
        })
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "Licensing" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Licensing Report" -Tenant $tenantName -Subtitle "License assignment and SKU consumption" -Kpis @(
    @{ label = "Users"; value = $users.Count; sub = "Scoped users"; cls = "neutral" },
    @{ label = "Licensed"; value = $licensedUsers; sub = if ($PaidLicensesOnly) { "Users with paid licenses" } else { "Users with licenses" }; cls = "ok" },
    @{ label = "Unlicensed"; value = $unlicensedUsers.Count; sub = if ($PaidLicensesOnly) { "Users without paid licenses" } else { "Users without licenses" }; cls = if ($unlicensedUsers.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "SKUs"; value = $skuRows.Count; sub = if ($PaidLicensesOnly) { "Paid subscribed products" } else { "Subscribed products" }; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "User Scope"; value = if ($IncludeDisabledUsers) { "All users" } else { "Enabled only" } },
    @{ label = "License Scope"; value = if ($PaidLicensesOnly) { "Paid only" } else { "All licenses" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = if ($PaidLicensesOnly) { "Paid Subscribed SKUs" } else { "Subscribed SKUs" }
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
        title = if ($PaidLicensesOnly) { "Users With Multiple Paid Licenses" } else { "Users With Multiple Licenses" }
        badge = "$($multiLicenseUsers.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "LicenseCount"; header = "Count" },
            @{ key = "Licenses"; header = "Licenses" }
        )
        rows = @($multiLicenseUsers | Sort-Object LicenseCount -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
