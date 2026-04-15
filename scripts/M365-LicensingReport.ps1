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
    $skuMap[[string]$sku.skuId] = [string]$sku.skuPartNumber
    [pscustomobject]@{
        Sku            = [string]$sku.skuPartNumber
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
            @{ key = "Sku"; header = "SKU" },
            @{ key = "Consumed"; header = "Consumed" },
            @{ key = "Purchased"; header = "Purchased" },
            @{ key = "Available"; header = "Available" },
            @{ key = "Warning"; header = "Warning" },
            @{ key = "Suspended"; header = "Suspended" }
        )
        rows = @($skuRows | Sort-Object Sku)
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
