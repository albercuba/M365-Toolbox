[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [switch]$IncludeDisabledPolicies
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING CONDITIONAL ACCESS COVERAGE"

$policies = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999')
if (-not $IncludeDisabledPolicies) {
    $policies = @($policies | Where-Object { $_.state -ne "disabled" })
}

$rows = foreach ($policy in $policies) {
    $includeUsers = @($policy.conditions.users.includeUsers).Count
    $excludeUsers = @($policy.conditions.users.excludeUsers).Count
    $includeGroups = @($policy.conditions.users.includeGroups).Count
    $excludeGroups = @($policy.conditions.users.excludeGroups).Count
    $includeGuests = @($policy.conditions.users.includeGuestsOrExternalUsers).Count
    [pscustomobject]@{
        Policy          = [string]$policy.displayName
        State           = [string]$policy.state
        IncludeUsers    = $includeUsers
        ExcludeUsers    = $excludeUsers
        IncludeGroups   = $includeGroups
        ExcludeGroups   = $excludeGroups
        GuestScope      = if ($includeGuests -gt 0) { "Included" } else { "Not Explicit" }
        Apps            = @($policy.conditions.applications.includeApplications).Count
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "CAPolicyCoverage" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 CA Policy Coverage Report" -Tenant $tenantName -Subtitle "Conditional Access scope and exclusions overview" -Kpis @(
    @{ label = "Policies"; value = $rows.Count; sub = "Scoped policies"; cls = "neutral" },
    @{ label = "Enabled"; value = @($rows | Where-Object { $_.State -eq "enabled" }).Count; sub = "Enforced"; cls = "ok" },
    @{ label = "Report Only"; value = @($rows | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count; sub = "Testing"; cls = "warn" },
    @{ label = "With Exclusions"; value = @($rows | Where-Object { $_.ExcludeUsers -gt 0 -or $_.ExcludeGroups -gt 0 }).Count; sub = "Review scope"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Disabled Included"; value = if ($IncludeDisabledPolicies) { "Yes" } else { "No" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Policy Coverage"
        badge = "$($rows.Count) policies"
        columns = @(
            @{ key = "Policy"; header = "Policy" },
            @{ key = "State"; header = "State"; type = "pill" },
            @{ key = "IncludeUsers"; header = "Inc Users" },
            @{ key = "ExcludeUsers"; header = "Exc Users" },
            @{ key = "IncludeGroups"; header = "Inc Groups" },
            @{ key = "ExcludeGroups"; header = "Exc Groups" },
            @{ key = "GuestScope"; header = "Guests"; type = "pill" },
            @{ key = "Apps"; header = "Apps" }
        )
        rows = @($rows | Sort-Object Policy)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
