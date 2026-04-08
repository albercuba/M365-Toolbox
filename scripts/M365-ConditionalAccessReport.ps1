[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [switch]$IncludeDisabledPolicies
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING CONDITIONAL ACCESS DATA"

$policies = @()
$policyWarning = $null
try {
    $policies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
}
catch {
    $policyWarning = $_.Exception.Message
    Write-Warning "  [!] Conditional Access policies could not be retrieved. $policyWarning"
}

if (-not $IncludeDisabledPolicies) {
    $policies = @($policies | Where-Object { $_.state -ne "disabled" })
}

$policyRows = foreach ($policy in $policies) {
    $includeUsers = @($policy.conditions.users.includeUsers).Count
    $excludeUsers = @($policy.conditions.users.excludeUsers).Count
    $includeApps = @($policy.conditions.applications.includeApplications).Count
    $grantControls = @($policy.grantControls.builtInControls) -join ", "

    [pscustomobject]@{
        Name          = [string]$policy.displayName
        State         = [string]$policy.state
        IncludeUsers  = $includeUsers
        ExcludeUsers  = $excludeUsers
        IncludeApps   = $includeApps
        GrantControls = if ($grantControls) { $grantControls } else { "Not set" }
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$enabledCount = @($policyRows | Where-Object { $_.State -eq "enabled" }).Count
$reportOnlyCount = @($policyRows | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count
$disabledCount = @($policyRows | Where-Object { $_.State -eq "disabled" }).Count
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ConditionalAccess" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Conditional Access Report" -Tenant $tenantName -Subtitle "Policy inventory and enforcement status" -Kpis @(
    @{ label = "Policies"; value = $policyRows.Count; sub = "Scoped policies"; cls = "neutral" },
    @{ label = "Enabled"; value = $enabledCount; sub = "Enforced policies"; cls = "ok" },
    @{ label = "Report Only"; value = $reportOnlyCount; sub = "Testing policies"; cls = if ($reportOnlyCount -gt 0) { "warn" } else { "neutral" } },
    @{ label = "Disabled"; value = $disabledCount; sub = "Inactive policies"; cls = if ($disabledCount -gt 0) { "crit" } else { "ok" } }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Disabled Included"; value = if ($IncludeDisabledPolicies) { "Yes" } else { "No" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Collection Status"
        badge = if ($policyWarning) { "Warning" } else { "Healthy" }
        text = if ($policyWarning) { "Conditional Access policies could not be retrieved: $policyWarning" } else { "Conditional Access policies were collected successfully." }
    },
    @{
        title = "Policy Inventory"
        badge = "$($policyRows.Count) policies"
        columns = @(
            @{ key = "Name"; header = "Policy" },
            @{ key = "State"; header = "State"; type = "pill" },
            @{ key = "IncludeUsers"; header = "Included Users" },
            @{ key = "ExcludeUsers"; header = "Excluded Users" },
            @{ key = "IncludeApps"; header = "Included Apps" },
            @{ key = "GrantControls"; header = "Grant Controls" }
        )
        rows = @($policyRows | Sort-Object Name)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
