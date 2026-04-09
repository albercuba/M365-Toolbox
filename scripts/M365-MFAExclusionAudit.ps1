[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING MFA EXCLUSION DATA"

$policies = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies")
$rows = foreach ($policy in $policies) {
    $grants = if ($policy.grantControls) { @($policy.grantControls.builtInControls) } else { @() }
    $usesMfa = $grants -contains "mfa" -or [string]$policy.grantControls.authenticationStrength?.id
    if (-not $usesMfa) { continue }

    [pscustomobject]@{
        Policy          = [string]$policy.displayName
        State           = [string]$policy.state
        ExcludedUsers   = @($policy.conditions.users.excludeUsers).Count
        ExcludedGroups  = @($policy.conditions.users.excludeGroups).Count
        ExcludedRoles   = @($policy.conditions.users.excludeRoles).Count
        IncludesGuests  = if (@($policy.conditions.users.includeGuestsOrExternalUsers).Count -gt 0) { "Yes" } else { "No" }
        GrantControl    = if ($grants.Count -gt 0) { $grants -join ", " } else { "Authentication strength" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MFAExclusionAudit" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 MFA Exclusion Audit" -Tenant $tenantName -Subtitle "Conditional Access policies that enforce MFA but still contain exclusions" -Kpis @(
    @{ label = "MFA Policies"; value = $rows.Count; sub = "Policies with MFA or auth strength"; cls = "neutral" },
    @{ label = "User Exclusions"; value = @($rows | Measure-Object -Property ExcludedUsers -Sum).Sum; sub = "Excluded users"; cls = "crit" },
    @{ label = "Group Exclusions"; value = @($rows | Measure-Object -Property ExcludedGroups -Sum).Sum; sub = "Excluded groups"; cls = "warn" },
    @{ label = "Role Exclusions"; value = @($rows | Measure-Object -Property ExcludedRoles -Sum).Sum; sub = "Excluded roles"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Policies With MFA Exclusions"
        badge = "$($rows.Count) policy/policies"
        columns = @(
            @{ key = "Policy"; header = "Policy" },
            @{ key = "State"; header = "State"; type = "pill" },
            @{ key = "ExcludedUsers"; header = "Excluded Users" },
            @{ key = "ExcludedGroups"; header = "Excluded Groups" },
            @{ key = "ExcludedRoles"; header = "Excluded Roles" },
            @{ key = "IncludesGuests"; header = "Guests Included"; type = "pill" },
            @{ key = "GrantControl"; header = "Grant Control" }
        )
        rows = @($rows | Sort-Object Policy)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
