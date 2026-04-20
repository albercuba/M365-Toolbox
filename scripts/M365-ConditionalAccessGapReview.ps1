[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING CONDITIONAL ACCESS POLICY GAPS"

$policies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
$rows = foreach ($policy in $policies) {
    $signals = [System.Collections.Generic.List[string]]::new()
    if ([string]$policy.state -ne 'enabled') { [void]$signals.Add("Policy state is $($policy.state)") }
    if (@($policy.conditions.users.excludeUsers).Count -gt 0) { [void]$signals.Add("Excluded users present") }
    if (@($policy.conditions.users.excludeGroups).Count -gt 0) { [void]$signals.Add("Excluded groups present") }
    if (-not $policy.grantControls -or @($policy.grantControls.builtInControls).Count -eq 0) { [void]$signals.Add("No grant controls defined") }

    [pscustomobject]@{
        Name           = [string]$policy.displayName
        State          = [string]$policy.state
        IncludedUsers  = @($policy.conditions.users.includeUsers).Count
        ExcludedUsers  = @($policy.conditions.users.excludeUsers).Count
        ExcludedGroups = @($policy.conditions.users.excludeGroups).Count
        Grants         = if ($policy.grantControls) { (@($policy.grantControls.builtInControls) -join ", ") } else { "None" }
        GapSignals     = if ($signals.Count -gt 0) { $signals -join "; " } else { "No obvious gap signal" }
    }
}

$gapRows = @($rows | Where-Object { $_.GapSignals -ne 'No obvious gap signal' })
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ConditionalAccessGapReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Conditional Access Gap Review" -Tenant $tenantName -Subtitle "Conditional Access coverage issues, exclusions, and non-enforced states" -Kpis @(
    @{ label = "Policies"; value = $rows.Count; sub = "Conditional Access policies"; cls = "neutral" },
    @{ label = "Gap Signals"; value = $gapRows.Count; sub = "Policies needing review"; cls = if ($gapRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Disabled/Report"; value = @($rows | Where-Object { $_.State -ne 'enabled' }).Count; sub = "Non-enforced policies"; cls = "warn" },
    @{ label = "User Exclusions"; value = @($rows | Measure-Object -Property ExcludedUsers -Sum).Sum; sub = "Excluded users across policies"; cls = "crit" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Conditional Access Policy Gaps"
        badge = "$($rows.Count) policy/policies"
        columns = @(
            @{ key = "Name"; header = "Policy" },
            @{ key = "State"; header = "State"; type = "pill" },
            @{ key = "IncludedUsers"; header = "Included Users" },
            @{ key = "ExcludedUsers"; header = "Excluded Users" },
            @{ key = "ExcludedGroups"; header = "Excluded Groups" },
            @{ key = "Grants"; header = "Grant Controls" },
            @{ key = "GapSignals"; header = "Gap Signals" }
        )
        rows = @($rows | Sort-Object Name)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
