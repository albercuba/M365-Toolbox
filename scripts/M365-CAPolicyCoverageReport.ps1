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

Write-SectionHeader "COLLECTING CONDITIONAL ACCESS COVERAGE"

$userNameCache = @{}
$groupNameCache = @{}
$appNameCache = @{}

function Expand-CaIds {
    param([object[]]$Ids)

    return @(
        $Ids |
            Where-Object { $_ -ne $null } |
            ForEach-Object { [string]$_ } |
            ForEach-Object { $_ -split ',' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
    )
}

function Get-CaUserNames {
    param([object[]]$Ids)

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($id in @(Expand-CaIds -Ids $Ids)) {
        if (-not $id) { continue }

        switch -Regex ($id) {
            '^All$' {
                [void]$names.Add("All users")
                continue
            }
            '^GuestsOrExternalUsers$' {
                [void]$names.Add("Guests or external users")
                continue
            }
            '^None$' {
                [void]$names.Add("None")
                continue
            }
        }

        if (-not $userNameCache.ContainsKey($id)) {
            try {
                $user = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}?`$select=displayName,userPrincipalName" -f $id) -ErrorAction Stop
                $label = if ($user.userPrincipalName) {
                    "{0} ({1})" -f [string]$user.displayName, [string]$user.userPrincipalName
                }
                else {
                    [string]$user.displayName
                }
                if ($label) {
                    $userNameCache[$id] = $label
                }
                else {
                    $userNameCache[$id] = [string]$id
                }
            }
            catch {
                $userNameCache[$id] = [string]$id
            }
        }

        [void]$names.Add([string]$userNameCache[$id])
    }

    if ($names.Count -gt 0) {
        return ($names -join "`n")
    }

    return "None"
}

function Get-CaGroupNames {
    param([object[]]$Ids)

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($id in @(Expand-CaIds -Ids $Ids)) {
        if (-not $id) { continue }

        if (-not $groupNameCache.ContainsKey($id)) {
            try {
                $group = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/groups/{0}?`$select=displayName" -f $id) -ErrorAction Stop
                if ($group.displayName) {
                    $groupNameCache[$id] = [string]$group.displayName
                }
                else {
                    $groupNameCache[$id] = [string]$id
                }
            }
            catch {
                $groupNameCache[$id] = [string]$id
            }
        }

        [void]$names.Add([string]$groupNameCache[$id])
    }

    if ($names.Count -gt 0) {
        return ($names -join "`n")
    }

    return "None"
}

function Get-CaAppNames {
    param([object[]]$Ids)

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($id in @(Expand-CaIds -Ids $Ids)) {
        if (-not $id) { continue }

        switch -Regex ($id) {
            '^All$' {
                [void]$names.Add("All cloud apps")
                continue
            }
            '^Office365$' {
                [void]$names.Add("Office 365")
                continue
            }
            '^MicrosoftAdminPortals$' {
                [void]$names.Add("Microsoft Admin Portals")
                continue
            }
            '^None$' {
                [void]$names.Add("None")
                continue
            }
        }

        if (-not $appNameCache.ContainsKey($id)) {
            try {
                $servicePrincipals = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '{0}'&`$select=displayName,appId" -f $id))
                $servicePrincipal = @($servicePrincipals | Select-Object -First 1)[0]
                if ($servicePrincipal.displayName) {
                    $appNameCache[$id] = [string]$servicePrincipal.displayName
                }
                else {
                    $appNameCache[$id] = [string]$id
                }
            }
            catch {
                $appNameCache[$id] = [string]$id
            }
        }

        [void]$names.Add([string]$appNameCache[$id])
    }

    if ($names.Count -gt 0) {
        return ($names -join "`n")
    }

    return "None"
}

$policies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
if (-not $IncludeDisabledPolicies) {
    $policies = @($policies | Where-Object { $_.state -ne "disabled" })
}

$rows = foreach ($policy in $policies) {
    $includeUsers = @($policy.conditions.users.includeUsers).Count
    $excludeUsers = @($policy.conditions.users.excludeUsers).Count
    $includeGroups = @($policy.conditions.users.includeGroups).Count
    $excludeGroups = @($policy.conditions.users.excludeGroups).Count
    $includeGuests = @($policy.conditions.users.includeGuestsOrExternalUsers).Count
    $includeApps = @($policy.conditions.applications.includeApplications)
    [pscustomobject]@{
        Policy             = [string]$policy.displayName
        State              = [string]$policy.state
        IncludeUsers       = $includeUsers
        IncludeUsersNames  = Get-CaUserNames -Ids @($policy.conditions.users.includeUsers)
        ExcludeUsers       = $excludeUsers
        ExcludeUsersNames  = Get-CaUserNames -Ids @($policy.conditions.users.excludeUsers)
        IncludeGroups      = $includeGroups
        IncludeGroupsNames = Get-CaGroupNames -Ids @($policy.conditions.users.includeGroups)
        ExcludeGroups      = $excludeGroups
        ExcludeGroupsNames = Get-CaGroupNames -Ids @($policy.conditions.users.excludeGroups)
        GuestScope         = $(if ($includeGuests -gt 0) { "Included" } else { "Not Explicit" })
        Apps               = $includeApps.Count
        AppNames           = Get-CaAppNames -Ids $includeApps
    }
}

$detailSections = foreach ($row in ($rows | Sort-Object Policy)) {
    @{
        title = $row.Policy
        badge = $row.State
        columns = @(
            @{ key = "IncludeUsersNames"; header = "Included Users"; type = "multiline" },
            @{ key = "ExcludeUsersNames"; header = "Excluded Users"; type = "multiline" },
            @{ key = "IncludeGroupsNames"; header = "Included Groups"; type = "multiline" },
            @{ key = "ExcludeGroupsNames"; header = "Excluded Groups"; type = "multiline" },
            @{ key = "AppNames"; header = "Applications"; type = "badgelist" }
        )
        rows = @(
            [pscustomobject]@{
                IncludeUsersNames  = $row.IncludeUsersNames
                ExcludeUsersNames  = $row.ExcludeUsersNames
                IncludeGroupsNames = $row.IncludeGroupsNames
                ExcludeGroupsNames = $row.ExcludeGroupsNames
                AppNames           = $row.AppNames
            }
        )
    }
}

$sections = @(
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
    },
    @{
        title = "Policy Details"
        badge = "$($detailSections.Count) detail cards"
        text = "Detailed policy scope is shown in the cards below."
    }
) + @($detailSections)

$tenantName = $(if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" })
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "CAPolicyCoverage" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 CA Policy Coverage Report" -Tenant $tenantName -Subtitle "Conditional Access scope and exclusions overview" -Kpis @(
    @{ label = "Policies"; value = $rows.Count; sub = "Scoped policies"; cls = "neutral" },
    @{ label = "Enabled"; value = @($rows | Where-Object { $_.State -eq "enabled" }).Count; sub = "Enforced"; cls = "ok" },
    @{ label = "Report Only"; value = @($rows | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }).Count; sub = "Testing"; cls = "warn" },
    @{ label = "With Exclusions"; value = @($rows | Where-Object { $_.ExcludeUsers -gt 0 -or $_.ExcludeGroups -gt 0 }).Count; sub = "Review scope"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Disabled Included"; value = $(if ($IncludeDisabledPolicies) { "Yes" } else { "No" }) },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections $sections

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
