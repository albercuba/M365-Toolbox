[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$MaxTeamsToInspect = 150
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Group.Read.All", "Directory.Read.All", "User.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING TEAMS OWNERSHIP DATA"

$teams = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,visibility,createdDateTime&`$top={0}" -f $MaxTeamsToInspect))
$rows = foreach ($team in $teams) {
    $owners = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners/microsoft.graph.user?`$select=id,displayName,userPrincipalName" -f $team.id))
    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/microsoft.graph.user?`$select=id,userType" -f $team.id))
    $guestMembers = @($members | Where-Object { $_.userType -eq "Guest" })
    [pscustomobject]@{
        TeamName      = [string]$team.displayName
        Visibility    = [string]$team.visibility
        Owners        = $owners.Count
        Members       = $members.Count
        Guests        = $guestMembers.Count
        OwnerNames    = if ($owners.Count -gt 0) { (@($owners | ForEach-Object { if ($_.userPrincipalName) { "{0} ({1})" -f $_.displayName, $_.userPrincipalName } else { [string]$_.displayName } }) -join ", ") } else { "None" }
        OwnershipRisk = if ($owners.Count -eq 0) { "Ownerless" } elseif ($owners.Count -eq 1) { "Single Owner" } else { "Healthy" }
        Created       = if ($team.createdDateTime) { (Get-Date $team.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "TeamsOwnershipAudit" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Teams Ownership Audit" -Tenant $tenantName -Subtitle "Ownership resilience and guest-heavy teams across the inspected sample" -Kpis @(
    @{ label = "Teams"; value = $rows.Count; sub = "Inspected teams"; cls = "neutral" },
    @{ label = "Ownerless"; value = @($rows | Where-Object { $_.OwnershipRisk -eq "Ownerless" }).Count; sub = "No owners"; cls = "crit" },
    @{ label = "Single Owner"; value = @($rows | Where-Object { $_.OwnershipRisk -eq "Single Owner" }).Count; sub = "Fragile ownership"; cls = "warn" },
    @{ label = "Guest-Heavy"; value = @($rows | Where-Object { $_.Guests -ge 5 }).Count; sub = "5+ guests"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Inspection Limit"; value = $MaxTeamsToInspect },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Teams Ownership"
        badge = "$($rows.Count) teams"
        columns = @(
            @{ key = "TeamName"; header = "Team" },
            @{ key = "Visibility"; header = "Visibility"; type = "pill" },
            @{ key = "Owners"; header = "Owners" },
            @{ key = "Members"; header = "Members" },
            @{ key = "Guests"; header = "Guests" },
            @{ key = "OwnershipRisk"; header = "Ownership Risk"; type = "pill" },
            @{ key = "OwnerNames"; header = "Owner Names" }
        )
        rows = @($rows | Sort-Object OwnershipRisk, Guests -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
