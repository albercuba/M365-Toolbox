[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$MaxTeamsToInspect = 100
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Group.Read.All", "User.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING TEAMS EXTERNAL ACCESS DATA"

$teams = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,visibility,createdDateTime&`$top={0}" -f $MaxTeamsToInspect))
$teamRows = [System.Collections.Generic.List[object]]::new()
$totalGuestMembers = 0
$teamsWithGuests = 0
$teamsWithoutOwners = 0

foreach ($team in $teams) {
    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,userType" -f $team.id))
    $owners = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners?`$select=id" -f $team.id))
    $guestMembers = @($members | Where-Object { $_.userType -eq "Guest" })

    if ($guestMembers.Count -gt 0) {
        $teamsWithGuests++
        $totalGuestMembers += $guestMembers.Count
    }

    if ($owners.Count -eq 0) {
        $teamsWithoutOwners++
    }

    [void]$teamRows.Add([pscustomobject]@{
        TeamName       = [string]$team.displayName
        Visibility     = [string]$team.visibility
        Owners         = $owners.Count
        Members        = $members.Count
        Guests         = $guestMembers.Count
        ExternalAccess = if ($guestMembers.Count -gt 0) { "Guest Members" } else { "Internal Only" }
        Created        = if ($team.createdDateTime) { (Get-Date $team.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
    })
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "TeamsExternalAccess" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Teams External Access Report" -Tenant $tenantName -Subtitle "Team guest membership exposure and ownership gaps" -Kpis @(
    @{ label = "Teams"; value = $teamRows.Count; sub = "Inspected teams"; cls = "neutral" },
    @{ label = "With Guests"; value = $teamsWithGuests; sub = "Teams with external members"; cls = if ($teamsWithGuests -gt 0) { "warn" } else { "ok" } },
    @{ label = "Guest Members"; value = $totalGuestMembers; sub = "Total guest memberships"; cls = if ($totalGuestMembers -gt 0) { "warn" } else { "ok" } },
    @{ label = "No Owners"; value = $teamsWithoutOwners; sub = "Teams without owners"; cls = if ($teamsWithoutOwners -gt 0) { "crit" } else { "ok" } }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Inspection Limit"; value = $MaxTeamsToInspect },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Teams Exposure"
        badge = "$($teamRows.Count) teams"
        columns = @(
            @{ key = "TeamName"; header = "Team" },
            @{ key = "Visibility"; header = "Visibility"; type = "pill" },
            @{ key = "Owners"; header = "Owners" },
            @{ key = "Members"; header = "Members" },
            @{ key = "Guests"; header = "Guests" },
            @{ key = "ExternalAccess"; header = "Exposure"; type = "pill" },
            @{ key = "Created"; header = "Created" }
        )
        rows = @($teamRows | Sort-Object Guests -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
