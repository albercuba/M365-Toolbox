[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$MaxGroupsToInspect = 200
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Group.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING GROUP LIFECYCLE DATA"

$groups = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,createdDateTime,renewedDateTime,visibility&`$top={0}" -f $MaxGroupsToInspect))
$rows = foreach ($group in $groups) {
    $owners = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners?`$select=id" -f $group.id))
    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members?`$select=id" -f $group.id))

    [pscustomobject]@{
        GroupName  = [string]$group.displayName
        Visibility = [string]$group.visibility
        Owners     = $owners.Count
        Members    = $members.Count
        Created    = if ($group.createdDateTime) { (Get-Date $group.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
        Renewed    = if ($group.renewedDateTime) { (Get-Date $group.renewedDateTime).ToString("yyyy-MM-dd") } else { "Never" }
        OwnerState = if ($owners.Count -eq 0) { "Ownerless" } else { "Owned" }
    }
}

$ownerless = @($rows | Where-Object { $_.OwnerState -eq "Ownerless" })
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "GroupLifecycle" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Group Lifecycle Report" -Tenant $tenantName -Subtitle "Unified group ownership and renewal hygiene" -Kpis @(
    @{ label = "Groups"; value = $rows.Count; sub = "Unified groups"; cls = "neutral" },
    @{ label = "Ownerless"; value = $ownerless.Count; sub = "Groups without owners"; cls = if ($ownerless.Count -gt 0) { "crit" } else { "ok" } },
    @{ label = "Private"; value = @($rows | Where-Object { $_.Visibility -eq "Private" }).Count; sub = "Private groups"; cls = "neutral" },
    @{ label = "Inspection Limit"; value = $MaxGroupsToInspect; sub = "Groups queried"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Group Lifecycle"
        badge = "$($rows.Count) groups"
        columns = @(
            @{ key = "GroupName"; header = "Group" },
            @{ key = "Visibility"; header = "Visibility"; type = "pill" },
            @{ key = "Owners"; header = "Owners" },
            @{ key = "Members"; header = "Members" },
            @{ key = "Created"; header = "Created" },
            @{ key = "Renewed"; header = "Renewed" },
            @{ key = "OwnerState"; header = "Owner State"; type = "pill" }
        )
        rows = @($rows | Sort-Object Members -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
