[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$MaxGroupsToInspect = 200
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Groups")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Group.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING GROUP LIFECYCLE DATA"

$groups = @(Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property Id,DisplayName,CreatedDateTime,RenewedDateTime,Visibility -ErrorAction Stop | Select-Object -First $MaxGroupsToInspect)
$detailRows = [System.Collections.Generic.List[object]]::new()
$rows = foreach ($group in $groups) {
    $owners = @(Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop)
    $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
    $groupName = [string]$group.displayName

    foreach ($owner in $owners) {
        [void]$detailRows.Add([pscustomobject]@{
            Group      = $groupName
            MemberName = Get-DirectoryObjectLabel -DirectoryObject $owner
            Role       = "Owner"
        })
    }

    foreach ($member in $members) {
        [void]$detailRows.Add([pscustomobject]@{
            Group      = $groupName
            MemberName = Get-DirectoryObjectLabel -DirectoryObject $member
            Role       = "Member"
        })
    }

    [pscustomobject]@{
        GroupName    = $groupName
        GroupFilter  = $groupName
        Visibility   = [string]$group.visibility
        Owners       = $owners.Count
        Members      = $members.Count
        Created      = if ($group.createdDateTime) { (Get-Date $group.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
        Renewed      = if ($group.renewedDateTime) { (Get-Date $group.renewedDateTime).ToString("yyyy-MM-dd") } else { "Never" }
        OwnerState   = if ($owners.Count -eq 0) { "Ownerless" } else { "Owned" }
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
        id = "group-lifecycle-summary"
        title = "Group Lifecycle"
        badge = "$($rows.Count) groups"
        rowAction = @{
            targetSectionId = "group-lifecycle-details"
            sourceKey = "GroupFilter"
            ariaLabel = "Open group lifecycle details for this group"
        }
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
    },
    @{
        id = "group-lifecycle-details"
        title = "Group Lifecycle Details"
        badge = "$($detailRows.Count) assignment(s)"
        filterKey = "Group"
        filterLabel = "Group Filter"
        filterPlaceholder = "Type a Group name"
        columns = @(
            @{ key = "Group"; header = "Group" },
            @{ key = "MemberName"; header = "Member Name" },
            @{ key = "Role"; header = "Role"; type = "pill" }
        )
        rows = @($detailRows | Sort-Object Group, Role, MemberName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
