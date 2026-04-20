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

function Get-DirectoryObjectValue {
    param(
        [Parameter(Mandatory)]
        $DirectoryObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $DirectoryObject) {
        return $null
    }

    $property = $DirectoryObject.PSObject.Properties |
        Where-Object { $_.Name -ieq $Name } |
        Select-Object -First 1
    if ($property) {
        return $property.Value
    }

    if ($DirectoryObject -is [System.Collections.IDictionary]) {
        foreach ($key in $DirectoryObject.Keys) {
            if ([string]$key -ieq $Name) {
                return $DirectoryObject[$key]
            }
        }
    }

    $additionalProperties = $DirectoryObject.PSObject.Properties |
        Where-Object { $_.Name -eq 'AdditionalProperties' } |
        Select-Object -First 1
    if ($additionalProperties -and $additionalProperties.Value -is [System.Collections.IDictionary]) {
        foreach ($key in $additionalProperties.Value.Keys) {
            if ([string]$key -ieq $Name) {
                return $additionalProperties.Value[$key]
            }
        }
    }

    return $null
}

function Get-DirectoryObjectLabel {
    param($DirectoryObject)

    if (-not $DirectoryObject) {
        return ""
    }

    $displayName = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'displayName')
    $userPrincipalName = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'userPrincipalName')

    if ($userPrincipalName) {
        return "{0} ({1})" -f $displayName, $userPrincipalName
    }

    $mail = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'mail')

    if ($mail) {
        return "{0} ({1})" -f $displayName, $mail
    }

    if ($displayName) {
        return $displayName
    }

    $id = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'id')
    if ($id) {
        return $id
    }

    return ""
}

$groups = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,createdDateTime,renewedDateTime,visibility&`$top={0}" -f $MaxGroupsToInspect))
$detailRows = [System.Collections.Generic.List[object]]::new()
$rows = foreach ($group in $groups) {
    $owners = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners?`$select=id,displayName,userPrincipalName,mail" -f $group.id))
    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members?`$select=id,displayName,userPrincipalName,mail" -f $group.id))
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
