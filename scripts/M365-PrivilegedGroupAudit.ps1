[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Group.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING PRIVILEGED GROUP DATA"

$groupNames = @("Company Administrator","Privileged Role Administrator","Helpdesk Administrator","User Administrator","Exchange Administrator","SharePoint Administrator","Teams Administrator")
$groups = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mailEnabled,securityEnabled,visibility&$top=999')
$rows = foreach ($group in $groups | Where-Object { $_.displayName -in $groupNames }) {
    $owners = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners?`$select=id" -f $group.id))
    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members?`$select=id" -f $group.id))
    [pscustomobject]@{
        Group      = [string]$group.displayName
        Visibility = [string]$group.visibility
        Owners     = $owners.Count
        Members    = $members.Count
        OwnerState = if ($owners.Count -eq 0) { "Ownerless" } else { "Owned" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "PrivilegedGroups" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Privileged Group Audit" -Tenant $tenantName -Subtitle "Sensitive groups, owners, and membership exposure" -Kpis @(
    @{ label = "Groups"; value = $rows.Count; sub = "Matched privileged groups"; cls = "neutral" },
    @{ label = "Ownerless"; value = @($rows | Where-Object { $_.OwnerState -eq "Ownerless" }).Count; sub = "Needs review"; cls = "crit" },
    @{ label = "Members"; value = (($rows | Measure-Object Members -Sum).Sum); sub = "Total memberships"; cls = "warn" },
    @{ label = "Visible"; value = @($rows | Where-Object { $_.Visibility -ne "HiddenMembership" }).Count; sub = "Non-hidden groups"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Privileged Groups"
        badge = "$($rows.Count) groups"
        columns = @(
            @{ key = "Group"; header = "Group" },
            @{ key = "Visibility"; header = "Visibility"; type = "pill" },
            @{ key = "Owners"; header = "Owners" },
            @{ key = "Members"; header = "Members" },
            @{ key = "OwnerState"; header = "Owner State"; type = "pill" }
        )
        rows = @($rows | Sort-Object Members -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
