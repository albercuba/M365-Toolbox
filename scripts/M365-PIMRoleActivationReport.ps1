[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.Governance")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("RoleManagement.Read.Directory", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING PIM ROLE DATA"

$eligibilities = @(Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ErrorAction Stop)
$activations = @(Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All -ErrorAction Stop)

$rows = foreach ($item in $activations) {
    [pscustomobject]@{
        PrincipalId = [string]$item.principalId
        RoleId      = [string]$item.roleDefinitionId
        Assignment  = "Active"
        Start       = if ($item.startDateTime) { (Get-Date $item.startDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
        End         = if ($item.endDateTime) { (Get-Date $item.endDateTime).ToString("yyyy-MM-dd HH:mm") } else { "Permanent" }
    }
}

$eligibilityRows = foreach ($item in $eligibilities) {
    [pscustomobject]@{
        PrincipalId = [string]$item.principalId
        RoleId      = [string]$item.roleDefinitionId
        Assignment  = "Eligible"
        Start       = if ($item.startDateTime) { (Get-Date $item.startDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
        End         = if ($item.endDateTime) { (Get-Date $item.endDateTime).ToString("yyyy-MM-dd HH:mm") } else { "Permanent" }
    }
}

$allRows = @($rows + $eligibilityRows)
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "PIMRoleActivation" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 PIM Role Activation Report" -Tenant $tenantName -Subtitle "Active and eligible directory role schedules" -Kpis @(
    @{ label = "Active"; value = $rows.Count; sub = "Active schedules"; cls = "warn" },
    @{ label = "Eligible"; value = $eligibilityRows.Count; sub = "Eligible schedules"; cls = "ok" },
    @{ label = "Permanent"; value = @($allRows | Where-Object { $_.End -eq "Permanent" }).Count; sub = "No end time"; cls = "warn" },
    @{ label = "Total"; value = $allRows.Count; sub = "Schedule instances"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Role Schedules"
        badge = "$($allRows.Count) entries"
        columns = @(
            @{ key = "PrincipalId"; header = "Principal ID" },
            @{ key = "RoleId"; header = "Role ID" },
            @{ key = "Assignment"; header = "Assignment"; type = "pill" },
            @{ key = "Start"; header = "Start" },
            @{ key = "End"; header = "End" }
        )
        rows = @($allRows | Sort-Object Assignment, End)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
