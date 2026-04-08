[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("RoleManagement.Read.Directory", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING ELIGIBLE ROLE ASSIGNMENTS"

$items = @()
$warningMessage = $null
try { $items = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$select=id,principalId,roleDefinitionId,startDateTime,endDateTime,memberType&$top=999') } catch { $warningMessage = $_.Exception.Message; Write-Warning "  [!] Eligible role assignments could not be retrieved. $warningMessage" }

$roleCache = @{}
$principalCache = @{}

function Resolve-RoleDefinitionName {
    param([string]$Id)
    if (-not $Id) { return "" }
    if ($roleCache.ContainsKey($Id)) { return $roleCache[$Id] }
    try { $roleCache[$Id] = [string](Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/{0}" -f $Id) -ErrorAction Stop).displayName } catch { $roleCache[$Id] = $Id }
    return $roleCache[$Id]
}

function Resolve-PrincipalName {
    param([string]$Id)
    if (-not $Id) { return "" }
    if ($principalCache.ContainsKey($Id)) { return $principalCache[$Id] }
    try {
        $principal = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/directoryObjects/{0}" -f $Id) -ErrorAction Stop
        $principalCache[$Id] = if ($principal.userPrincipalName) { "{0} ({1})" -f $principal.displayName, $principal.userPrincipalName } else { [string]$principal.displayName }
    }
    catch { $principalCache[$Id] = $Id }
    return $principalCache[$Id]
}

$rows = foreach ($item in $items) {
    [pscustomobject]@{
        Principal  = Resolve-PrincipalName -Id ([string]$item.principalId)
        Role       = Resolve-RoleDefinitionName -Id ([string]$item.roleDefinitionId)
        MemberType = [string]$item.memberType
        Start      = if ($item.startDateTime) { (Get-Date $item.startDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
        End        = if ($item.endDateTime) { (Get-Date $item.endDateTime).ToString("yyyy-MM-dd HH:mm") } else { "Permanent" }
        Permanent  = if ($item.endDateTime) { "No" } else { "Yes" }
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "RoleEligibleAssignments" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Role Eligible Assignments Report" -Tenant $tenantName -Subtitle "Eligible privileged roles and assignment duration posture" -Kpis @(
    @{ label = "Assignments"; value = $rows.Count; sub = "Eligible assignments"; cls = "neutral" },
    @{ label = "Roles"; value = @($rows | Select-Object -ExpandProperty Role -Unique).Count; sub = "Distinct roles"; cls = "neutral" },
    @{ label = "Permanent"; value = @($rows | Where-Object { $_.Permanent -eq "Yes" }).Count; sub = "No end date"; cls = if (@($rows | Where-Object { $_.Permanent -eq "Yes" }).Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Status"; value = if ($warningMessage) { "Warning" } else { "Healthy" }; sub = "Collection state"; cls = if ($warningMessage) { "warn" } else { "ok" } }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Collection Status"
        badge = if ($warningMessage) { "Warning" } else { "Healthy" }
        text = if ($warningMessage) { "Eligible role assignments could not be retrieved: $warningMessage" } else { "Eligible role assignments were collected successfully." }
    },
    @{
        title = "Eligible Assignments"
        badge = "$($rows.Count) assignments"
        columns = @(
            @{ key = "Principal"; header = "Principal" },
            @{ key = "Role"; header = "Role" },
            @{ key = "MemberType"; header = "Member Type"; type = "pill" },
            @{ key = "Start"; header = "Start" },
            @{ key = "End"; header = "End" },
            @{ key = "Permanent"; header = "Permanent"; type = "pill" }
        )
        rows = @($rows | Sort-Object Role, Principal)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
