[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("RoleManagement.Read.Directory", "Directory.Read.All", "User.Read.All", "UserAuthenticationMethod.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING ADMIN ROLE DATA"

$privilegedRoleNames = @(
    "Global Administrator",
    "Privileged Role Administrator",
    "Security Administrator",
    "Exchange Administrator",
    "SharePoint Administrator",
    "Teams Administrator",
    "User Administrator",
    "Billing Administrator",
    "Conditional Access Administrator"
)

$directoryRoles = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directoryRoles?$top=200')
$roleRows = [System.Collections.Generic.List[object]]::new()
$adminMap = @{}

foreach ($role in $directoryRoles) {
    if ($role.displayName -notin $privilegedRoleNames) {
        continue
    }

    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/directoryRoles/{0}/members?`$select=id,displayName,userPrincipalName" -f $role.id))
    [void]$roleRows.Add([pscustomobject]@{
        RoleName = [string]$role.displayName
        Members  = $members.Count
    })

    foreach ($member in $members) {
        if (-not $member.id) { continue }
        if (-not $adminMap.ContainsKey([string]$member.id)) {
            $methods = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $member.id))
            $methodNames = @($methods | ForEach-Object { [string]$_.'@odata.type' })
            $hasMfa = @($methodNames | Where-Object { $_ -notmatch "passwordAuthenticationMethod$" }).Count -gt 0

            $adminMap[[string]$member.id] = [pscustomobject]@{
                DisplayName       = [string]$member.displayName
                UserPrincipalName = [string]$member.userPrincipalName
                MfaRegistered     = if ($hasMfa) { "Registered" } else { "Not Registered" }
                Methods           = ($methodNames -replace "^#microsoft.graph\.", "" -join ", ")
            }
        }
    }
}

$admins = @($adminMap.Values | Sort-Object DisplayName)
$adminsWithoutMfa = @($admins | Where-Object { $_.MfaRegistered -eq "Not Registered" })
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "AdminRoleAudit" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Admin Role Audit" -Tenant $tenantName -Subtitle "Privileged roles and MFA hygiene" -Kpis @(
    @{ label = "Admins"; value = $admins.Count; sub = "Unique privileged users"; cls = "neutral" },
    @{ label = "Roles"; value = $roleRows.Count; sub = "Privileged role types"; cls = "neutral" },
    @{ label = "No MFA"; value = $adminsWithoutMfa.Count; sub = "Privileged users without MFA"; cls = if ($adminsWithoutMfa.Count -gt 0) { "crit" } else { "ok" } },
    @{ label = "MFA Coverage"; value = if ($admins.Count -gt 0) { [math]::Round((($admins.Count - $adminsWithoutMfa.Count) / $admins.Count) * 100, 1) } else { 100 }; sub = "Percent"; cls = "ok" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Role Scope"; value = "$($roleRows.Count) privileged roles" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Role Assignment Summary"
        badge = "$($roleRows.Count) roles"
        columns = @(
            @{ key = "RoleName"; header = "Role" },
            @{ key = "Members"; header = "Members" }
        )
        rows = @($roleRows | Sort-Object Members -Descending)
    },
    @{
        title = "Privileged Users"
        badge = "$($admins.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "MfaRegistered"; header = "MFA"; type = "pill" },
            @{ key = "Methods"; header = "Methods" }
        )
        rows = $admins
    },
    @{
        title = "Admins Without MFA"
        badge = "$($adminsWithoutMfa.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "MfaRegistered"; header = "MFA"; type = "pill" },
            @{ key = "Methods"; header = "Methods" }
        )
        rows = $adminsWithoutMfa
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
