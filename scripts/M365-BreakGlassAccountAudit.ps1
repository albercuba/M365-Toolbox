[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "Directory.Read.All", "RoleManagement.Read.Directory", "AuditLog.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING BREAK-GLASS ACCOUNT DATA"

$candidateUsers = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,JobTitle,Department,SignInActivity -ErrorAction Stop)
$directoryRoles = @()
try { $directoryRoles = @(Get-MgDirectoryRole -All -ErrorAction Stop) } catch { $directoryRoles = @() }

$privilegedMembers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($role in $directoryRoles) {
    try {
        $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop)
        foreach ($member in $members) {
            if ($member.Id) { [void]$privilegedMembers.Add([string]$member.Id) }
        }
    }
    catch {}
}

$rows = [System.Collections.Generic.List[object]]::new()
foreach ($user in $candidateUsers) {
    $userText = @(
        [string]$user.displayName
        [string]$user.userPrincipalName
        [string]$user.jobTitle
        [string]$user.department
    ) -join " "

    if ($userText -notmatch "(?i)break.?glass|emergency|bgadmin|tier0|tier-0") { continue }

    $methods = @()
    try { $methods = @(Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop) } catch {}
    $methodNames = @($methods | ForEach-Object { Get-GraphAuthMethodLabel -Method $_ } | Where-Object { $_ -and $_ -ne 'Password' })

    $lastSignIn = if ($user.signInActivity -and $user.signInActivity.lastSignInDateTime) { (Get-Date $user.signInActivity.lastSignInDateTime).ToString("yyyy-MM-dd") } else { "Unknown" }
    $mfaState = if ($methodNames.Count -gt 0) { "Registered" } else { "Not Registered" }
    $risk = if (-not $user.accountEnabled) { "Disabled" } elseif ($methodNames.Count -eq 0) { "No MFA" } elseif (-not $privilegedMembers.Contains([string]$user.id)) { "Not Privileged" } else { "Healthy" }

    [void]$rows.Add([pscustomobject]@{
        DisplayName       = [string]$user.displayName
        UserPrincipalName = [string]$user.userPrincipalName
        Enabled           = if ($user.accountEnabled) { "Enabled" } else { "Disabled" }
        Privileged        = if ($privilegedMembers.Contains([string]$user.id)) { "Yes" } else { "No" }
        MfaRegistered     = $mfaState
        Methods           = if ($methodNames.Count -gt 0) { ($methodNames | Sort-Object -Unique) -join ", " } else { "None" }
        LastSignIn        = $lastSignIn
        Risk              = $risk
    })
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "BreakGlassAudit" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Break-Glass Account Audit" -Tenant $tenantName -Subtitle "Emergency access account posture, privilege, and MFA visibility" -Kpis @(
    @{ label = "Accounts"; value = $rows.Count; sub = "Matched break-glass candidates"; cls = "neutral" },
    @{ label = "Privileged"; value = @($rows | Where-Object { $_.Privileged -eq "Yes" }).Count; sub = "Privileged accounts"; cls = "ok" },
    @{ label = "No MFA"; value = @($rows | Where-Object { $_.MfaRegistered -eq "Not Registered" }).Count; sub = "Need attention"; cls = if (@($rows | Where-Object { $_.MfaRegistered -eq "Not Registered" }).Count -gt 0) { "crit" } else { "ok" } },
    @{ label = "Disabled"; value = @($rows | Where-Object { $_.Enabled -eq "Disabled" }).Count; sub = "Disabled candidates"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Break-Glass Accounts"
        badge = "$($rows.Count) accounts"
        columns = @(
            @{ key = "DisplayName"; header = "Account" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "Enabled"; header = "Enabled"; type = "pill" },
            @{ key = "Privileged"; header = "Privileged"; type = "pill" },
            @{ key = "MfaRegistered"; header = "MFA"; type = "pill" },
            @{ key = "Methods"; header = "Methods" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "Risk"; header = "Risk"; type = "pill" }
        )
        rows = @($rows | Sort-Object Risk, DisplayName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
