[CmdletBinding()]
param(
    [string]$TenantId,
    [int]$LookbackDays = 30,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns", "Microsoft.Graph.DirectoryObjects")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("RoleManagement.Read.Directory", "AuditLog.Read.All", "Directory.Read.All", "User.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING PRIVILEGED USER SIGN-INS"

$roles = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/directoryRoles")
$privilegedRows = [System.Collections.Generic.List[object]]::new()
$membersByUser = @{}

foreach ($role in $roles) {
    $members = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/directoryRoles/{0}/members" -f $role.id))
    foreach ($member in $members) {
        $upn = [string]$member.userPrincipalName
        if (-not $upn) { continue }
        if (-not $membersByUser.ContainsKey($upn)) {
            $membersByUser[$upn] = [System.Collections.Generic.List[string]]::new()
        }
        [void]$membersByUser[$upn].Add([string]$role.displayName)
    }
}

$startUtc = (Get-Date).AddDays(-1 * $LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$signIns = @(Get-MgAuditLogSignIn -Filter "createdDateTime ge $startUtc" -All -ErrorAction Stop)

foreach ($upn in $membersByUser.Keys) {
    $userSignIns = @($signIns | Where-Object { $_.UserPrincipalName -eq $upn } | Sort-Object CreatedDateTime -Descending)
    $latest = $userSignIns | Select-Object -First 1
    [void]$privilegedRows.Add([pscustomobject]@{
        UserPrincipalName = $upn
        Roles             = (($membersByUser[$upn] | Select-Object -Unique) -join ", ")
        LastSignIn        = if ($latest) { [string]$latest.CreatedDateTime } else { "No recent sign-in" }
        LastIP            = if ($latest) { [string]$latest.IpAddress } else { "" }
        LastLocation      = if ($latest) { Get-GeoLocationString -Location $latest.Location } else { "" }
        LastClient        = if ($latest) { [string]$latest.ClientAppUsed } else { "" }
        SignInCount       = $userSignIns.Count
    })
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "PrivilegedUserSignInReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Privileged User Sign-In Review" -Tenant $tenantName -Subtitle "Recent sign-in visibility for privileged identities" -Kpis @(
    @{ label = "Privileged Users"; value = $privilegedRows.Count; sub = "Users in active directory roles"; cls = "neutral" },
    @{ label = "Roles"; value = $roles.Count; sub = "Directory roles reviewed"; cls = "neutral" },
    @{ label = "No Recent Sign-In"; value = @($privilegedRows | Where-Object { $_.LastSignIn -eq 'No recent sign-in' }).Count; sub = "Privileged users with no recent activity"; cls = "warn" },
    @{ label = "Lookback"; value = "$LookbackDays d"; sub = "Sign-in review window"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Privileged User Sign-Ins"
        badge = "$($privilegedRows.Count) user(s)"
        columns = @(
            @{ key = "UserPrincipalName"; header = "User" },
            @{ key = "Roles"; header = "Roles" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "LastIP"; header = "IP" },
            @{ key = "LastLocation"; header = "Location" },
            @{ key = "LastClient"; header = "Client" },
            @{ key = "SignInCount"; header = "Sign-Ins" }
        )
        rows = @($privilegedRows | Sort-Object UserPrincipalName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
