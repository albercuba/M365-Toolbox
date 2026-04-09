[CmdletBinding()]
param(
    [string]$TenantId,
    [int]$DormantDays = 45,
    [int]$LookbackDays = 90,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns", "Microsoft.Graph.DirectoryObjects")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("RoleManagement.Read.Directory", "AuditLog.Read.All", "Directory.Read.All", "User.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING DORMANT ADMIN DATA"

$roles = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/directoryRoles")
$membersByUser = @{}
foreach ($role in $roles) {
    foreach ($member in @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/directoryRoles/{0}/members" -f $role.id))) {
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
$thresholdDate = (Get-Date).AddDays(-1 * $DormantDays)
$rows = [System.Collections.Generic.List[object]]::new()

foreach ($upn in $membersByUser.Keys) {
    $latest = @($signIns | Where-Object { $_.UserPrincipalName -eq $upn } | Sort-Object CreatedDateTime -Descending | Select-Object -First 1)
    $lastSeen = if ($latest) { Get-Date $latest[0].CreatedDateTime } else { $null }
    $status = if (-not $lastSeen) { "No sign-in seen" } elseif ($lastSeen -lt $thresholdDate) { "Dormant" } else { "Active" }
    [void]$rows.Add([pscustomobject]@{
        UserPrincipalName = $upn
        Roles             = (($membersByUser[$upn] | Select-Object -Unique) -join ", ")
        LastSignIn        = if ($lastSeen) { $lastSeen.ToString("yyyy-MM-dd HH:mm") } else { "No recent sign-in" }
        DaysSinceSignIn   = if ($lastSeen) { [int]((New-TimeSpan -Start $lastSeen -End (Get-Date)).TotalDays) } else { ">$LookbackDays" }
        Status            = $status
    })
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "DormantAdminAccountReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Dormant Admin Account Review" -Tenant $tenantName -Subtitle "Privileged identities with little or no recent sign-in activity" -Kpis @(
    @{ label = "Privileged Users"; value = $rows.Count; sub = "Users in active admin roles"; cls = "neutral" },
    @{ label = "Dormant"; value = @($rows | Where-Object { $_.Status -eq 'Dormant' }).Count; sub = "Sign-in older than threshold"; cls = "warn" },
    @{ label = "No Sign-In"; value = @($rows | Where-Object { $_.Status -eq 'No sign-in seen' }).Count; sub = "No sign-in in lookback"; cls = "crit" },
    @{ label = "Threshold"; value = "$DormantDays d"; sub = "Dormant threshold"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Lookback"; value = "$LookbackDays day(s)" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Dormant Privileged Accounts"
        badge = "$($rows.Count) account(s)"
        columns = @(
            @{ key = "UserPrincipalName"; header = "User" },
            @{ key = "Roles"; header = "Roles" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "DaysSinceSignIn"; header = "Days Since" },
            @{ key = "Status"; header = "Status"; type = "pill" }
        )
        rows = @($rows | Sort-Object Status, UserPrincipalName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
