[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$InactiveDays = 90,
    [switch]$IncludeDisabledUsers
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "AuditLog.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING INACTIVE USER DATA"

$uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,department,createdDateTime,signInActivity,assignedLicenses&$top=999'
$users = @(Invoke-GraphCollection -Uri $uri)
if (-not $IncludeDisabledUsers) {
    $users = @($users | Where-Object { $_.accountEnabled -eq $true })
}

$cutoff = (Get-Date).AddDays(-1 * $InactiveDays)
$rows = foreach ($user in $users) {
    $lastSignIn = $null
    if ($user.signInActivity.lastSignInDateTime) {
        $lastSignIn = Get-Date $user.signInActivity.lastSignInDateTime
    }

    [pscustomobject]@{
        DisplayName       = [string]$user.displayName
        UserPrincipalName = [string]$user.userPrincipalName
        Department        = [string]$user.department
        Enabled           = if ($user.accountEnabled) { "Enabled" } else { "Disabled" }
        LastSignIn        = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd") } else { "Never" }
        Licensed          = if (@($user.assignedLicenses).Count -gt 0) { "Yes" } else { "No" }
        Inactive          = if (-not $lastSignIn -or $lastSignIn -lt $cutoff) { "Yes" } else { "No" }
        Created           = if ($user.createdDateTime) { (Get-Date $user.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
    }
}

$inactiveRows = @($rows | Where-Object { $_.Inactive -eq "Yes" })
$licensedInactiveRows = @($inactiveRows | Where-Object { $_.Licensed -eq "Yes" })
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "InactiveUsers" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Inactive Users Report" -Tenant $tenantName -Subtitle "Inactive identities and potential license reclaim candidates" -Kpis @(
    @{ label = "Users"; value = $rows.Count; sub = "Scoped users"; cls = "neutral" },
    @{ label = "Inactive"; value = $inactiveRows.Count; sub = "No recent sign-in"; cls = if ($inactiveRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Licensed Inactive"; value = $licensedInactiveRows.Count; sub = "Possible reclaim"; cls = if ($licensedInactiveRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Threshold"; value = $InactiveDays; sub = "Days"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Include Disabled"; value = if ($IncludeDisabledUsers) { "Yes" } else { "No" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Inactive Users"
        badge = "$($inactiveRows.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "Department"; header = "Department" },
            @{ key = "Enabled"; header = "Enabled"; type = "pill" },
            @{ key = "Licensed"; header = "Licensed"; type = "pill" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "Created"; header = "Created" }
        )
        rows = @($inactiveRows | Sort-Object LastSignIn, DisplayName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
