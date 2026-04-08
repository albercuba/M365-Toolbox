[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$StaleDays = 30
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "Directory.Read.All", "AuditLog.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING GUEST INVITATION FAILURES"

$guests = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'&`$select=id,displayName,userPrincipalName,mail,createdDateTime,externalUserState,externalUserStateChangeDateTime,signInActivity&`$top=999")
$staleCutoff = (Get-Date).AddDays(-1 * $StaleDays)

$rows = foreach ($guest in $guests) {
    $lastSignIn = $null
    if ($guest.signInActivity -and $guest.signInActivity.lastSignInDateTime) { $lastSignIn = Get-Date $guest.signInActivity.lastSignInDateTime }
    if ($guest.externalUserState -eq "Accepted" -and $lastSignIn) { continue }

    [pscustomobject]@{
        DisplayName       = [string]$guest.displayName
        UserPrincipalName = [string]$guest.userPrincipalName
        Mail              = [string]$guest.mail
        ExternalState     = if ($guest.externalUserState) { [string]$guest.externalUserState } else { "Unknown" }
        Created           = if ($guest.createdDateTime) { (Get-Date $guest.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
        StateChanged      = if ($guest.externalUserStateChangeDateTime) { (Get-Date $guest.externalUserStateChangeDateTime).ToString("yyyy-MM-dd") } else { "" }
        LastSignIn        = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd") } else { "Never" }
        Stale             = if (-not $lastSignIn -or $lastSignIn -lt $staleCutoff) { "Yes" } else { "No" }
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "GuestInvitationFailures" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Guest Invitation Failures" -Tenant $tenantName -Subtitle "Guests with pending acceptance or no successful sign-in activity" -Kpis @(
    @{ label = "Guests"; value = $rows.Count; sub = "Problem invitations"; cls = if ($rows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Pending"; value = @($rows | Where-Object { $_.ExternalState -ne "Accepted" }).Count; sub = "Not accepted"; cls = "warn" },
    @{ label = "Never Signed In"; value = @($rows | Where-Object { $_.LastSignIn -eq "Never" }).Count; sub = "No activity"; cls = "crit" },
    @{ label = "Stale"; value = @($rows | Where-Object { $_.Stale -eq "Yes" }).Count; sub = "$StaleDays+ days"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Stale Threshold"; value = "$StaleDays days" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Guest Invitation Issues"
        badge = "$($rows.Count) guests"
        columns = @(
            @{ key = "DisplayName"; header = "Guest" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "ExternalState"; header = "Invitation State"; type = "pill" },
            @{ key = "Created"; header = "Created" },
            @{ key = "StateChanged"; header = "State Changed" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "Stale"; header = "Stale"; type = "pill" }
        )
        rows = @($rows | Sort-Object ExternalState, LastSignIn)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
