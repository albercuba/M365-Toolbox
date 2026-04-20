[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$StaleDays = 90
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "Directory.Read.All", "AuditLog.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING GUEST ACCESS DATA"

$guests = @(Get-MgUser -Filter "userType eq 'Guest'" -All -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,CreatedDateTime,ExternalUserState,SignInActivity -ErrorAction Stop)
$staleCutoff = (Get-Date).AddDays(-1 * $StaleDays)

$domainGroups = $guests |
    ForEach-Object {
        $upn = [string]$_.userPrincipalName
        $domain = if ($upn -match "@") { $upn.Split("@")[-1] } else { "unknown" }
        [pscustomobject]@{ Domain = $domain }
    } |
    Group-Object Domain |
    ForEach-Object {
        [pscustomobject]@{
            Domain = $_.Name
            Guests = $_.Count
        }
    }

$guestRows = foreach ($guest in $guests) {
    $lastSignIn = $null
    if ($guest.signInActivity.lastSignInDateTime) {
        $lastSignIn = Get-Date $guest.signInActivity.lastSignInDateTime
    }

    [pscustomobject]@{
        DisplayName       = [string]$guest.displayName
        UserPrincipalName = [string]$guest.userPrincipalName
        State             = [string]$guest.externalUserState
        Enabled           = if ($guest.accountEnabled) { "Enabled" } else { "Disabled" }
        Created           = if ($guest.createdDateTime) { (Get-Date $guest.createdDateTime).ToString("yyyy-MM-dd") } else { "" }
        LastSignIn        = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd") } else { "Never" }
        Stale             = if (-not $lastSignIn -or $lastSignIn -lt $staleCutoff) { "Yes" } else { "No" }
    }
}

$pendingGuests = @($guestRows | Where-Object { $_.State -match "Pending" })
$staleGuests = @($guestRows | Where-Object { $_.Stale -eq "Yes" })
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "GuestAccess" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Guest Access Report" -Tenant $tenantName -Subtitle "Guest inventory and stale external identities" -Kpis @(
    @{ label = "Guests"; value = $guestRows.Count; sub = "Guest accounts"; cls = "neutral" },
    @{ label = "Pending"; value = $pendingGuests.Count; sub = "Invitation acceptance pending"; cls = if ($pendingGuests.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Stale"; value = $staleGuests.Count; sub = "No recent sign-in"; cls = if ($staleGuests.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Domains"; value = $domainGroups.Count; sub = "External organizations"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Stale Threshold"; value = "$StaleDays days" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Guest Accounts"
        badge = "$($guestRows.Count) guests"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "State"; header = "Invitation State"; type = "pill" },
            @{ key = "Enabled"; header = "Enabled"; type = "pill" },
            @{ key = "Created"; header = "Created" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "Stale"; header = "Stale"; type = "pill" }
        )
        rows = @($guestRows | Sort-Object DisplayName)
    },
    @{
        title = "External Domains"
        badge = "$($domainGroups.Count) domains"
        columns = @(
            @{ key = "Domain"; header = "Domain" },
            @{ key = "Guests"; header = "Guest Count" }
        )
        rows = @($domainGroups | Sort-Object Guests -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
