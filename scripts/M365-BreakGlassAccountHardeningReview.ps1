[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$AccountNameHints = "breakglass,emergency",
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "AuditLog.Read.All", "RoleManagement.Read.Directory", "UserAuthenticationMethod.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING BREAK-GLASS ACCOUNT DATA"

$terms = Normalize-DelimitedValue -Value @($AccountNameHints)
$users = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,UserType)
$matchingUsers = @(
    $users | Where-Object {
        $label = ("{0} {1}" -f $_.DisplayName, $_.UserPrincipalName).ToLowerInvariant()
        @($terms | Where-Object { $label.Contains($_.ToLowerInvariant()) }).Count -gt 0
    }
)

$signIns = @(Get-MgAuditLogSignIn -Filter ("createdDateTime ge {0}" -f (Get-Date).AddDays(-90).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")) -All -ErrorAction Stop)
$rows = foreach ($user in $matchingUsers) {
    $methods = @(Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction SilentlyContinue | Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.passwordAuthenticationMethod' })
    $latest = @($signIns | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName } | Sort-Object CreatedDateTime -Descending | Select-Object -First 1)
    [pscustomobject]@{
        UserPrincipalName = [string]$user.UserPrincipalName
        DisplayName       = [string]$user.DisplayName
        AccountEnabled    = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
        MFARegistered     = if ($methods.Count -gt 0) { "Registered" } else { "No MFA" }
        LastSignIn        = if ($latest) { [string]$latest[0].CreatedDateTime } else { "No recent sign-in" }
        HardeningNotes    = if ($methods.Count -gt 0 -and $user.AccountEnabled) { "Review CA exclusions and emergency access documentation." } else { "Needs hardening review." }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "BreakGlassAccountHardeningReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Break-Glass Account Hardening Review" -Tenant $tenantName -Subtitle "Likely emergency access accounts and their hardening posture" -Kpis @(
    @{ label = "Candidates"; value = $rows.Count; sub = "Accounts matching break-glass hints"; cls = "neutral" },
    @{ label = "No MFA"; value = @($rows | Where-Object { $_.MFARegistered -eq 'No MFA' }).Count; sub = "Emergency accounts without MFA"; cls = "crit" },
    @{ label = "Enabled"; value = @($rows | Where-Object { $_.AccountEnabled -eq 'Enabled' }).Count; sub = "Enabled emergency accounts"; cls = "warn" },
    @{ label = "Hints"; value = ($terms -join ", "); sub = "Search terms"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Break-Glass Hardening Review"
        badge = "$($rows.Count) account(s)"
        columns = @(
            @{ key = "UserPrincipalName"; header = "User" },
            @{ key = "DisplayName"; header = "Display Name" },
            @{ key = "AccountEnabled"; header = "State"; type = "pill" },
            @{ key = "MFARegistered"; header = "MFA"; type = "pill" },
            @{ key = "LastSignIn"; header = "Last Sign-In" },
            @{ key = "HardeningNotes"; header = "Notes" }
        )
        rows = @($rows | Sort-Object UserPrincipalName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
