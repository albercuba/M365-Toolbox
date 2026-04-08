[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$MaxUsersToInspect = 250
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "Directory.Read.All", "AuditLog.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING MFA CAMPAIGN DATA"

$users = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,department,country,jobTitle,accountEnabled,signInActivity&`$top={0}" -f $MaxUsersToInspect))
$rows = [System.Collections.Generic.List[object]]::new()

foreach ($user in $users) {
    if (-not $user.userPrincipalName) { continue }
    $methods = @()
    try { $methods = @(Invoke-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/methods" -f $user.id)) } catch {}

    $methodNames = @(
        foreach ($method in $methods) {
            switch -Regex ([string]$method.'@odata.type') {
                'microsoftAuthenticatorAuthenticationMethod' { 'Authenticator App'; break }
                'phoneAuthenticationMethod' { 'Phone'; break }
                'fido2AuthenticationMethod' { 'FIDO2 Key'; break }
                'softwareOathAuthenticationMethod' { 'Software OATH'; break }
                'emailAuthenticationMethod' { 'Email OTP'; break }
                'windowsHelloForBusinessAuthenticationMethod' { 'Windows Hello'; break }
            }
        }
    )

    $lastSignIn = $null
    if ($user.signInActivity -and $user.signInActivity.lastSignInDateTime) { $lastSignIn = Get-Date $user.signInActivity.lastSignInDateTime }

    [void]$rows.Add([pscustomobject]@{
        DisplayName       = [string]$user.displayName
        UserPrincipalName = [string]$user.userPrincipalName
        Department        = [string]$user.department
        Country           = [string]$user.country
        JobTitle          = [string]$user.jobTitle
        Enabled           = if ($user.accountEnabled) { "Enabled" } else { "Disabled" }
        MfaRegistered     = if ($methodNames.Count -gt 0) { "Registered" } else { "Not Registered" }
        MethodCount       = $methodNames.Count
        Methods           = if ($methodNames.Count -gt 0) { ($methodNames | Sort-Object -Unique) -join ", " } else { "None" }
        LastSignIn        = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd") } else { "Unknown" }
    })
}

$unregistered = @($rows | Where-Object { $_.MfaRegistered -eq "Not Registered" })
$deptBreakdown = @($rows | Group-Object Department | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
    [pscustomobject]@{
        Department   = if ($_.Name) { $_.Name } else { "Unassigned" }
        Users        = $_.Count
        Unregistered = @($_.Group | Where-Object { $_.MfaRegistered -eq "Not Registered" }).Count
    }
})

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MfaRegistrationCampaign" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 MFA Registration Campaign Report" -Tenant $tenantName -Subtitle "User-by-user MFA registration readiness for adoption campaigns" -Kpis @(
    @{ label = "Users"; value = $rows.Count; sub = "Inspected users"; cls = "neutral" },
    @{ label = "Registered"; value = @($rows | Where-Object { $_.MfaRegistered -eq "Registered" }).Count; sub = "With methods"; cls = "ok" },
    @{ label = "Not Registered"; value = $unregistered.Count; sub = "Campaign targets"; cls = if ($unregistered.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Departments"; value = @($rows | Select-Object -ExpandProperty Department -Unique).Count; sub = "Distinct departments"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Inspection Limit"; value = $MaxUsersToInspect },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Department Breakdown"
        badge = "$($deptBreakdown.Count) departments"
        columns = @(
            @{ key = "Department"; header = "Department" },
            @{ key = "Users"; header = "Users" },
            @{ key = "Unregistered"; header = "Unregistered" }
        )
        rows = $deptBreakdown
    },
    @{
        title = "User Registration Status"
        badge = "$($rows.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "User" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "Department"; header = "Department" },
            @{ key = "Country"; header = "Country" },
            @{ key = "MfaRegistered"; header = "MFA"; type = "pill" },
            @{ key = "Methods"; header = "Methods" },
            @{ key = "LastSignIn"; header = "Last Sign-In" }
        )
        rows = @($rows | Sort-Object MfaRegistered, Department, DisplayName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
