[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [switch]$IncludeGuests = $false,

    [string[]]$AdminRoles = @(
        "Global Administrator",
        "Privileged Role Administrator",
        "Security Administrator",
        "Exchange Administrator",
        "SharePoint Administrator",
        "Teams Administrator",
        "User Administrator",
        "Billing Administrator",
        "Application Administrator",
        "Cloud Application Administrator",
        "Conditional Access Administrator"
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$script:TenantDomain = ""
$script:CanResolveAdminRoles = $true

function Write-SectionHeader {
    param([string]$Title)
    $line = "=" * 50
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host (" " * [Math]::Floor((50 - $Title.Length) / 2) + $Title) -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Assert-RequiredModules {
    $required = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.SignIns",
        "Microsoft.Graph.DirectoryObjects"
    )

    Write-Host ""
    Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan

    foreach ($moduleName in $required) {
        $installed = Get-Module -ListAvailable -Name $moduleName |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $installed) {
            throw "Required module '$moduleName' is not installed."
        }

        Import-Module $moduleName -Force -WarningAction SilentlyContinue
        Write-Host "  [+] $moduleName v$($installed.Version)" -ForegroundColor Green
    }

    $directoryManagementModule = Get-Module -ListAvailable -Name "Microsoft.Graph.Identity.DirectoryManagement" |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($directoryManagementModule) {
        Import-Module "Microsoft.Graph.Identity.DirectoryManagement" -Force -WarningAction SilentlyContinue
        Write-Host "  [+] Microsoft.Graph.Identity.DirectoryManagement v$($directoryManagementModule.Version)" -ForegroundColor Green
        $script:CanResolveAdminRoles = $true
    }
    else {
        Write-Warning "  [!] Microsoft.Graph.Identity.DirectoryManagement is not installed. Admin role resolution will be skipped."
        $script:CanResolveAdminRoles = $false
    }
}

function Connect-ToGraph {
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )

    $ErrorActionPreference = 'Stop'

    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "          CONNECTING TO MICROSOFT GRAPH" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "[*] Requested tenant: $TenantId"
    Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
    Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

    Disconnect-MgGraph -ErrorAction SilentlyContinue

    try {
        Connect-MgGraph `
            -TenantId $TenantId `
            -Scopes "User.Read.All","Directory.Read.All","AuditLog.Read.All" `
            -UseDeviceAuthentication `
            -NoWelcome `
            -ContextScope Process `
            -ErrorAction Stop

        Write-Host "[+] Connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        $msg = $_.Exception.Message

        if ($msg -match "AADSTS90002") {
            throw "Tenant '$TenantId' was not found. Check the tenant name."
        }

        throw "Failed to connect to Microsoft Graph: $msg"
    }
}

function Ensure-ReportScopes {
    $requiredScopes = @(
        "User.Read.All",
        "UserAuthenticationMethod.Read.All",
        "Directory.Read.All",
        "RoleManagement.Read.Directory",
        "AuditLog.Read.All",
        "Organization.Read.All"
    )

    $ctx = Get-MgContext
    $grantedScopes = @()
    if ($ctx -and $ctx.Scopes) {
        $grantedScopes = @($ctx.Scopes)
    }

    $missingScopes = @($requiredScopes | Where-Object { $_ -notin $grantedScopes })
    if ($missingScopes.Count -eq 0) {
        return
    }

    Write-Host "[*] Reconnecting with additional report scopes..." -ForegroundColor Cyan
    Write-Host "[*] Additional consent may be required to read authentication methods and role memberships." -ForegroundColor DarkCyan

    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Connect-MgGraph `
        -TenantId $TenantId `
        -Scopes $requiredScopes `
        -UseDeviceAuthentication `
        -NoWelcome `
        -ContextScope Process `
        -ErrorAction Stop
}

function Resolve-TenantDomain {
    try {
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        if ($org -and $org.VerifiedDomains) {
            $defaultDomain = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
            $script:TenantDomain = if ($defaultDomain) { [string]$defaultDomain.Name } else { [string]$org.VerifiedDomains[0].Name }
        }
    }
    catch {
        $ctx = Get-MgContext
        if ($ctx -and $ctx.TenantId) {
            $script:TenantDomain = [string]$ctx.TenantId
        }
    }
}

function Get-AdminUserIds {
    param([string[]]$RoleNames)

    Write-Host "[*] Resolving privileged role assignments..." -ForegroundColor Cyan
    $adminIds = [System.Collections.Generic.HashSet[string]]::new()

    if (-not $script:CanResolveAdminRoles) {
        Write-Warning "  [!] Skipping privileged role lookup because Microsoft.Graph.Identity.DirectoryManagement is unavailable."
        return $adminIds
    }

    try {
        $roles = Get-MgDirectoryRole -All -ErrorAction Stop
        foreach ($role in $roles) {
            if ($RoleNames -contains $role.DisplayName) {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
                foreach ($member in $members) {
                    [void]$adminIds.Add([string]$member.Id)
                }
            }
        }
        Write-Host "  [+] $($adminIds.Count) privileged user(s) found." -ForegroundColor $(if ($adminIds.Count -gt 0) { "Green" } else { "Yellow" })
    }
    catch {
        Write-Warning "  [!] Could not retrieve role memberships: $_"
    }

    return $adminIds
}

function Get-MfaReport {
    param(
        [bool]$IncludeGuests,
        [System.Collections.Generic.HashSet[string]]$AdminIds
    )

    Write-SectionHeader "COLLECTING MFA AUTHENTICATION DATA"

    $report = [System.Collections.Generic.List[object]]::new()
    $filter = if ($IncludeGuests) { "accountEnabled eq true" } else { "accountEnabled eq true and userType eq 'Member'" }

    $hasPremiumLicense = $false
    Write-Host "[*] Probing for Entra ID P1/P2 (SignInActivity support)..." -ForegroundColor Cyan
    try {
        Get-MgUser -Filter $filter -Top 1 `
            -Property Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,SignInActivity,Mail `
            -ErrorAction Stop | Out-Null
        $hasPremiumLicense = $true
        Write-Host "  [+] Premium license detected - Last Sign-In data will be included." -ForegroundColor Green
    }
    catch {
        if ($_ -match "403" -or $_ -match "Forbidden" -or $_ -match "NonPremium" -or $_ -match "RequestFromNonPremiumTenant") {
            Write-Host "  [!] No Entra ID P1/P2 license detected - Last Sign-In will show as N/A for this tenant." -ForegroundColor Yellow
        }
        else {
            throw
        }
    }

    Write-Host "[*] Fetching users (filter: $filter)..." -ForegroundColor Cyan
    $userProps = "Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,Mail"
    if ($hasPremiumLicense) {
        $userProps += ",SignInActivity"
    }

    $users = @(Get-MgUser -Filter $filter -All -Property $userProps -ErrorAction Stop)
    Write-Host "  [+] $($users.Count) user(s) to process." -ForegroundColor Green

    $i = 0
    foreach ($user in $users) {
        $i++
        if ($i -eq 1 -or $i % 25 -eq 0 -or $i -eq $users.Count) {
            Write-Host "  [*] Processing user $i / $($users.Count) ..." -ForegroundColor Cyan
        }

        $methodNames = [System.Collections.Generic.List[string]]::new()
        $hasAuthApp = $false
        $hasFido2 = $false
        $hasPhone = $false
        $hasTotp = $false
        $hasEmail = $false
        $hasWhfb = $false
        $hasPasskey = $false
        $defaultMethod = "None"

        try {
            $methods = @(Get-MgUserAuthenticationMethod -UserId $user.Id -All -ErrorAction Stop)

            foreach ($method in $methods) {
                $additionalProperties = $method.AdditionalProperties
                if (-not $additionalProperties) {
                    $additionalProperties = @{}
                }

                $odata = [string]$additionalProperties['@odata.type']
                switch ($odata) {
                    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                        $hasAuthApp = $true
                        $appMode = [string]$additionalProperties['authenticationMode']
                        [void]$methodNames.Add("Authenticator App" + $(if ($appMode -eq "passwordless") { " (Passwordless)" } else { "" }))
                    }
                    "#microsoft.graph.fido2AuthenticationMethod" {
                        $hasFido2 = $true
                        [void]$methodNames.Add("FIDO2 Security Key")
                    }
                    "#microsoft.graph.phoneAuthenticationMethod" {
                        $hasPhone = $true
                        $phoneType = [string]$additionalProperties['phoneType']
                        [void]$methodNames.Add("Phone ($phoneType)")
                    }
                    "#microsoft.graph.softwareOathAuthenticationMethod" {
                        $hasTotp = $true
                        [void]$methodNames.Add("Software TOTP (OATH)")
                    }
                    "#microsoft.graph.emailAuthenticationMethod" {
                        $hasEmail = $true
                        [void]$methodNames.Add("Email OTP")
                    }
                    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                        $hasWhfb = $true
                        [void]$methodNames.Add("Windows Hello for Business")
                    }
                    "#microsoft.graph.passkeyAuthenticationMethod" {
                        $hasPasskey = $true
                        [void]$methodNames.Add("Passkey")
                    }
                    "#microsoft.graph.passwordAuthenticationMethod" {
                    }
                    default {
                        if ($odata -and $odata -ne "#microsoft.graph.passwordAuthenticationMethod") {
                            $friendly = $odata -replace ".*\.", "" -replace "AuthenticationMethod", ""
                            [void]$methodNames.Add($friendly)
                        }
                    }
                }
            }

            $defaultMethod = if ($hasFido2) { "FIDO2 Key" }
                elseif ($hasPasskey) { "Passkey" }
                elseif ($hasWhfb) { "WHfB" }
                elseif ($hasAuthApp) { "Authenticator App" }
                elseif ($hasTotp) { "TOTP" }
                elseif ($hasPhone) { "Phone" }
                elseif ($hasEmail) { "Email OTP" }
                else { "Password Only" }
        }
        catch {
            $defaultMethod = "Error reading methods"
        }

        $hasMfa = $hasAuthApp -or $hasFido2 -or $hasPhone -or $hasTotp -or $hasEmail -or $hasWhfb -or $hasPasskey
        $lastSignIn = if (-not $hasPremiumLicense) { "N/A (no P1/P2 license)" } else { "Never / Unknown" }
        if ($hasPremiumLicense -and $user.SignInActivity -and $user.SignInActivity.LastSignInDateTime) {
            $lastSignIn = "{0:yyyy-MM-dd HH:mm}" -f $user.SignInActivity.LastSignInDateTime
        }

        $report.Add([PSCustomObject]@{
            DisplayName   = $user.DisplayName
            UPN           = $user.UserPrincipalName
            UserType      = $user.UserType
            IsAdmin       = $AdminIds.Contains([string]$user.Id)
            MfaStatus     = if ($hasMfa) { "Registered" } else { "NOT Registered" }
            MfaRegistered = $hasMfa
            DefaultMethod = $defaultMethod
            MethodCount   = @($methodNames | Select-Object -Unique).Count
            Methods       = (@($methodNames | Select-Object -Unique) -join "; ")
            HasAuthApp    = $hasAuthApp
            HasFIDO2      = $hasFido2
            HasPhone      = $hasPhone
            HasTOTP       = $hasTotp
            HasEmailOTP   = $hasEmail
            HasWHfB       = $hasWhfb
            HasPasskey    = $hasPasskey
            LastSignIn    = $lastSignIn
        }) | Out-Null
    }

    Write-Host "  [+] Data collection complete. $($report.Count) records." -ForegroundColor Green
    return $report
}

function Show-ConsoleSummary {
    param([object[]]$Report)

    Write-SectionHeader "MFA SUMMARY"

    $total = $Report.Count
    $registered = @($Report | Where-Object { $_.MfaRegistered }).Count
    $notRegistered = $total - $registered
    $admins = @($Report | Where-Object { $_.IsAdmin }).Count
    $adminNoMfa = @($Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
    $pct = if ($total -gt 0) { [math]::Round(($registered / $total) * 100, 1) } else { 0 }

    Write-Host ""
    Write-Host ("  Tenant              : {0}" -f $(if ($script:TenantDomain) { $script:TenantDomain } else { $TenantId })) -ForegroundColor Cyan
    Write-Host ("  Total Users         : {0,6}" -f $total) -ForegroundColor Cyan
    Write-Host ("  MFA Registered      : {0,6}  ({1}%)" -f $registered, $pct) -ForegroundColor $(if ($pct -ge 90) { "Green" } elseif ($pct -ge 70) { "Yellow" } else { "Red" })
    Write-Host ("  MFA NOT Registered  : {0,6}" -f $notRegistered) -ForegroundColor $(if ($notRegistered -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host ("  Admin Accounts      : {0,6}" -f $admins) -ForegroundColor Cyan
    Write-Host ("  Admins without MFA  : {0,6}" -f $adminNoMfa) -ForegroundColor $(if ($adminNoMfa -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host "  Method breakdown:" -ForegroundColor Yellow
    Write-Host ("    Authenticator App         : {0}" -f @($Report | Where-Object { $_.HasAuthApp }).Count) -ForegroundColor Cyan
    Write-Host ("    FIDO2 Security Key        : {0}" -f @($Report | Where-Object { $_.HasFIDO2 }).Count) -ForegroundColor Cyan
    Write-Host ("    Phone (SMS/Voice)         : {0}" -f @($Report | Where-Object { $_.HasPhone }).Count) -ForegroundColor Cyan
    Write-Host ("    Software TOTP (OATH)      : {0}" -f @($Report | Where-Object { $_.HasTOTP }).Count) -ForegroundColor Cyan
    Write-Host ("    Email OTP                 : {0}" -f @($Report | Where-Object { $_.HasEmailOTP }).Count) -ForegroundColor Cyan
    Write-Host ("    Windows Hello for Business: {0}" -f @($Report | Where-Object { $_.HasWHfB }).Count) -ForegroundColor Cyan
    Write-Host ("    Passkey                   : {0}" -f @($Report | Where-Object { $_.HasPasskey }).Count) -ForegroundColor Cyan
    Write-Host ""

    if ($adminNoMfa -gt 0) {
        Write-Host "  [!] CRITICAL - Admin accounts without MFA:" -ForegroundColor Red
        $Report |
            Where-Object { $_.IsAdmin -and -not $_.MfaRegistered } |
            Select-Object DisplayName, UPN, LastSignIn |
            Format-Table -AutoSize |
            Out-Host
    }
}

Assert-RequiredModules
Connect-ToGraph -TenantId $TenantId
Ensure-ReportScopes
Resolve-TenantDomain
$adminIds = Get-AdminUserIds -RoleNames $AdminRoles
$report = Get-MfaReport -IncludeGuests $IncludeGuests.IsPresent -AdminIds $adminIds
Show-ConsoleSummary -Report $report
