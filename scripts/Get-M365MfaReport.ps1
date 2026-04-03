<#
.SYNOPSIS
    Generates a comprehensive MFA authentication methods report for all users
    in a Microsoft 365 tenant.

.DESCRIPTION
    Connects to Microsoft Graph with delegated device code authentication and reports on:
      - Per-user MFA registration status and enabled methods
      - Authenticator App, FIDO2, Phone, TOTP, Email, Windows Hello, etc.
      - Users with MFA disabled / not registered
      - Privileged / admin account MFA coverage
      - Tenant-wide summary statistics

    Outputs a formatted report to the console and exports to Excel (.xlsx)
    and/or a self-contained HTML dashboard matching the RegITs dark theme.

    Requires the Microsoft.Graph PowerShell SDK (auto-installed if missing).

.PARAMETER ExportXlsx
    Folder or file path for the Excel export.
    Example: -ExportXlsx "C:\Reports\"

.PARAMETER ExportHtml
    Folder or file path for the HTML dashboard export.
    Example: -ExportHtml "C:\Reports\"

.PARAMETER IncludeGuests
    Include guest/external users in the report. Default: $false

.PARAMETER TenantId
    Optional Microsoft Entra tenant ID or primary tenant domain
    (for example: contoso.onmicrosoft.com). Recommended when the
    sign-in flow cannot infer the tenant automatically.

.PARAMETER AdminRoles
    List of role display-names to flag as privileged.
    Default covers the most critical built-in roles.

.EXAMPLE
    .\Get-M365MfaReport.ps1 -ExportXlsx "C:\Reports\" -ExportHtml "C:\Reports\"

.EXAMPLE
    .\Get-M365MfaReport.ps1 -ExportHtml "C:\Reports\" -IncludeGuests
#>

[CmdletBinding()]
param (
    [string]   $ExportXlsx,
    [string]   $ExportHtml,
    [switch]   $IncludeGuests    = $false,
    [string]   $TenantId         = $(if ($env:M365_TENANT_ID) { $env:M365_TENANT_ID } elseif ($env:AZURE_TENANT_ID) { $env:AZURE_TENANT_ID } else { $null }),
    [string[]] $AdminRoles       = @(
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
$script:GraphDeviceCodeClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
$script:ImportExcelAutoSizeSupported = $IsWindows

# ─────────────────────────────────────────────
# HELPER: Timestamp + filename builder
# ─────────────────────────────────────────────

$script:TenantDomain = ""

function Add-TimestampToPath {
    param ([string]$Path, [string]$DefaultExtension = "")
    if (-not $Path) { return $Path }
    $stamp  = Get-Date -Format "dd.MM.yy-HH.mm.ss"
    $tenant = if ($script:TenantDomain) { $script:TenantDomain.Split('.')[0] } else { "M365" }
    $base   = "MFA_${tenant}_$stamp"
    if (Test-Path $Path -PathType Container) { return Join-Path $Path "$base$DefaultExtension" }
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    $ext = [System.IO.Path]::GetExtension($Path)
    if (-not $ext -and $DefaultExtension) { $ext = $DefaultExtension }
    if ($dir) { return [System.IO.Path]::Combine($dir, "$base$ext") }
    return "$base$ext"
}

# ─────────────────────────────────────────────
# HELPER: Section header
# ─────────────────────────────────────────────

function Write-SectionHeader {
    param ([string]$Title)
    $line = "=" * 50
    Write-Host "`n$line" -ForegroundColor Cyan
    $pad = [math]::Floor((50 - $Title.Length) / 2)
    Write-Host (" " * $pad + $Title) -ForegroundColor Cyan
    Write-Host "$line" -ForegroundColor Cyan
}

function Get-HttpErrorDetails {
    param(
        [Parameter(Mandatory = $true)]
        $ErrorRecord
    )

    $message = $ErrorRecord.Exception.Message

    try {
        if ($ErrorRecord.Exception.Response) {
            $responseStream = $ErrorRecord.Exception.Response.GetResponseStream()
            if ($responseStream) {
                $reader = [System.IO.StreamReader]::new($responseStream)
                try {
                    $body = $reader.ReadToEnd()
                    if ($body) {
                        return "$message`n$body"
                    }
                }
                finally {
                    $reader.Dispose()
                }
            }
        }
    }
    catch {}

    return $message
}

function Invoke-OAuthFormRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [hashtable]$Body,

        [int]$TimeoutSec = 30
    )

    $response = Invoke-WebRequest -Method Post -Uri $Uri -Body $Body `
        -ContentType "application/x-www-form-urlencoded" -TimeoutSec $TimeoutSec `
        -SkipHttpErrorCheck -ErrorAction Stop

    $statusCode = [int]$response.StatusCode
    $content = [string]$response.Content
    $json = $null

    if ($content) {
        try {
            $json = $content | ConvertFrom-Json -ErrorAction Stop
        }
        catch {}
    }

    return [PSCustomObject]@{
        StatusCode = $statusCode
        Content    = $content
        Json       = $json
    }
}

# ─────────────────────────────────────────────
# PREREQUISITES: Module check / install
# ─────────────────────────────────────────────

function Assert-RequiredModules {
    param(
        [bool]$NeedExcelExport = $true
    )

    $required = @(
        @{ Name = "Microsoft.Graph.Authentication";   MinVersion = "2.0.0" }
        @{ Name = "Microsoft.Graph.Users";            MinVersion = "2.0.0" }
        @{ Name = "Microsoft.Graph.Identity.SignIns"; MinVersion = "2.0.0" }
        @{ Name = "Microsoft.Graph.DirectoryObjects"; MinVersion = "2.0.0" }
    )

    if ($NeedExcelExport) {
        $required += @{ Name = "ImportExcel"; MinVersion = "7.0.0" }
    }

    Write-Host "`n[*] Checking required PowerShell modules..." -ForegroundColor Cyan
    $xlAvailable = $false

    foreach ($mod in $required) {
        $installed = Get-Module -ListAvailable -Name $mod.Name |
                     Sort-Object Version -Descending | Select-Object -First 1

        if (-not $installed) {
            Write-Host "  [!] Module '$($mod.Name)' not found. Installing..." -ForegroundColor Yellow
            try {
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
                }
                Install-Module -Name $mod.Name -MinimumVersion $mod.MinVersion `
                    -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Host "  [+] Installed $($mod.Name)" -ForegroundColor Green
                $installed = Get-Module -ListAvailable -Name $mod.Name |
                             Sort-Object Version -Descending | Select-Object -First 1
            }
            catch {
                Write-Warning "  [!] Could not install '$($mod.Name)': $_"
                continue
            }
        }
        else {
            Write-Host "  [+] $($mod.Name) v$($installed.Version)" -ForegroundColor Green
        }

        try {
            $originalWarningPreference = $WarningPreference
            if ($mod.Name -eq "ImportExcel") {
                $WarningPreference = "SilentlyContinue"
            }

            if ($mod.Name -eq "ImportExcel") {
                & { Import-Module $mod.Name -Force -WarningAction SilentlyContinue } 3>$null 6>$null
            }
            else {
                Import-Module $mod.Name -Force -WarningAction SilentlyContinue
            }
            if ($mod.Name -eq "ImportExcel") { $xlAvailable = $true }
        }
        catch { Write-Warning "  [!] Could not import '$($mod.Name)': $_" }
        finally {
            $WarningPreference = $originalWarningPreference
        }
    }

    return $xlAvailable
}

# ─────────────────────────────────────────────
# SECTION 1 — Connect to Microsoft Graph
# ─────────────────────────────────────────────

function Connect-ToGraph {
    Write-SectionHeader "CONNECTING TO MICROSOFT GRAPH"

    $scopes = @(
        "User.Read.All",
        "UserAuthenticationMethod.Read.All",
        "Directory.Read.All",
        "RoleManagement.Read.Directory",
        "AuditLog.Read.All",
        "Reports.Read.All",
        "Organization.Read.All"
    )
    $deviceCodeScopes = @(
        $scopes,
        "offline_access",
        "openid",
        "profile"
    )

    try {
        $tenantSegment = "organizations"
        $deviceCodeUri = "https://login.microsoftonline.com/$tenantSegment/oauth2/v2.0/devicecode"
        $tokenUri = "https://login.microsoftonline.com/$tenantSegment/oauth2/v2.0/token"
        $deviceCodeScope = ($deviceCodeScopes | Select-Object -Unique) -join " "

        if ($TenantId) {
            Write-Host "[*] Requested tenant: $TenantId" -ForegroundColor Cyan
            Write-Host "[*] Using the organizations sign-in endpoint and validating the tenant after authentication." -ForegroundColor DarkCyan
        }

        Write-Host "[*] Starting device code sign-in for Microsoft Graph..." -ForegroundColor Cyan
        Write-Host "[*] Use an admin account that can read users, authentication methods, roles, audit logs, and reports." -ForegroundColor DarkCyan
        Write-Host "[*] Requesting device code from Microsoft Entra ID..." -ForegroundColor Cyan

        try {
            $deviceCodeResult = Invoke-OAuthFormRequest -Uri $deviceCodeUri -Body @{
                client_id = $script:GraphDeviceCodeClientId
                scope     = $deviceCodeScope
            } -TimeoutSec 30
        }
        catch {
            throw "Failed to request a device code from Microsoft Entra ID.`n$(Get-HttpErrorDetails -ErrorRecord $_)"
        }

        if ($deviceCodeResult.StatusCode -lt 200 -or $deviceCodeResult.StatusCode -ge 300) {
            $deviceCodeError = if ($deviceCodeResult.Content) { $deviceCodeResult.Content } else { "No response body returned." }
            throw "Failed to request a device code from Microsoft Entra ID. HTTP $($deviceCodeResult.StatusCode).`n$deviceCodeError"
        }

        $deviceCodeResponse = $deviceCodeResult.Json
        if (-not $deviceCodeResponse) {
            throw "Failed to parse the device code response from Microsoft Entra ID.`n$($deviceCodeResult.Content)"
        }

        Write-Host ""
        if ($deviceCodeResponse.message) {
            Write-Host $deviceCodeResponse.message -ForegroundColor Yellow
            Write-Host ""
        }
        Write-Host "Sign in with your admin account using the device code flow:" -ForegroundColor Yellow
        Write-Host "1. Open: $($deviceCodeResponse.verification_uri)" -ForegroundColor Yellow
        $verificationUriComplete = $null
        if ($deviceCodeResponse.PSObject.Properties.Match("verification_uri_complete").Count -gt 0) {
            $verificationUriComplete = [string]$deviceCodeResponse.verification_uri_complete
        }
        if ($verificationUriComplete) {
            Write-Host "   Direct link: $verificationUriComplete" -ForegroundColor Yellow
        }
        Write-Host "2. Enter code: $($deviceCodeResponse.user_code)" -ForegroundColor Yellow
        Write-Host "3. Complete the sign-in and consent prompt for the tenant you want to review." -ForegroundColor Yellow
        Write-Host ""

        $pollIntervalSeconds = [int]$deviceCodeResponse.interval
        $deadline = (Get-Date).AddSeconds([int]$deviceCodeResponse.expires_in)
        $accessToken = $null

        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds $pollIntervalSeconds

            try {
                $tokenResult = Invoke-OAuthFormRequest -Uri $tokenUri -Body @{
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    client_id   = $script:GraphDeviceCodeClientId
                    device_code = $deviceCodeResponse.device_code
                } -TimeoutSec 30
            }
            catch {
                $errorDetails = Get-HttpErrorDetails -ErrorRecord $_
                throw "Device code token polling failed.`n$errorDetails"
            }

            if ($tokenResult.StatusCode -ge 200 -and $tokenResult.StatusCode -lt 300) {
                if (-not $tokenResult.Json) {
                    throw "Failed to parse the token response from Microsoft Entra ID.`n$($tokenResult.Content)"
                }

                $accessToken = $tokenResult.Json.access_token
                break
            }

            $oauthError = $tokenResult.Json
            if ($oauthError) {
                switch ($oauthError.error) {
                    "authorization_pending" { continue }
                    "slow_down" {
                        $pollIntervalSeconds += 5
                        continue
                    }
                    "authorization_declined" { throw "Device code sign-in was declined." }
                    "expired_token" { throw "Device code expired before sign-in completed." }
                    "bad_verification_code" { throw "Invalid device code returned by the authorization server." }
                    default {
                        if ($oauthError.error_description) {
                            throw $oauthError.error_description
                        }
                    }
                }
            }

            $tokenError = if ($tokenResult.Content) { $tokenResult.Content } else { "No response body returned." }
            throw "Device code token polling failed with HTTP $($tokenResult.StatusCode).`n$tokenError"
        }

        if (-not $accessToken) {
            throw "Timed out waiting for device code sign-in to complete."
        }

        $secureAccessToken = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
        Write-Host "[*] Device code authentication completed. Connecting the Graph PowerShell session..." -ForegroundColor Cyan
        Connect-MgGraph -AccessToken $secureAccessToken -NoWelcome -ContextScope Process -ErrorAction Stop
        $ctx = Get-MgContext
        Write-Host "[+] Connected account: $($ctx.Account)" -ForegroundColor Green
        Write-Host "[+] Tenant ID        : $($ctx.TenantId)" -ForegroundColor Green

        # Resolve tenant domain
        try {
            $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($org -and $org.VerifiedDomains) {
                $primary = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -First 1
                $script:TenantDomain = if ($primary) { $primary.Name } else { $org.VerifiedDomains[0].Name }
                Write-Host "[+] Tenant domain: $($script:TenantDomain)" -ForegroundColor Green
            }
        }
        catch { $script:TenantDomain = $ctx.TenantId }

        if ($TenantId) {
            $requestedTenant = $TenantId.Trim().ToLowerInvariant()
            $tenantMatches = $false
            $parsedTenantGuid = [guid]::Empty

            if ([guid]::TryParse($requestedTenant, [ref]$parsedTenantGuid)) {
                $tenantMatches = [string]::Equals([string]$ctx.TenantId, $requestedTenant, [System.StringComparison]::OrdinalIgnoreCase)
            }
            else {
                $verifiedDomains = @()
                if ($org -and $org.VerifiedDomains) {
                    $verifiedDomains = @($org.VerifiedDomains | ForEach-Object { [string]$_.Name })
                }

                $tenantMatches = $verifiedDomains | Where-Object {
                    [string]::Equals($_, $requestedTenant, [System.StringComparison]::OrdinalIgnoreCase)
                } | Select-Object -First 1
            }

            if (-not $tenantMatches) {
                $resolvedTenant = if ($script:TenantDomain) { $script:TenantDomain } else { [string]$ctx.TenantId }
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
                throw "Signed in to tenant '$resolvedTenant', but the requested tenant was '$TenantId'. Sign in again with an account from the requested tenant."
            }
        }

        return $ctx
    }
    catch {
        $errorText = $_.Exception.Message

        if ($errorText -match "AADSTS50059") {
            $suggestion = if ($TenantId) {
                "The provided tenant value '$TenantId' could not be resolved. Verify that it is a valid tenant GUID or domain such as 'contoso.onmicrosoft.com'."
            }
            else {
                "No tenant-identifying information was supplied. Re-run the report with a Tenant ID or set the environment variable M365_TENANT_ID."
            }

            Write-Error "Failed to connect to Microsoft Graph: $suggestion`nOriginal error: $errorText"
        }
        else {
            Write-Error "Failed to connect to Microsoft Graph: $errorText"
        }

        exit 1
    }
}

# ─────────────────────────────────────────────
# SECTION 2 — Collect admin role memberships
# ─────────────────────────────────────────────

function Get-AdminUserIds {
    param ([string[]]$RoleNames)
    Write-Host "[*] Resolving privileged role assignments..." -ForegroundColor Cyan
    $adminIds = [System.Collections.Generic.HashSet[string]]::new()

    try {
        $roles = Get-MgDirectoryRole -All -ErrorAction Stop
        foreach ($role in $roles) {
            if ($RoleNames -contains $role.DisplayName) {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
                foreach ($m in $members) { $adminIds.Add($m.Id) | Out-Null }
            }
        }
        Write-Host "  [+] $($adminIds.Count) privileged user(s) found." -ForegroundColor $(if ($adminIds.Count -gt 0) { "Green" } else { "Yellow" })
    }
    catch { Write-Warning "  [!] Could not retrieve role memberships: $_" }

    return $adminIds
}

# ─────────────────────────────────────────────
# SECTION 3 — Main MFA data collection
# ─────────────────────────────────────────────

function Get-MfaReport {
    param (
        [bool]   $IncludeGuests,
        [System.Collections.Generic.HashSet[string]]$AdminIds
    )

    Write-SectionHeader "COLLECTING MFA AUTHENTICATION DATA"

    $report  = [System.Collections.Generic.List[PSObject]]::new()
    $filter  = if ($IncludeGuests) { "accountEnabled eq true" } `
               else { "accountEnabled eq true and userType eq 'Member'" }

    # ── Detect whether the tenant has Entra ID P1/P2 (needed for SignInActivity) ──
    # We probe with a single user; if the tenant lacks the license we get a 403
    # and silently fall back to fetching without SignInActivity.
    $hasPremiumLicense = $false
    Write-Host "[*] Probing for Entra ID P1/P2 (SignInActivity support)..." -ForegroundColor Cyan
    try {
        Get-MgUser -Filter $filter -Top 1 `
            -Property Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,SignInActivity,Mail `
            -ErrorAction Stop | Out-Null
        $hasPremiumLicense = $true
        Write-Host "  [+] Premium license detected — Last Sign-In data will be included." -ForegroundColor Green
    }
    catch {
        if ($_ -match "403" -or $_ -match "Forbidden" -or $_ -match "NonPremium" -or $_ -match "RequestFromNonPremiumTenant") {
            Write-Host "  [!] No Entra ID P1/P2 license detected — Last Sign-In will show as N/A for this tenant." -ForegroundColor Yellow
        }
        else { throw }
    }

    Write-Host "[*] Fetching users (filter: $filter)..." -ForegroundColor Cyan
    $userProps = 'Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,Mail'
    if ($hasPremiumLicense) { $userProps += ',SignInActivity' }

    $users = Get-MgUser -Filter $filter -All -Property $userProps -ErrorAction Stop
    Write-Host "  [+] $($users.Count) user(s) to process." -ForegroundColor Green

    $i = 0
    foreach ($user in $users) {
        $i++
        if ($i % 25 -eq 0 -or $i -eq 1) {
            Write-Host "  [*] Processing user $i / $($users.Count) ..." -ForegroundColor Cyan
        }

        # Defaults
        $methodNames   = [System.Collections.Generic.List[string]]::new()
        $hasMfa        = $false
        $hasAuthApp    = $false
        $hasFido2      = $false
        $hasPhone      = $false
        $hasTotp       = $false
        $hasEmail      = $false
        $hasWhfb       = $false
        $hasPasskey    = $false
        $defaultMethod = "None"
        $methodCount   = 0

        try {
            $methods = Get-MgUserAuthenticationMethod -UserId $user.Id -All -ErrorAction Stop

            foreach ($m in $methods) {
                $odata = $m.AdditionalProperties['@odata.type']
                switch ($odata) {
                    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                        $hasAuthApp = $true
                        $appType    = $m.AdditionalProperties['authenticationMode']
                        $methodNames.Add("Authenticator App" + $(if ($appType -eq "passwordless") { " (Passwordless)" } else { "" })) | Out-Null
                    }
                    "#microsoft.graph.fido2AuthenticationMethod" {
                        $hasFido2 = $true
                        $methodNames.Add("FIDO2 Security Key") | Out-Null
                    }
                    "#microsoft.graph.phoneAuthenticationMethod" {
                        $hasPhone  = $true
                        $phoneType = $m.AdditionalProperties['phoneType']
                        $methodNames.Add("Phone ($phoneType)") | Out-Null
                    }
                    "#microsoft.graph.softwareOathAuthenticationMethod" {
                        $hasTotp = $true
                        $methodNames.Add("Software TOTP (OATH)") | Out-Null
                    }
                    "#microsoft.graph.emailAuthenticationMethod" {
                        $hasEmail = $true
                        $methodNames.Add("Email OTP") | Out-Null
                    }
                    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                        $hasWhfb = $true
                        $methodNames.Add("Windows Hello for Business") | Out-Null
                    }
                    "#microsoft.graph.passkeyAuthenticationMethod" {
                        $hasPasskey = $true
                        $methodNames.Add("Passkey") | Out-Null
                    }
                    "#microsoft.graph.passwordAuthenticationMethod" {
                        # Password — not an MFA method, skip count
                    }
                    default {
                        if ($odata -and $odata -ne "#microsoft.graph.passwordAuthenticationMethod") {
                            $friendly = $odata -replace ".*\.", "" -replace "AuthenticationMethod", ""
                            $methodNames.Add($friendly) | Out-Null
                        }
                    }
                }
            }

            # Determine "has MFA" = at least one real second factor registered
            $hasMfa = $hasAuthApp -or $hasFido2 -or $hasPhone -or $hasTotp -or $hasEmail -or $hasWhfb -or $hasPasskey
            $methodCount = $methodNames.Count

            # Derive default/strongest method label
            $defaultMethod = if ($hasFido2)      { "FIDO2 Key" }
                             elseif ($hasPasskey) { "Passkey" }
                             elseif ($hasWhfb)    { "WHfB" }
                             elseif ($hasAuthApp) { "Authenticator App" }
                             elseif ($hasTotp)    { "TOTP" }
                             elseif ($hasPhone)   { "Phone" }
                             elseif ($hasEmail)   { "Email OTP" }
                             else                 { "Password Only" }
        }
        catch {
            $defaultMethod = "Error reading methods"
        }

        # Last sign-in — only available on tenants with Entra ID P1/P2
        $lastSignIn = if (-not $hasPremiumLicense) { "N/A (no P1/P2 license)" } else { "Never / Unknown" }
        if ($hasPremiumLicense -and $user.SignInActivity -and $user.SignInActivity.LastSignInDateTime) {
            $lastSignIn = "{0:yyyy-MM-dd HH:mm}" -f $user.SignInActivity.LastSignInDateTime
        }

        # MFA status label
        $mfaStatus = if ($hasMfa) { "Registered" } else { "NOT Registered" }
        $isAdmin   = $AdminIds.Contains($user.Id)

        $report.Add([PSCustomObject]@{
            DisplayName     = $user.DisplayName
            UPN             = $user.UserPrincipalName
            UserType        = $user.UserType
            IsAdmin         = $isAdmin
            MfaStatus       = $mfaStatus
            MfaRegistered   = $hasMfa
            DefaultMethod   = $defaultMethod
            MethodCount     = $methodCount
            Methods         = ($methodNames | Select-Object -Unique) -join "; "
            HasAuthApp      = $hasAuthApp
            HasFIDO2        = $hasFido2
            HasPhone        = $hasPhone
            HasTOTP         = $hasTotp
            HasEmailOTP     = $hasEmail
            HasWHfB         = $hasWhfb
            HasPasskey      = $hasPasskey
            LastSignIn      = $lastSignIn
        })
    }

    Write-Host "  [+] Data collection complete. $($report.Count) records." -ForegroundColor Green
    return $report
}

# ─────────────────────────────────────────────
# SECTION 4 — Print console summary
# ─────────────────────────────────────────────

function Show-ConsoleSummary {
    param ($Report)

    Write-SectionHeader "MFA SUMMARY"

    $total        = $Report.Count
    $registered   = @($Report | Where-Object { $_.MfaRegistered }).Count
    $notReg       = $total - $registered
    $admins       = @($Report | Where-Object { $_.IsAdmin }).Count
    $adminNoMfa   = @($Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
    $pct          = if ($total -gt 0) { [math]::Round(($registered / $total) * 100, 1) } else { 0 }

    Write-Host ""
    Write-Host ("  Total Users         : {0,6}" -f $total)       -ForegroundColor Cyan
    Write-Host ("  MFA Registered      : {0,6}  ({1}%)" -f $registered, $pct) `
        -ForegroundColor $(if ($pct -ge 90) { "Green" } elseif ($pct -ge 70) { "Yellow" } else { "Red" })
    Write-Host ("  MFA NOT Registered  : {0,6}" -f $notReg) `
        -ForegroundColor $(if ($notReg -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host ("  Admin Accounts      : {0,6}" -f $admins)       -ForegroundColor Cyan
    Write-Host ("  Admins without MFA  : {0,6}" -f $adminNoMfa) `
        -ForegroundColor $(if ($adminNoMfa -gt 0) { "Red" } else { "Green" })
    Write-Host ""

    Write-Host "  Method breakdown:" -ForegroundColor Yellow
    Write-Host ("    Authenticator App        : {0}" -f @($Report | Where-Object { $_.HasAuthApp  }).Count) -ForegroundColor Cyan
    Write-Host ("    FIDO2 Security Key        : {0}" -f @($Report | Where-Object { $_.HasFIDO2   }).Count) -ForegroundColor Cyan
    Write-Host ("    Phone (SMS/Voice)         : {0}" -f @($Report | Where-Object { $_.HasPhone   }).Count) -ForegroundColor Cyan
    Write-Host ("    Software TOTP (OATH)      : {0}" -f @($Report | Where-Object { $_.HasTOTP    }).Count) -ForegroundColor Cyan
    Write-Host ("    Email OTP                 : {0}" -f @($Report | Where-Object { $_.HasEmailOTP}).Count) -ForegroundColor Cyan
    Write-Host ("    Windows Hello for Business: {0}" -f @($Report | Where-Object { $_.HasWHfB    }).Count) -ForegroundColor Cyan
    Write-Host ("    Passkey                   : {0}" -f @($Report | Where-Object { $_.HasPasskey }).Count) -ForegroundColor Cyan
    Write-Host ""

    if ($adminNoMfa -gt 0) {
        Write-Host "  [!] CRITICAL — Admin accounts without MFA:" -ForegroundColor Red
        $Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered } |
            Select-Object DisplayName, UPN, LastSignIn |
            Format-Table -AutoSize | Out-Host
    }
}

# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────

function Export-MfaXlsx {
    param ([string]$Path, $Report, [string]$TenantDomain, [string]$ReportDate)

    Write-Host "[*] Exporting Excel workbook..." -ForegroundColor Cyan

    $total      = $Report.Count
    $registered = @($Report | Where-Object { $_.MfaRegistered }).Count
    $notReg     = $total - $registered
    $pct        = if ($total -gt 0) { [math]::Round(($registered / $total) * 100, 1) } else { 0 }
    $admins     = @($Report | Where-Object { $_.IsAdmin }).Count
    $adminNoMfa = @($Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count

    if (-not $script:ImportExcelAutoSizeSupported) {
        Write-Host "[*] ImportExcel autosize disabled on this platform. Exporting workbook without column autosizing." -ForegroundColor Yellow
    }

    $xl = @{
        Path         = $Path
        AutoSize     = $script:ImportExcelAutoSizeSupported
        FreezeTopRow = $true
        BoldTopRow   = $true
        Append       = $true
    }

    @(
        [PSCustomObject]@{ Category = "Report Generated";     Value = $ReportDate }
        [PSCustomObject]@{ Category = "Tenant Domain";        Value = $TenantDomain }
        [PSCustomObject]@{ Category = "";                     Value = "" }
        [PSCustomObject]@{ Category = "Total Users";          Value = $total }
        [PSCustomObject]@{ Category = "MFA Registered";       Value = "$registered ($pct%)" }
        [PSCustomObject]@{ Category = "MFA NOT Registered";   Value = $notReg }
        [PSCustomObject]@{ Category = "";                     Value = "" }
        [PSCustomObject]@{ Category = "Admin Accounts";       Value = $admins }
        [PSCustomObject]@{ Category = "Admins without MFA";   Value = $adminNoMfa }
        [PSCustomObject]@{ Category = "";                     Value = "" }
        [PSCustomObject]@{ Category = "Has Authenticator App"; Value = @($Report | Where-Object { $_.HasAuthApp   }).Count }
        [PSCustomObject]@{ Category = "Has FIDO2 Key";        Value = @($Report | Where-Object { $_.HasFIDO2     }).Count }
        [PSCustomObject]@{ Category = "Has Phone";            Value = @($Report | Where-Object { $_.HasPhone     }).Count }
        [PSCustomObject]@{ Category = "Has TOTP";             Value = @($Report | Where-Object { $_.HasTOTP      }).Count }
        [PSCustomObject]@{ Category = "Has Email OTP";        Value = @($Report | Where-Object { $_.HasEmailOTP  }).Count }
        [PSCustomObject]@{ Category = "Has WHfB";             Value = @($Report | Where-Object { $_.HasWHfB      }).Count }
        [PSCustomObject]@{ Category = "Has Passkey";          Value = @($Report | Where-Object { $_.HasPasskey   }).Count }
    ) | Export-Excel @xl -WorksheetName "Summary"

    $Report | Select-Object DisplayName,UPN,UserType,IsAdmin,MfaStatus,DefaultMethod,MethodCount,Methods,HasAuthApp,HasFIDO2,HasPhone,HasTOTP,HasEmailOTP,HasWHfB,HasPasskey,LastSignIn |
        Export-Excel @xl -WorksheetName "All Users"

    $noMfa = $Report | Where-Object { -not $_.MfaRegistered }
    if (@($noMfa).Count -gt 0) {
        $noMfa | Select-Object DisplayName,UPN,UserType,IsAdmin,LastSignIn |
            Export-Excel @xl -WorksheetName "No MFA"
    } else {
        [PSCustomObject]@{ Info = "All users have MFA registered." } | Export-Excel @xl -WorksheetName "No MFA"
    }

    $adminsNoMfa = $Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }
    if (@($adminsNoMfa).Count -gt 0) {
        $adminsNoMfa | Select-Object DisplayName,UPN,DefaultMethod,LastSignIn |
            Export-Excel @xl -WorksheetName "Admins No MFA"
    } else {
        [PSCustomObject]@{ Info = "All admin accounts have MFA registered." } | Export-Excel @xl -WorksheetName "Admins No MFA"
    }

    Write-Host "[+] Excel exported to: $Path" -ForegroundColor Green
}

# ─────────────────────────────────────────────
# HTML DASHBOARD EXPORT
# ─────────────────────────────────────────────

function Export-MfaHtmlReport {
    param ([string]$Path, $Report, [string]$TenantDomain, [string]$ReportDate)

    Write-Host "[*] Building HTML dashboard..." -ForegroundColor Cyan

    $total        = $Report.Count
    $registered   = @($Report | Where-Object { $_.MfaRegistered }).Count
    $notReg       = $total - $registered
    $pct          = if ($total -gt 0) { [math]::Round(($registered / $total) * 100, 1) } else { 0 }
    $admins       = @($Report | Where-Object { $_.IsAdmin }).Count
    $adminNoMfa   = @($Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
    $cntAuthApp   = @($Report | Where-Object { $_.HasAuthApp   }).Count
    $cntFido2     = @($Report | Where-Object { $_.HasFIDO2     }).Count
    $cntPhone     = @($Report | Where-Object { $_.HasPhone     }).Count
    $cntTotp      = @($Report | Where-Object { $_.HasTOTP      }).Count
    $cntEmail     = @($Report | Where-Object { $_.HasEmailOTP  }).Count
    $cntWhfb      = @($Report | Where-Object { $_.HasWHfB      }).Count
    $cntPasskey   = @($Report | Where-Object { $_.HasPasskey   }).Count

    # JSON helpers
    function jStr([string]$s) { '"' + ($s -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' ') + '"' }
    function jBool([bool]$b)  { $b.ToString().ToLower() }

    $userCols = @('DisplayName','UPN','UserType','IsAdmin','MfaStatus','DefaultMethod','MethodCount','Methods','HasAuthApp','HasFIDO2','HasPhone','HasTOTP','HasEmailOTP','HasWHfB','HasPasskey','LastSignIn')

    $userJsonParts = foreach ($r in $Report) {
        $pairs = foreach ($c in $userCols) {
            $v = $r.$c
            if ($v -eq $null) { $v = "" }
            if ($v -is [bool]) { (jStr $c) + ":" + (jBool $v) }
            else               { (jStr $c) + ":" + (jStr ([string]$v)) }
        }
        "{" + ($pairs -join ",") + "}"
    }
    $usersJson = "[" + ($userJsonParts -join ",") + "]"

    $htmlPart1 = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 MFA Report — $TenantDomain</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg:#080c14; --bg2:#0c1220; --bg3:#101828; --bg4:#141e2e;
    --border:#1a2840; --border2:#223050;
    --text:#c8d8f0; --text2:#6a88b0; --text3:#3a5070;
    --accent:#38bdf8; --accent2:#0ea5e9;
    --ok:#10b981; --ok-bg:rgba(16,185,129,.1);
    --warn:#f59e0b; --warn-bg:rgba(245,158,11,.1);
    --crit:#f43f5e; --crit-bg:rgba(244,63,94,.1);
    --err:#fb923c; --err-bg:rgba(251,146,60,.1);
    --mono:'JetBrains Mono',monospace; --sans:'DM Sans',sans-serif;
    --r:8px; --r2:12px;
  }
  html { scroll-behavior:smooth; }
  body { font-family:var(--sans); background:var(--bg); color:var(--text); min-height:100vh; line-height:1.6; }
  body::before {
    content:''; position:fixed; inset:0;
    background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='0.04'/%3E%3C/svg%3E");
    pointer-events:none; z-index:0; opacity:.4;
  }
  .topbar {
    position:sticky; top:0; z-index:100;
    background:rgba(8,12,20,.9); backdrop-filter:blur(16px);
    border-bottom:1px solid var(--border);
    padding:0 2rem; height:52px; display:flex; align-items:center; gap:1.5rem;
  }
  .topbar-logo { font-family:var(--mono); font-size:.72rem; font-weight:700; letter-spacing:.16em; text-transform:uppercase; color:var(--accent); }
  .topbar-sep  { color:var(--border2); }
  .topbar-server { font-family:var(--mono); font-size:.78rem; color:var(--text2); }
  .topbar-nav { display:flex; gap:.15rem; margin-left:.75rem; }
  .topbar-nav a { font-size:.72rem; font-weight:600; color:var(--text3); text-decoration:none; padding:.3rem .65rem; border-radius:var(--r); transition:all .15s; }
  .topbar-nav a:hover { background:var(--bg3); color:var(--text); }
  .topbar-right { margin-left:auto; display:flex; align-items:center; gap:1rem; }
  .topbar-time { font-family:var(--mono); font-size:.7rem; color:var(--text3); }
  .page { max-width:1440px; margin:0 auto; padding:1.75rem 2rem; position:relative; z-index:1; }
  .server-strip {
    background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2);
    padding:1rem 1.5rem; display:flex; flex-wrap:wrap; gap:2.5rem; margin-bottom:1.5rem;
  }
  .strip-item { display:flex; flex-direction:column; gap:.2rem; }
  .strip-label { font-size:.63rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:var(--text3); }
  .strip-value { font-family:var(--mono); font-size:.82rem; color:var(--accent); }
  .hero { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:.85rem; margin-bottom:1.5rem; }
  .kpi {
    background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2);
    padding:1rem 1.1rem; position:relative; overflow:hidden;
    transition:transform .2s,border-color .2s; cursor:default; animation:fadeIn .3s ease both;
  }
  .kpi:hover { transform:translateY(-2px); border-color:var(--border2); }
  .kpi::after { content:''; position:absolute; top:0;left:0;right:0; height:2px; border-radius:2px 2px 0 0; }
  .kpi.ok::after    { background:var(--ok); }
  .kpi.warn::after  { background:var(--warn); }
  .kpi.crit::after  { background:var(--crit); }
  .kpi.neutral::after { background:var(--accent2); }
  .kpi-label { font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); margin-bottom:.4rem; }
  .kpi-value { font-family:var(--mono); font-size:1.7rem; font-weight:700; line-height:1; }
  .kpi.ok .kpi-value     { color:var(--ok); }
  .kpi.warn .kpi-value   { color:var(--warn); }
  .kpi.crit .kpi-value   { color:var(--crit); }
  .kpi.neutral .kpi-value { color:var(--accent); }
  .kpi-sub { font-size:.68rem; color:var(--text3); margin-top:.35rem; font-family:var(--mono); }
  .sections { display:grid; gap:1.25rem; }
  .card { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2); overflow:hidden; animation:fadeIn .35s ease both; }
  .card:nth-child(1){animation-delay:.05s} .card:nth-child(2){animation-delay:.10s}
  .card:nth-child(3){animation-delay:.15s} .card:nth-child(4){animation-delay:.20s}
  .card-header { display:flex; align-items:center; gap:.75rem; padding:.85rem 1.25rem; border-bottom:1px solid var(--border); cursor:pointer; user-select:none; transition:background .15s; }
  .card-header:hover { background:var(--bg3); }
  .card-icon  { font-size:.95rem; width:1.4rem; text-align:center; opacity:.7; }
  .card-title { font-size:.72rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text2); flex:1; }
  .card-badge { font-family:var(--mono); font-size:.7rem; padding:.18rem .55rem; border-radius:4px; }
  .badge-ok      { background:var(--ok-bg);   color:var(--ok); }
  .badge-warn    { background:var(--warn-bg);  color:var(--warn); }
  .badge-crit    { background:var(--crit-bg);  color:var(--crit); }
  .badge-neutral { background:rgba(56,189,248,.1); color:var(--accent); }
  .card-chevron { font-size:.65rem; color:var(--text3); transition:transform .2s; }
  .card.collapsed .card-chevron { transform:rotate(-90deg); }
  .card-body { padding:1.25rem; }
  .card.collapsed .card-body { display:none; }

  /* Search / filter */
  .toolbar { display:flex; gap:.65rem; margin-bottom:1rem; flex-wrap:wrap; align-items:center; }
  .search-box {
    flex:1; min-width:220px; max-width:360px;
    background:var(--bg3); border:1px solid var(--border2); border-radius:var(--r);
    color:var(--text); font-family:var(--mono); font-size:.78rem;
    padding:.4rem .75rem; outline:none; transition:border-color .15s;
  }
  .search-box:focus { border-color:var(--accent2); }
  .search-box::placeholder { color:var(--text3); }
  .filter-btn {
    font-family:var(--sans); font-size:.72rem; font-weight:700; padding:.3rem .8rem;
    border-radius:5px; border:1px solid var(--border); background:var(--bg3);
    color:var(--text3); cursor:pointer; transition:all .15s;
  }
  .filter-btn:hover { color:var(--text); border-color:var(--border2); }
  .filter-btn.active-all  { border-color:var(--accent2); color:var(--accent); background:rgba(56,189,248,.08); }
  .filter-btn.active-crit { border-color:var(--crit); color:var(--crit); background:var(--crit-bg); }
  .filter-btn.active-ok   { border-color:var(--ok);   color:var(--ok);   background:var(--ok-bg); }
  .filter-btn.active-admin{ border-color:var(--warn);  color:var(--warn); background:var(--warn-bg); }

  /* Tables */
  .table-wrap { overflow-x:auto; border:1px solid var(--border); border-radius:var(--r); }
  .table-scroll { max-height:560px; overflow-y:auto; border:1px solid var(--border); border-radius:var(--r); }
  table { width:100%; border-collapse:collapse; font-size:.77rem; }
  thead { background:var(--bg3); position:sticky; top:0; z-index:1; }
  th { padding:.55rem .9rem; text-align:left; font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); white-space:nowrap; border-bottom:1px solid var(--border); }
  td { padding:.5rem .9rem; border-bottom:1px solid var(--border); color:var(--text); font-family:var(--mono); font-size:.76rem; vertical-align:top; max-width:340px; word-break:break-all; }
  tr:last-child td { border-bottom:none; }
  tbody tr:hover td { background:rgba(255,255,255,.02); }
  .pill { display:inline-block; padding:.1rem .45rem; border-radius:4px; font-size:.7rem; font-weight:600; white-space:nowrap; }
  .hidden-row { display:none; }

  /* Method badge grid */
  .method-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(190px,1fr)); gap:.85rem; }
  .method-item {
    background:var(--bg3); border:1px solid var(--border); border-radius:var(--r);
    padding:.85rem 1rem; display:flex; align-items:center; gap:.85rem;
  }
  .method-icon { font-size:1.25rem; opacity:.7; width:2rem; text-align:center; flex-shrink:0; }
  .method-info { display:flex; flex-direction:column; gap:.2rem; }
  .method-label { font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); }
  .method-count { font-family:var(--mono); font-size:1.05rem; font-weight:700; color:var(--accent); }

  /* Progress bar for coverage */
  .coverage-wrap { margin-bottom:1.25rem; }
  .coverage-label { display:flex; justify-content:space-between; margin-bottom:.4rem; }
  .coverage-label-text { font-size:.72rem; color:var(--text2); font-family:var(--mono); }
  .coverage-pct { font-family:var(--mono); font-size:.72rem; }
  .coverage-bar-bg { height:8px; background:var(--border); border-radius:4px; overflow:hidden; }
  .coverage-bar { height:100%; border-radius:4px; transition:width 1s cubic-bezier(.25,.46,.45,.94); }
  .coverage-bar.ok   { background:linear-gradient(90deg,var(--ok),#34d399); }
  .coverage-bar.warn { background:linear-gradient(90deg,var(--warn),#fcd34d); }
  .coverage-bar.crit { background:linear-gradient(90deg,var(--crit),#fb7185); }

  .empty { font-family:var(--mono); font-size:.78rem; color:var(--text3); padding:.75rem 0; font-style:italic; }
  .section-sub { font-size:.7rem; color:var(--text3); margin-bottom:.75rem; font-family:var(--mono); font-weight:600; letter-spacing:.08em; text-transform:uppercase; }

  @keyframes fadeIn { from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)} }
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-track{background:var(--bg2)}
  ::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
  ::-webkit-scrollbar-thumb:hover{background:var(--text3)}
  @media(max-width:640px){.topbar-nav{display:none}.hero{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-logo">&#x2B21; RegITs GmbH M365 MFA Monitor</div>
  <span class="topbar-sep">|</span>
  <div class="topbar-server">$TenantDomain</div>
  <nav class="topbar-nav">
    <a href="#sec-overview">Overview</a>
    <a href="#sec-methods">Methods</a>
    <a href="#sec-users">All Users</a>
    <a href="#sec-nomfa">No MFA</a>
    <a href="#sec-admins">Admins</a>
  </nav>
  <div class="topbar-right">
    <div class="topbar-time">Generated: $ReportDate</div>
  </div>
</div>

<div class="page">
  <div id="server-strip" class="server-strip"></div>
  <div id="hero" class="hero"></div>
  <div id="sections" class="sections"></div>
</div>

<script>
const DATA = {
  tenant:     "$TenantDomain",
  reportDate: "$ReportDate",
  total:      $total,
  registered: $registered,
  notReg:     $notReg,
  pct:        $pct,
  admins:     $admins,
  adminNoMfa: $adminNoMfa,
  cntAuthApp: $cntAuthApp,
  cntFido2:   $cntFido2,
  cntPhone:   $cntPhone,
  cntTotp:    $cntTotp,
  cntEmail:   $cntEmail,
  cntWhfb:    $cntWhfb,
  cntPasskey: $cntPasskey,
  users:      $usersJson,
};
"@

    $htmlPart2 = @'
function esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function kpi(value,label,sub,cls){return `<div class="kpi ${cls}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div></div>`;}
function strip(items){return items.map(([l,v])=>`<div class="strip-item"><span class="strip-label">${l}</span><span class="strip-value">${esc(v)}</span></div>`).join('');}
function card(id,icon,title,badgeText,badgeClass,body){return `<div class="card" id="${id}"><div class="card-header" onclick="toggleCard(this)"><span class="card-icon">${icon}</span><span class="card-title">${title}</span><span class="card-badge ${badgeClass}">${badgeText}</span><span class="card-chevron">&#x25BC;</span></div><div class="card-body">${body}</div></div>`;}
function toggleCard(h){h.closest('.card').classList.toggle('collapsed');}
function methodPill(b){return b?'<span class="pill badge-ok">Yes</span>':'<span class="pill" style="background:rgba(244,63,94,.08);color:#f43f5e">No</span>';}

function mfaStatusPill(v){
  if(v==='Registered') return `<span class="pill badge-ok">${esc(v)}</span>`;
  return `<span class="pill badge-crit">${esc(v)}</span>`;
}
function adminPill(v){
  if(v===true||v==='true'||v==='True') return `<span class="pill badge-warn">Admin</span>`;
  return `<span class="pill" style="background:rgba(56,189,248,.07);color:var(--accent)">User</span>`;
}
function methodLabel(v){
  const map={'Authenticator App':'ok','FIDO2 Key':'ok','Passkey':'ok','WHfB':'ok','TOTP':'ok','Phone':'warn','Email OTP':'warn','Password Only':'crit','None':'crit'};
  const cls=map[v]||'neutral';
  return `<span class="pill badge-${cls}">${esc(v)}</span>`;
}

function buildUserTable(rows){
  if(!rows||rows.length===0)return '<p class="empty">No users match this filter.</p>';
  const cols=['DisplayName','UPN','MfaStatus','IsAdmin','DefaultMethod','MethodCount','Methods','LastSignIn'];
  let h='<table><thead><tr>';
  for(const c of cols)h+=`<th>${esc(c)}</th>`;
  h+='</tr></thead><tbody>';
  for(const r of rows){
    h+=`<tr>`;
    h+=`<td>${esc(r.DisplayName)}</td>`;
    h+=`<td>${esc(r.UPN)}</td>`;
    h+=`<td>${mfaStatusPill(r.MfaStatus)}</td>`;
    h+=`<td>${adminPill(r.IsAdmin)}</td>`;
    h+=`<td>${methodLabel(r.DefaultMethod)}</td>`;
    h+=`<td>${esc(r.MethodCount)}</td>`;
    h+=`<td style="font-size:.7rem;color:var(--text2)">${esc(r.Methods||'—')}</td>`;
    h+=`<td>${esc(r.LastSignIn)}</td>`;
    h+='</tr>';
  }
  h+='</tbody></table>';
  return h;
}

// ── Live search + filter for user table ──
let _activeFilter='all';
let _searchTerm='';
function applyUserFilter(){
  const rows=document.querySelectorAll('#user-tbody tr');
  rows.forEach(row=>{
    const cells=Array.from(row.cells).map(c=>c.textContent.toLowerCase()).join(' ');
    const matchSearch=!_searchTerm||cells.includes(_searchTerm);
    const mfaCell=row.cells[2]?row.cells[2].textContent.toLowerCase():'';
    const adminCell=row.cells[3]?row.cells[3].textContent.toLowerCase():'';
    let matchFilter=true;
    if(_activeFilter==='nomfa') matchFilter=mfaCell.includes('not');
    if(_activeFilter==='mfa')   matchFilter=!mfaCell.includes('not');
    if(_activeFilter==='admin') matchFilter=adminCell.includes('admin');
    row.style.display=(matchSearch&&matchFilter)?'':'none';
  });
}
function setFilter(f,btn){
  _activeFilter=f;
  document.querySelectorAll('.uf-btn').forEach(b=>{
    b.classList.remove('active-all','active-crit','active-ok','active-admin');
  });
  if(f==='all')   btn.classList.add('active-all');
  if(f==='nomfa') btn.classList.add('active-crit');
  if(f==='mfa')   btn.classList.add('active-ok');
  if(f==='admin') btn.classList.add('active-admin');
  applyUserFilter();
}

function render(){
  const d=DATA;
  const pct=d.pct;
  const coverageClass=pct>=90?'ok':pct>=70?'warn':'crit';
  const coveragePctColor=pct>=90?'var(--ok)':pct>=70?'var(--warn)':'var(--crit)';

  // Strip
  document.getElementById('server-strip').innerHTML=strip([
    ['Tenant',d.tenant],
    ['Total Users',String(d.total)],
    ['MFA Coverage',pct+'%'],
    ['Admins',String(d.admins)],
    ['Admins Without MFA',String(d.adminNoMfa)],
    ['Report Generated',d.reportDate],
  ]);

  // KPIs
  const mfaPctCls=pct>=90?'ok':pct>=70?'warn':'crit';
  document.getElementById('hero').innerHTML=[
    kpi(d.total,'Total<br>Users','active members','neutral'),
    kpi(d.registered,'MFA<br>Registered','have 2nd factor',d.registered===d.total?'ok':'warn'),
    kpi(d.notReg,'MFA<br>NOT Registered','password only',d.notReg>0?'crit':'ok'),
    kpi(pct+'%','MFA<br>Coverage','of all users',mfaPctCls),
    kpi(d.admins,'Admin<br>Accounts','privileged roles','neutral'),
    kpi(d.adminNoMfa,'Admins<br>Without MFA','critical risk',d.adminNoMfa>0?'crit':'ok'),
    kpi(d.cntAuthApp,'Authenticator<br>App','microsoft/3rd party','neutral'),
    kpi(d.cntFido2,'FIDO2<br>Keys','hardware tokens',d.cntFido2>0?'ok':'neutral'),
    kpi(d.cntPhone,'Phone<br>SMS/Voice','lower assurance method',d.cntPhone>0?'warn':'neutral'),
  ].join('');

  // Sections
  const sec=document.getElementById('sections');
  sec.innerHTML='';

  // ── Overview card ──
  const overviewBody=`
    <div class="coverage-wrap">
      <div class="coverage-label">
        <span class="coverage-label-text">MFA Coverage — ${d.registered} of ${d.total} users</span>
        <span class="coverage-pct" style="color:${coveragePctColor}">${pct}%</span>
      </div>
      <div class="coverage-bar-bg"><div class="coverage-bar ${coverageClass}" style="width:${pct}%"></div></div>
    </div>
    ${d.adminNoMfa>0?`<div style="background:var(--crit-bg);border:1px solid var(--crit);border-radius:var(--r);padding:.75rem 1rem;margin-bottom:1rem;font-size:.78rem;font-family:var(--mono);color:var(--crit)"><i class="fa-solid fa-triangle-exclamation"></i>&nbsp; <strong>${d.adminNoMfa} admin account(s)</strong> have no MFA registered — immediate action required.</div>`:''}
    <div class="method-grid">
      <div class="method-item"><div class="method-icon" style="color:#38bdf8"><i class="fa-solid fa-mobile-screen-button"></i></div><div class="method-info"><div class="method-label">Authenticator App</div><div class="method-count">${d.cntAuthApp}</div></div></div>
      <div class="method-item"><div class="method-icon" style="color:#a78bfa"><i class="fa-solid fa-key"></i></div><div class="method-info"><div class="method-label">FIDO2 Security Key</div><div class="method-count">${d.cntFido2}</div></div></div>
      <div class="method-item"><div class="method-icon" style="color:#f59e0b"><i class="fa-solid fa-phone"></i></div><div class="method-info"><div class="method-label">Phone (SMS/Voice)</div><div class="method-count">${d.cntPhone}</div></div></div>
      <div class="method-item"><div class="method-icon" style="color:#10b981"><i class="fa-solid fa-clock-rotate-left"></i></div><div class="method-info"><div class="method-label">TOTP / OATH</div><div class="method-count">${d.cntTotp}</div></div></div>
      <div class="method-item"><div class="method-icon" style="color:#6a88b0"><i class="fa-solid fa-envelope"></i></div><div class="method-info"><div class="method-label">Email OTP</div><div class="method-count">${d.cntEmail}</div></div></div>
      <div class="method-item"><div class="method-icon" style="color:#38bdf8"><i class="fa-brands fa-windows"></i></div><div class="method-info"><div class="method-label">Windows Hello (WHfB)</div><div class="method-count">${d.cntWhfb}</div></div></div>
      <div class="method-item"><div class="method-icon" style="color:#fb923c"><i class="fa-solid fa-fingerprint"></i></div><div class="method-info"><div class="method-label">Passkey</div><div class="method-count">${d.cntPasskey}</div></div></div>
    </div>`;
  const overviewBadge=`${pct}% covered`;
  const overviewBadgeClass=pct>=90?'badge-ok':pct>=70?'badge-warn':'badge-crit';
  sec.innerHTML+=card('sec-overview','<i class="fa-solid fa-chart-pie"></i>','MFA Coverage Overview',overviewBadge,overviewBadgeClass,overviewBody);

  // ── Methods details card ──
  const methodCols=['DisplayName','UPN','DefaultMethod','HasAuthApp','HasFIDO2','HasPhone','HasTOTP','HasEmailOTP','HasWHfB','HasPasskey','LastSignIn'];
  let mTable='<div class="table-scroll"><table><thead><tr>';
  for(const c of methodCols)mTable+=`<th>${esc(c.replace(/Has/,''))}</th>`;
  mTable+='</tr></thead><tbody>';
  for(const r of d.users){
    mTable+=`<tr>
      <td>${esc(r.DisplayName)}</td>
      <td style="font-size:.72rem">${esc(r.UPN)}</td>
      <td>${methodLabel(r.DefaultMethod)}</td>
      <td>${methodPill(r.HasAuthApp)}</td>
      <td>${methodPill(r.HasFIDO2)}</td>
      <td>${methodPill(r.HasPhone)}</td>
      <td>${methodPill(r.HasTOTP)}</td>
      <td>${methodPill(r.HasEmailOTP)}</td>
      <td>${methodPill(r.HasWHfB)}</td>
      <td>${methodPill(r.HasPasskey)}</td>
      <td>${esc(r.LastSignIn)}</td>
    </tr>`;
  }
  mTable+='</tbody></table></div>';
  sec.innerHTML+=card('sec-methods','<i class="fa-solid fa-shield-halved"></i>','Authentication Methods per User',`${d.total} users`,'badge-neutral',mTable);

  // ── All users card (with live filter) ──
  let uTable=`
    <div class="toolbar">
      <input class="search-box" type="text" placeholder="Search name or UPN..." oninput="_searchTerm=this.value.toLowerCase();applyUserFilter()">
      <button class="filter-btn uf-btn active-all" onclick="setFilter('all',this)">All (${d.total})</button>
      <button class="filter-btn uf-btn" onclick="setFilter('mfa',this)">MFA Registered (${d.registered})</button>
      <button class="filter-btn uf-btn" onclick="setFilter('nomfa',this)">No MFA (${d.notReg})</button>
      <button class="filter-btn uf-btn" onclick="setFilter('admin',this)">Admins (${d.admins})</button>
    </div>
    <div class="table-scroll"><table><thead><tr>
      <th>Display Name</th><th>UPN</th><th>MFA Status</th><th>Role</th>
      <th>Default Method</th><th># Methods</th><th>All Methods</th><th>Last Sign-In</th>
    </tr></thead><tbody id="user-tbody">`;
  for(const r of d.users){
    uTable+=`<tr>
      <td>${esc(r.DisplayName)}</td>
      <td style="font-size:.72rem">${esc(r.UPN)}</td>
      <td>${mfaStatusPill(r.MfaStatus)}</td>
      <td>${adminPill(r.IsAdmin)}</td>
      <td>${methodLabel(r.DefaultMethod)}</td>
      <td>${esc(r.MethodCount)}</td>
      <td style="font-size:.7rem;color:var(--text2)">${esc(r.Methods||'—')}</td>
      <td>${esc(r.LastSignIn)}</td>
    </tr>`;
  }
  uTable+='</tbody></table></div>';
  sec.innerHTML+=card('sec-users','<i class="fa-solid fa-users"></i>','All Users',`${d.total} users`,'badge-neutral',uTable);

  // ── No MFA card ──
  const noMfaUsers=d.users.filter(u=>u.MfaStatus!=='Registered');
  const noMfaBadge=noMfaUsers.length===0?'All users protected':`${noMfaUsers.length} user(s) at risk`;
  const noMfaBadgeClass=noMfaUsers.length===0?'badge-ok':'badge-crit';
  let noMfaBody='';
  if(noMfaUsers.length===0){
    noMfaBody='<p class="empty"><i class="fa-solid fa-circle-check" style="color:var(--ok)"></i>&nbsp; Excellent! Every user has at least one MFA method registered.</p>';
  } else {
    noMfaBody='<div class="table-scroll"><table><thead><tr><th>Display Name</th><th>UPN</th><th>Role</th><th>Last Sign-In</th></tr></thead><tbody>';
    for(const r of noMfaUsers){
      noMfaBody+=`<tr>
        <td>${esc(r.DisplayName)}</td>
        <td style="font-size:.72rem">${esc(r.UPN)}</td>
        <td>${adminPill(r.IsAdmin)}</td>
        <td>${esc(r.LastSignIn)}</td>
      </tr>`;
    }
    noMfaBody+='</tbody></table></div>';
  }
  sec.innerHTML+=card('sec-nomfa','<i class="fa-solid fa-triangle-exclamation"></i>','Users Without MFA',noMfaBadge,noMfaBadgeClass,noMfaBody);

  // ── Admins card ──
  const adminUsers=d.users.filter(u=>u.IsAdmin===true||u.IsAdmin==='true'||u.IsAdmin==='True');
  const adminBadge=`${d.adminNoMfa} without MFA`;
  const adminBadgeClass=d.adminNoMfa>0?'badge-crit':d.admins>0?'badge-ok':'badge-neutral';
  let adminBody='';
  if(adminUsers.length===0){
    adminBody='<p class="empty">No privileged admin accounts found in the scanned roles.</p>';
  } else {
    adminBody='<div class="table-scroll"><table><thead><tr><th>Display Name</th><th>UPN</th><th>MFA Status</th><th>Default Method</th><th>All Methods</th><th>Last Sign-In</th></tr></thead><tbody>';
    for(const r of adminUsers){
      adminBody+=`<tr>
        <td>${esc(r.DisplayName)}</td>
        <td style="font-size:.72rem">${esc(r.UPN)}</td>
        <td>${mfaStatusPill(r.MfaStatus)}</td>
        <td>${methodLabel(r.DefaultMethod)}</td>
        <td style="font-size:.7rem;color:var(--text2)">${esc(r.Methods||'—')}</td>
        <td>${esc(r.LastSignIn)}</td>
      </tr>`;
    }
    adminBody+='</tbody></table></div>';
  }
  sec.innerHTML+=card('sec-admins','<i class="fa-solid fa-user-shield"></i>','Privileged Admin Accounts',adminBadge,adminBadgeClass,adminBody);
}

render();
'@

    $htmlClose = @'
</script>
</body>
</html>
'@

    $html = $htmlPart1 + $htmlPart2 + $htmlClose

    try {
        $html | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "[+] HTML dashboard exported to: $Path" -ForegroundColor Green
    }
    catch { Write-Error "Failed to export HTML report: $_" }
}

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

Write-Host "`n====================================================" -ForegroundColor Cyan
Write-Host "       M365 MFA AUTHENTICATION REPORT" -ForegroundColor Cyan
Write-Host "       $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "====================================================`n" -ForegroundColor Cyan

$xlAvailable = Assert-RequiredModules -NeedExcelExport ([bool]$ExportXlsx)
Connect-ToGraph | Out-Null

$adminIds   = Get-AdminUserIds -RoleNames $AdminRoles
$mfaReport  = Get-MfaReport -IncludeGuests $IncludeGuests.IsPresent -AdminIds $adminIds
Show-ConsoleSummary -Report $mfaReport

$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$xlsxPath   = $null
$htmlPath   = $null

if ($ExportXlsx) {
    if (-not $xlAvailable) {
        Write-Warning "[!] ImportExcel not available. Skipping .xlsx export."
    } else {
        $xlsxPath = Add-TimestampToPath -Path $ExportXlsx -DefaultExtension ".xlsx"
        try {
            if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
            Export-MfaXlsx -Path $xlsxPath -Report $mfaReport `
                -TenantDomain $script:TenantDomain -ReportDate $reportDate
        }
        catch { Write-Error "Excel export failed: $_" }
    }
}

if ($ExportHtml) {
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -DefaultExtension ".html"
    Export-MfaHtmlReport -Path $htmlPath -Report $mfaReport `
        -TenantDomain $script:TenantDomain -ReportDate $reportDate
}

# Disconnect gracefully
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
Write-Host "[*] Disconnected from Microsoft Graph." -ForegroundColor Cyan
Write-Host "[*] MFA report complete.`n" -ForegroundColor Cyan
