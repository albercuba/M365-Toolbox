[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [string]$ExportHtml,

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

function Add-TimestampToPath {
    param(
        [string]$Path,
        [string]$DefaultExtension = ""
    )

    if (-not $Path) { return $Path }

    $stamp = Get-Date -Format "dd.MM.yy-HH.mm.ss"
    $tenant = if ($script:TenantDomain) { $script:TenantDomain.Split('.')[0] } else { "M365" }
    $base = "MFA_${tenant}_$stamp"

    if (Test-Path $Path -PathType Container) {
        return Join-Path $Path "$base$DefaultExtension"
    }

    $dir = [System.IO.Path]::GetDirectoryName($Path)
    $ext = [System.IO.Path]::GetExtension($Path)
    if (-not $ext -and $DefaultExtension) {
        $ext = $DefaultExtension
    }

    if ($dir) {
        return [System.IO.Path]::Combine($dir, "$base$ext")
    }

    return "$base$ext"
}

function Write-SectionHeader {
    param([string]$Title)
    $line = "=" * 50
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host (" " * [Math]::Floor((50 - $Title.Length) / 2) + $Title) -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Assert-RequiredModules {
    $coreModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.SignIns",
        "Microsoft.Graph.DirectoryObjects"
    )
    $optionalRoleModule = "Microsoft.Graph.Identity.DirectoryManagement"

    Write-Host ""
    Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan

    $coreModuleInventory = @{}
    $commonVersions = $null

    foreach ($moduleName in $coreModules) {
        $installed = @(Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending)

        if (-not $installed -or $installed.Count -eq 0) {
            throw "Required module '$moduleName' is not installed."
        }

        $coreModuleInventory[$moduleName] = $installed
        $versionStrings = @($installed | ForEach-Object { $_.Version.ToString() })

        if ($null -eq $commonVersions) {
            $commonVersions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($versionString in $versionStrings) {
                [void]$commonVersions.Add($versionString)
            }
        }
        else {
            $nextCommonVersions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($versionString in $versionStrings) {
                if ($commonVersions.Contains($versionString)) {
                    [void]$nextCommonVersions.Add($versionString)
                }
            }
            $commonVersions = $nextCommonVersions
        }
    }

    if (-not $commonVersions -or $commonVersions.Count -eq 0) {
        throw "No common Microsoft Graph module version is installed across: $($coreModules -join ', ')."
    }

    $selectedGraphVersion = $commonVersions |
        ForEach-Object { [version]$_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    foreach ($moduleName in $coreModules) {
        Import-Module $moduleName -RequiredVersion $selectedGraphVersion -Force -WarningAction SilentlyContinue
        Write-Host "  [+] $moduleName v$selectedGraphVersion" -ForegroundColor Green
    }

    $directoryManagementModule = Get-Module -ListAvailable -Name $optionalRoleModule |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($directoryManagementModule -and ($directoryManagementModule.Version -eq $selectedGraphVersion)) {
        Import-Module $optionalRoleModule -RequiredVersion $selectedGraphVersion -Force -WarningAction SilentlyContinue
        Write-Host "  [+] $optionalRoleModule v$selectedGraphVersion" -ForegroundColor Green
        $script:CanResolveAdminRoles = $true
    }
    else {
        Write-Warning "  [!] $optionalRoleModule v$selectedGraphVersion is not installed. Admin role resolution will be skipped."
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
        return ,$adminIds
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

    return ,$adminIds
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

function Export-MfaHtmlReport {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [object[]]$Report,
        [string]$TenantDomain,
        [string]$ReportDate
    )

    Write-Host "[*] Building HTML dashboard..." -ForegroundColor Cyan

    $total = $Report.Count
    $registered = @($Report | Where-Object { $_.MfaRegistered }).Count
    $notRegistered = $total - $registered
    $admins = @($Report | Where-Object { $_.IsAdmin }).Count
    $adminNoMfa = @($Report | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
    $pct = if ($total -gt 0) { [math]::Round(($registered / $total) * 100, 1) } else { 0 }

    $htmlData = [PSCustomObject]@{
        Tenant       = if ($TenantDomain) { $TenantDomain } else { $TenantId }
        ReportDate   = $ReportDate
        TotalUsers   = $total
        Registered   = $registered
        NotRegistered = $notRegistered
        CoveragePct  = $pct
        Admins       = $admins
        AdminNoMfa   = $adminNoMfa
        MethodCounts = [PSCustomObject]@{
            AuthenticatorApp = @($Report | Where-Object { $_.HasAuthApp }).Count
            Fido2            = @($Report | Where-Object { $_.HasFIDO2 }).Count
            Phone            = @($Report | Where-Object { $_.HasPhone }).Count
            Totp             = @($Report | Where-Object { $_.HasTOTP }).Count
            EmailOtp         = @($Report | Where-Object { $_.HasEmailOTP }).Count
            Whfb             = @($Report | Where-Object { $_.HasWHfB }).Count
            Passkey          = @($Report | Where-Object { $_.HasPasskey }).Count
        }
        Users        = @($Report | Sort-Object IsAdmin -Descending, MfaRegistered, DisplayName)
    } | ConvertTo-Json -Depth 6 -Compress

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>M365 MFA Report</title>
<style>
  :root {
    --bg: #09111f;
    --panel: #101a2b;
    --panel-2: #14233a;
    --line: #223553;
    --text: #e6efff;
    --muted: #93a8c9;
    --ok: #1fb56c;
    --warn: #f0ad2c;
    --crit: #ef5b6b;
    --accent: #4db6ff;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: radial-gradient(circle at top, #13233d 0%, var(--bg) 45%);
    color: var(--text);
    font: 14px/1.5 Segoe UI, Arial, sans-serif;
  }
  .wrap {
    max-width: 1380px;
    margin: 0 auto;
    padding: 28px 20px 40px;
  }
  .hero {
    display: flex;
    justify-content: space-between;
    gap: 16px;
    flex-wrap: wrap;
    margin-bottom: 20px;
  }
  .hero h1 {
    margin: 0;
    font-size: 28px;
  }
  .meta {
    color: var(--muted);
    font-size: 13px;
  }
  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 14px;
    margin: 20px 0;
  }
  .card {
    background: linear-gradient(180deg, var(--panel), var(--panel-2));
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 16px;
    box-shadow: 0 12px 30px rgba(0,0,0,.2);
  }
  .label {
    color: var(--muted);
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .08em;
    margin-bottom: 10px;
  }
  .value {
    font-size: 28px;
    font-weight: 700;
  }
  .ok { color: var(--ok); }
  .warn { color: var(--warn); }
  .crit { color: var(--crit); }
  .accent { color: var(--accent); }
  .section {
    margin-top: 22px;
  }
  .section h2 {
    margin: 0 0 12px;
    font-size: 18px;
  }
  .toolbar {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    margin-bottom: 12px;
  }
  .toolbar input {
    background: #0c1524;
    color: var(--text);
    border: 1px solid var(--line);
    border-radius: 10px;
    padding: 10px 12px;
    min-width: 260px;
  }
  .toolbar button {
    background: #0c1524;
    color: var(--muted);
    border: 1px solid var(--line);
    border-radius: 999px;
    padding: 9px 12px;
    cursor: pointer;
  }
  .toolbar button.active {
    color: var(--text);
    border-color: var(--accent);
  }
  .coverage {
    height: 10px;
    background: #0c1524;
    border-radius: 999px;
    overflow: hidden;
    border: 1px solid var(--line);
  }
  .coverage > div {
    height: 100%;
    background: linear-gradient(90deg, var(--ok), #52d89a);
  }
  .method-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
  }
  .method-item {
    background: #0c1524;
    border: 1px solid var(--line);
    border-radius: 12px;
    padding: 12px;
  }
  .table-wrap {
    overflow: auto;
    border: 1px solid var(--line);
    border-radius: 12px;
    background: #0c1524;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    min-width: 1080px;
  }
  th, td {
    padding: 10px 12px;
    border-bottom: 1px solid var(--line);
    text-align: left;
    vertical-align: top;
  }
  th {
    position: sticky;
    top: 0;
    background: #122038;
    color: var(--muted);
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .06em;
  }
  tr:hover td {
    background: rgba(255,255,255,.02);
  }
  .pill {
    display: inline-block;
    padding: 3px 8px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 600;
  }
  .pill-ok {
    background: rgba(31,181,108,.14);
    color: #74e2aa;
  }
  .pill-crit {
    background: rgba(239,91,107,.14);
    color: #ff97a3;
  }
  .pill-admin {
    background: rgba(240,173,44,.14);
    color: #ffd07a;
  }
  .muted { color: var(--muted); }
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <div>
      <h1>M365 MFA Dashboard</h1>
      <div class="meta">Tenant: <strong>$([System.Web.HttpUtility]::HtmlEncode((if ($TenantDomain) { $TenantDomain } else { $TenantId })))</strong></div>
      <div class="meta">Generated: $([System.Web.HttpUtility]::HtmlEncode($ReportDate))</div>
    </div>
  </div>

  <div class="grid">
    <div class="card"><div class="label">Total Users</div><div class="value accent" id="totalUsers"></div></div>
    <div class="card"><div class="label">MFA Registered</div><div class="value ok" id="registered"></div></div>
    <div class="card"><div class="label">MFA Not Registered</div><div class="value crit" id="notRegistered"></div></div>
    <div class="card"><div class="label">Admin Accounts</div><div class="value warn" id="admins"></div></div>
    <div class="card"><div class="label">Admins Without MFA</div><div class="value crit" id="adminNoMfa"></div></div>
  </div>

  <div class="section card">
    <h2>MFA Coverage</h2>
    <div class="meta" id="coverageText"></div>
    <div class="coverage" style="margin-top:10px;"><div id="coverageBar"></div></div>
  </div>

  <div class="section card">
    <h2>Method Breakdown</h2>
    <div class="method-grid" id="methodGrid"></div>
  </div>

  <div class="section card">
    <h2>User Detail</h2>
    <div class="toolbar">
      <input id="searchBox" type="text" placeholder="Filter by name, UPN, method or status">
      <button data-filter="all" class="active">All</button>
      <button data-filter="no-mfa">No MFA</button>
      <button data-filter="admin">Admins</button>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Display Name</th>
            <th>UPN</th>
            <th>User Type</th>
            <th>Admin</th>
            <th>MFA Status</th>
            <th>Default Method</th>
            <th>Method Count</th>
            <th>Methods</th>
            <th>Last Sign-In</th>
          </tr>
        </thead>
        <tbody id="userRows"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const report = $htmlData;
const userRows = document.getElementById('userRows');
const searchBox = document.getElementById('searchBox');
const filterButtons = [...document.querySelectorAll('[data-filter]')];
let activeFilter = 'all';

document.getElementById('totalUsers').textContent = report.TotalUsers;
document.getElementById('registered').textContent = report.Registered;
document.getElementById('notRegistered').textContent = report.NotRegistered;
document.getElementById('admins').textContent = report.Admins;
document.getElementById('adminNoMfa').textContent = report.AdminNoMfa;
document.getElementById('coverageText').textContent = report.CoveragePct + '% of users have MFA registered';
document.getElementById('coverageBar').style.width = report.CoveragePct + '%';

const methods = [
  ['Authenticator App', report.MethodCounts.AuthenticatorApp],
  ['FIDO2 Security Key', report.MethodCounts.Fido2],
  ['Phone', report.MethodCounts.Phone],
  ['Software TOTP', report.MethodCounts.Totp],
  ['Email OTP', report.MethodCounts.EmailOtp],
  ['Windows Hello', report.MethodCounts.Whfb],
  ['Passkey', report.MethodCounts.Passkey]
];

document.getElementById('methodGrid').innerHTML = methods.map(([label, count]) =>
  `<div class="method-item"><div class="label">${label}</div><div class="value accent" style="font-size:22px;">${count}</div></div>`
).join('');

function esc(value) {
  return String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('\"', '&quot;');
}

function statusPill(status) {
  return status === 'Registered'
    ? '<span class="pill pill-ok">Registered</span>'
    : '<span class="pill pill-crit">Not Registered</span>';
}

function adminPill(isAdmin) {
  return isAdmin ? '<span class="pill pill-admin">Admin</span>' : '<span class="muted">No</span>';
}

function renderRows() {
  const term = searchBox.value.trim().toLowerCase();
  const rows = report.Users.filter((user) => {
    const matchesFilter =
      activeFilter === 'all' ||
      (activeFilter === 'no-mfa' && !user.MfaRegistered) ||
      (activeFilter === 'admin' && user.IsAdmin);

    const haystack = [
      user.DisplayName,
      user.UPN,
      user.MfaStatus,
      user.DefaultMethod,
      user.Methods,
      user.LastSignIn
    ].join(' ').toLowerCase();

    return matchesFilter && (!term || haystack.includes(term));
  });

  userRows.innerHTML = rows.map((user) => `
    <tr>
      <td>${esc(user.DisplayName)}</td>
      <td>${esc(user.UPN)}</td>
      <td>${esc(user.UserType)}</td>
      <td>${adminPill(user.IsAdmin)}</td>
      <td>${statusPill(user.MfaStatus)}</td>
      <td>${esc(user.DefaultMethod)}</td>
      <td>${esc(user.MethodCount)}</td>
      <td>${esc(user.Methods || '')}</td>
      <td>${esc(user.LastSignIn)}</td>
    </tr>
  `).join('');
}

searchBox.addEventListener('input', renderRows);
for (const button of filterButtons) {
  button.addEventListener('click', () => {
    activeFilter = button.dataset.filter;
    filterButtons.forEach((b) => b.classList.toggle('active', b === button));
    renderRows();
  });
}

renderRows();
</script>
</body>
</html>
"@

    $outDir = [System.IO.Path]::GetDirectoryName($Path)
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $html | Out-File -LiteralPath $Path -Encoding UTF8
    Write-Host "[+] HTML dashboard exported to: $Path" -ForegroundColor Green
}

Assert-RequiredModules
Connect-ToGraph -TenantId $TenantId
Ensure-ReportScopes
Resolve-TenantDomain
$adminIds = Get-AdminUserIds -RoleNames $AdminRoles
$report = Get-MfaReport -IncludeGuests $IncludeGuests.IsPresent -AdminIds $adminIds
Show-ConsoleSummary -Report $report

if ($ExportHtml) {
    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -DefaultExtension ".html"
    Export-MfaHtmlReport -Path $htmlPath -Report $report -TenantDomain $script:TenantDomain -ReportDate $reportDate
}
