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

    $sortedUsers = @(
        $Report | Sort-Object -Property @(
            @{ Expression = 'IsAdmin'; Descending = $true },
            @{ Expression = 'MfaRegistered'; Descending = $true },
            @{ Expression = 'DisplayName'; Descending = $false }
        )
    )

    $htmlData = [PSCustomObject]@{
        tenant     = if ($TenantDomain) { $TenantDomain } else { $TenantId }
        reportDate = $ReportDate
        total      = $total
        registered = $registered
        notReg     = $notRegistered
        pct        = $pct
        admins     = $admins
        adminNoMfa = $adminNoMfa
        cntAuthApp = @($Report | Where-Object { $_.HasAuthApp }).Count
        cntFido2   = @($Report | Where-Object { $_.HasFIDO2 }).Count
        cntPhone   = @($Report | Where-Object { $_.HasPhone }).Count
        cntTotp    = @($Report | Where-Object { $_.HasTOTP }).Count
        cntEmail   = @($Report | Where-Object { $_.HasEmailOTP }).Count
        cntWhfb    = @($Report | Where-Object { $_.HasWHfB }).Count
        cntPasskey = @($Report | Where-Object { $_.HasPasskey }).Count
        users      = $sortedUsers
    } | ConvertTo-Json -Depth 6 -Compress

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 MFA Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg:#f5f7fb; --bg2:#ffffff; --bg3:#f7f9fc; --bg4:#eef3f9;
    --border:#d6dfec; --border2:#b8c8dc;
    --text:#17324d; --text2:#4b6887; --text3:#6f88a3;
    --accent:#0f7cc0; --accent2:#1091d5;
    --ok:#15803d; --ok-bg:rgba(21,128,61,.1);
    --warn:#b45309; --warn-bg:rgba(180,83,9,.1);
    --crit:#dc2626; --crit-bg:rgba(220,38,38,.1);
    --mono:'JetBrains Mono',monospace; --sans:'DM Sans',sans-serif;
    --r:8px; --r2:12px;
  }
  html { scroll-behavior:smooth; }
  body { font-family:var(--sans); background:var(--bg); color:var(--text); min-height:100vh; line-height:1.6; }
  body::before { content:''; position:fixed; inset:0; background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='0.025'/%3E%3C/svg%3E"); pointer-events:none; z-index:0; opacity:.35; }
  .topbar { position:sticky; top:0; z-index:100; background:rgba(245,247,251,.88); backdrop-filter:blur(16px); border-bottom:1px solid var(--border); padding:0 2rem; height:52px; display:flex; align-items:center; gap:1.5rem; }
  .topbar-logo { font-family:var(--mono); font-size:.72rem; font-weight:700; letter-spacing:.16em; text-transform:uppercase; color:var(--accent); }
  .topbar-sep  { color:var(--border2); }
  .topbar-server { font-family:var(--mono); font-size:.78rem; color:var(--text2); }
  .topbar-nav { display:flex; gap:.15rem; margin-left:.75rem; }
  .topbar-nav a { font-size:.72rem; font-weight:600; color:var(--text3); text-decoration:none; padding:.3rem .65rem; border-radius:var(--r); transition:all .15s; }
  .topbar-nav a:hover { background:var(--bg3); color:var(--text); }
  .topbar-right { margin-left:auto; display:flex; align-items:center; gap:1rem; }
  .topbar-time { font-family:var(--mono); font-size:.7rem; color:var(--text3); }
  .page { max-width:1440px; margin:0 auto; padding:1.75rem 2rem; position:relative; z-index:1; }
  .server-strip { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2); padding:1rem 1.5rem; display:flex; flex-wrap:wrap; gap:2.5rem; margin-bottom:1.5rem; }
  .strip-item { display:flex; flex-direction:column; gap:.2rem; }
  .strip-label { font-size:.63rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:var(--text3); }
  .strip-value { font-family:var(--mono); font-size:.82rem; color:var(--accent); }
  .hero { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:.85rem; margin-bottom:1.5rem; }
  .kpi { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2); padding:1rem 1.1rem; position:relative; overflow:hidden; transition:transform .2s,border-color .2s; cursor:default; }
  .kpi:hover { transform:translateY(-2px); border-color:var(--border2); }
  .kpi::after { content:''; position:absolute; top:0;left:0;right:0; height:2px; border-radius:2px 2px 0 0; }
  .kpi.ok::after { background:var(--ok); }
  .kpi.warn::after { background:var(--warn); }
  .kpi.crit::after { background:var(--crit); }
  .kpi.neutral::after { background:var(--accent2); }
  .kpi-label { font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); margin-bottom:.4rem; }
  .kpi-value { font-family:var(--mono); font-size:1.7rem; font-weight:700; line-height:1; }
  .kpi.ok .kpi-value { color:var(--ok); }
  .kpi.warn .kpi-value { color:var(--warn); }
  .kpi.crit .kpi-value { color:var(--crit); }
  .kpi.neutral .kpi-value { color:var(--accent); }
  .kpi-sub { font-size:.68rem; color:var(--text3); margin-top:.35rem; font-family:var(--mono); }
  .sections { display:grid; gap:1.25rem; }
  .card { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2); overflow:hidden; }
  .card-header { display:flex; align-items:center; gap:.75rem; padding:.85rem 1.25rem; border-bottom:1px solid var(--border); cursor:pointer; user-select:none; transition:background .15s; }
  .card-header:hover { background:var(--bg3); }
  .card-icon { font-size:.95rem; width:1.4rem; text-align:center; opacity:.7; }
  .card-title { font-size:.72rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text2); flex:1; }
  .card-badge { font-family:var(--mono); font-size:.7rem; padding:.18rem .55rem; border-radius:4px; }
  .badge-ok { background:var(--ok-bg); color:var(--ok); }
  .badge-warn { background:var(--warn-bg); color:var(--warn); }
  .badge-crit { background:var(--crit-bg); color:var(--crit); }
  .badge-neutral { background:rgba(15,124,192,.1); color:var(--accent); }
  .card-chevron { font-size:.65rem; color:var(--text3); transition:transform .2s; }
  .card.collapsed .card-chevron { transform:rotate(-90deg); }
  .card-body { padding:1.25rem; }
  .card.collapsed .card-body { display:none; }
  .toolbar { display:flex; gap:.65rem; margin-bottom:1rem; flex-wrap:wrap; align-items:center; }
  .search-box { flex:1; min-width:220px; max-width:360px; background:var(--bg3); border:1px solid var(--border2); border-radius:var(--r); color:var(--text); font-family:var(--mono); font-size:.78rem; padding:.4rem .75rem; outline:none; transition:border-color .15s; }
  .search-box:focus { border-color:var(--accent2); }
  .search-box::placeholder { color:var(--text3); }
  .filter-btn { font-family:var(--sans); font-size:.72rem; font-weight:700; padding:.3rem .8rem; border-radius:5px; border:1px solid var(--border); background:var(--bg3); color:var(--text3); cursor:pointer; transition:all .15s; }
  .filter-btn:hover { color:var(--text); border-color:var(--border2); }
  .filter-btn.active-all { border-color:var(--accent2); color:var(--accent); background:rgba(15,124,192,.08); }
  .filter-btn.active-crit { border-color:var(--crit); color:var(--crit); background:var(--crit-bg); }
  .filter-btn.active-ok { border-color:var(--ok); color:var(--ok); background:var(--ok-bg); }
  .filter-btn.active-admin { border-color:var(--warn); color:var(--warn); background:var(--warn-bg); }
  .table-scroll { max-height:560px; overflow-y:auto; border:1px solid var(--border); border-radius:var(--r); }
  table { width:100%; border-collapse:collapse; font-size:.77rem; }
  thead { background:var(--bg3); position:sticky; top:0; z-index:1; }
  th { padding:.55rem .9rem; text-align:left; font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); white-space:nowrap; border-bottom:1px solid var(--border); }
  td { padding:.5rem .9rem; border-bottom:1px solid var(--border); color:var(--text); font-family:var(--mono); font-size:.76rem; vertical-align:top; max-width:340px; word-break:break-all; }
  tr:last-child td { border-bottom:none; }
  tbody tr:hover td { background:rgba(15,124,192,.05); }
  .pill { display:inline-block; padding:.1rem .45rem; border-radius:4px; font-size:.7rem; font-weight:600; white-space:nowrap; }
  .method-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(190px,1fr)); gap:.85rem; }
  .method-item { background:var(--bg3); border:1px solid var(--border); border-radius:var(--r); padding:.85rem 1rem; display:flex; align-items:center; gap:.85rem; }
  .method-icon { font-size:1.25rem; opacity:.7; width:2rem; text-align:center; flex-shrink:0; }
  .method-info { display:flex; flex-direction:column; gap:.2rem; }
  .method-label { font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); }
  .method-count { font-family:var(--mono); font-size:1.05rem; font-weight:700; color:var(--accent); }
  .coverage-wrap { margin-bottom:1.25rem; }
  .coverage-label { display:flex; justify-content:space-between; margin-bottom:.4rem; }
  .coverage-label-text { font-size:.72rem; color:var(--text2); font-family:var(--mono); }
  .coverage-pct { font-family:var(--mono); font-size:.72rem; }
  .coverage-bar-bg { height:8px; background:var(--border); border-radius:4px; overflow:hidden; }
  .coverage-bar { height:100%; border-radius:4px; transition:width 1s cubic-bezier(.25,.46,.45,.94); }
  .coverage-bar.ok { background:linear-gradient(90deg,var(--ok),#34d399); }
  .coverage-bar.warn { background:linear-gradient(90deg,var(--warn),#fcd34d); }
  .coverage-bar.crit { background:linear-gradient(90deg,var(--crit),#fb7185); }
  .empty { font-family:var(--mono); font-size:.78rem; color:var(--text3); padding:.75rem 0; font-style:italic; }
  @media(max-width:640px){.topbar-nav{display:none}.hero{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-logo">&#x2B21; RegITs GmbH M365 MFA Monitor</div>
  <span class="topbar-sep">|</span>
  <div class="topbar-server"></div>
  <nav class="topbar-nav">
    <a href="#sec-overview">Overview</a>
    <a href="#sec-methods">Methods</a>
    <a href="#sec-users">All Users</a>
    <a href="#sec-nomfa">No MFA</a>
    <a href="#sec-admins">Admins</a>
  </nav>
  <div class="topbar-right"><div class="topbar-time"></div></div>
</div>
<div class="page">
  <div id="server-strip" class="server-strip"></div>
  <div id="hero" class="hero"></div>
  <div id="sections" class="sections"></div>
</div>
<script>
const DATA = $htmlData;
document.querySelector('.topbar-server').textContent = DATA.tenant;
document.querySelector('.topbar-time').textContent = 'Generated: ' + DATA.reportDate;
document.title = 'M365 MFA Report - ' + DATA.tenant;
function esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function kpi(value,label,sub,cls){return '<div class="kpi ' + cls + '"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>';}
function strip(items){return items.map(function(pair){return '<div class="strip-item"><span class="strip-label">' + pair[0] + '</span><span class="strip-value">' + esc(pair[1]) + '</span></div>';}).join('');}
function card(id,icon,title,badgeText,badgeClass,body){return '<div class="card" id="' + id + '"><div class="card-header" onclick="toggleCard(this)"><span class="card-icon">' + icon + '</span><span class="card-title">' + title + '</span><span class="card-badge ' + badgeClass + '">' + badgeText + '</span><span class="card-chevron">&#x25BC;</span></div><div class="card-body">' + body + '</div></div>';}
function toggleCard(h){h.closest('.card').classList.toggle('collapsed');}
  function methodPill(b){return b?'<span class="pill badge-ok">Yes</span>':'<span class="pill" style="background:rgba(220,38,38,.08);color:#dc2626">No</span>';}
function mfaStatusPill(v){return v==='Registered'?'<span class="pill badge-ok">' + esc(v) + '</span>':'<span class="pill badge-crit">' + esc(v) + '</span>';}
  function adminPill(v){return v===true?'<span class="pill badge-warn">Admin</span>':'<span class="pill" style="background:rgba(15,124,192,.07);color:var(--accent)">User</span>';}
function methodLabel(v){const map={'Authenticator App':'ok','FIDO2 Key':'ok','Passkey':'ok','WHfB':'ok','TOTP':'ok','Phone':'warn','Email OTP':'warn','Password Only':'crit','None':'crit'};const cls=map[v]||'neutral';return '<span class="pill badge-' + cls + '">' + esc(v) + '</span>';}
let _activeFilter='all'; let _searchTerm='';
function applyUserFilter(){const rows=document.querySelectorAll('#user-tbody tr');rows.forEach(function(row){const cells=Array.from(row.cells).map(function(c){return c.textContent.toLowerCase();}).join(' ');const matchSearch=!_searchTerm||cells.includes(_searchTerm);const mfaCell=row.cells[2]?row.cells[2].textContent.toLowerCase():'';const adminCell=row.cells[3]?row.cells[3].textContent.toLowerCase():'';let matchFilter=true;if(_activeFilter==='nomfa') matchFilter=mfaCell.includes('not');if(_activeFilter==='mfa') matchFilter=!mfaCell.includes('not');if(_activeFilter==='admin') matchFilter=adminCell.includes('admin');row.style.display=(matchSearch&&matchFilter)?'':'none';});}
function setFilter(f,btn){_activeFilter=f;document.querySelectorAll('.uf-btn').forEach(function(b){b.classList.remove('active-all','active-crit','active-ok','active-admin');});if(f==='all') btn.classList.add('active-all');if(f==='nomfa') btn.classList.add('active-crit');if(f==='mfa') btn.classList.add('active-ok');if(f==='admin') btn.classList.add('active-admin');applyUserFilter();}
function render(){const d=DATA;const pct=d.pct;const coverageClass=pct>=90?'ok':pct>=70?'warn':'crit';const coveragePctColor=pct>=90?'var(--ok)':pct>=70?'var(--warn)':'var(--crit)';document.getElementById('server-strip').innerHTML=strip([['Tenant',d.tenant],['Total Users',String(d.total)],['MFA Coverage',pct + '%'],['Admins',String(d.admins)],['Admins Without MFA',String(d.adminNoMfa)],['Report Generated',d.reportDate]]);const mfaPctCls=pct>=90?'ok':pct>=70?'warn':'crit';document.getElementById('hero').innerHTML=[kpi(d.total,'Total<br>Users','active members','neutral'),kpi(d.registered,'MFA<br>Registered','have 2nd factor',d.registered===d.total?'ok':'warn'),kpi(d.notReg,'MFA<br>NOT Registered','password only',d.notReg>0?'crit':'ok'),kpi(pct + '%','MFA<br>Coverage','of all users',mfaPctCls),kpi(d.admins,'Admin<br>Accounts','privileged roles','neutral'),kpi(d.adminNoMfa,'Admins<br>Without MFA','critical risk',d.adminNoMfa>0?'crit':'ok'),kpi(d.cntAuthApp,'Authenticator<br>App','microsoft/3rd party','neutral'),kpi(d.cntFido2,'FIDO2<br>Keys','hardware tokens',d.cntFido2>0?'ok':'neutral'),kpi(d.cntPhone,'Phone<br>SMS/Voice','legacy method',d.cntPhone>0?'warn':'neutral')].join('');const sec=document.getElementById('sections');sec.innerHTML='';const overviewBody='<div class="coverage-wrap"><div class="coverage-label"><span class="coverage-label-text">MFA Coverage - ' + d.registered + ' of ' + d.total + ' users</span><span class="coverage-pct" style="color:' + coveragePctColor + '">' + pct + '%</span></div><div class="coverage-bar-bg"><div class="coverage-bar ' + coverageClass + '" style="width:' + pct + '%"></div></div></div>' + (d.adminNoMfa>0?'<div style="background:var(--crit-bg);border:1px solid var(--crit);border-radius:var(--r);padding:.75rem 1rem;margin-bottom:1rem;font-size:.78rem;font-family:var(--mono);color:var(--crit)"><strong>' + d.adminNoMfa + ' admin account(s)</strong> have no MFA registered - immediate action required.</div>':'') + '<div class="method-grid"><div class="method-item"><div class="method-icon" style="color:#38bdf8"><i class="fa-solid fa-mobile-screen-button"></i></div><div class="method-info"><div class="method-label">Authenticator App</div><div class="method-count">' + d.cntAuthApp + '</div></div></div><div class="method-item"><div class="method-icon" style="color:#a78bfa"><i class="fa-solid fa-key"></i></div><div class="method-info"><div class="method-label">FIDO2 Security Key</div><div class="method-count">' + d.cntFido2 + '</div></div></div><div class="method-item"><div class="method-icon" style="color:#f59e0b"><i class="fa-solid fa-phone"></i></div><div class="method-info"><div class="method-label">Phone (SMS/Voice)</div><div class="method-count">' + d.cntPhone + '</div></div></div><div class="method-item"><div class="method-icon" style="color:#10b981"><i class="fa-solid fa-clock-rotate-left"></i></div><div class="method-info"><div class="method-label">TOTP / OATH</div><div class="method-count">' + d.cntTotp + '</div></div></div><div class="method-item"><div class="method-icon" style="color:#6a88b0"><i class="fa-solid fa-envelope"></i></div><div class="method-info"><div class="method-label">Email OTP</div><div class="method-count">' + d.cntEmail + '</div></div></div><div class="method-item"><div class="method-icon" style="color:#38bdf8"><i class="fa-brands fa-windows"></i></div><div class="method-info"><div class="method-label">Windows Hello (WHfB)</div><div class="method-count">' + d.cntWhfb + '</div></div></div><div class="method-item"><div class="method-icon" style="color:#fb923c"><i class="fa-solid fa-fingerprint"></i></div><div class="method-info"><div class="method-label">Passkey</div><div class="method-count">' + d.cntPasskey + '</div></div></div></div>';sec.innerHTML+=card('sec-overview','<i class="fa-solid fa-chart-pie"></i>','MFA Coverage Overview',pct + '% covered',pct>=90?'badge-ok':pct>=70?'badge-warn':'badge-crit',overviewBody);const methodCols=['DisplayName','UPN','DefaultMethod','HasAuthApp','HasFIDO2','HasPhone','HasTOTP','HasEmailOTP','HasWHfB','HasPasskey','LastSignIn'];let mTable='<div class="table-scroll"><table><thead><tr>';for(const c of methodCols){mTable+='<th>' + esc(c.replace(/Has/,'')) + '</th>';}mTable+='</tr></thead><tbody>';for(const r of d.users){mTable+='<tr><td>' + esc(r.DisplayName) + '</td><td style="font-size:.72rem">' + esc(r.UPN) + '</td><td>' + methodLabel(r.DefaultMethod) + '</td><td>' + methodPill(r.HasAuthApp) + '</td><td>' + methodPill(r.HasFIDO2) + '</td><td>' + methodPill(r.HasPhone) + '</td><td>' + methodPill(r.HasTOTP) + '</td><td>' + methodPill(r.HasEmailOTP) + '</td><td>' + methodPill(r.HasWHfB) + '</td><td>' + methodPill(r.HasPasskey) + '</td><td>' + esc(r.LastSignIn) + '</td></tr>';}mTable+='</tbody></table></div>';sec.innerHTML+=card('sec-methods','<i class="fa-solid fa-shield-halved"></i>','Authentication Methods per User',d.total + ' users','badge-neutral',mTable);let uTable='<div class="toolbar"><input class="search-box" type="text" placeholder="Search name or UPN..." oninput="_searchTerm=this.value.toLowerCase();applyUserFilter()"><button class="filter-btn uf-btn active-all" onclick="setFilter(\'all\',this)">All (' + d.total + ')</button><button class="filter-btn uf-btn" onclick="setFilter(\'mfa\',this)">MFA Registered (' + d.registered + ')</button><button class="filter-btn uf-btn" onclick="setFilter(\'nomfa\',this)">No MFA (' + d.notReg + ')</button><button class="filter-btn uf-btn" onclick="setFilter(\'admin\',this)">Admins (' + d.admins + ')</button></div><div class="table-scroll"><table><thead><tr><th>Display Name</th><th>UPN</th><th>MFA Status</th><th>Role</th><th>Default Method</th><th># Methods</th><th>All Methods</th><th>Last Sign-In</th></tr></thead><tbody id="user-tbody">';for(const r of d.users){uTable+='<tr><td>' + esc(r.DisplayName) + '</td><td style="font-size:.72rem">' + esc(r.UPN) + '</td><td>' + mfaStatusPill(r.MfaStatus) + '</td><td>' + adminPill(r.IsAdmin) + '</td><td>' + methodLabel(r.DefaultMethod) + '</td><td>' + esc(r.MethodCount) + '</td><td style="font-size:.7rem;color:var(--text2)">' + esc(r.Methods || '—') + '</td><td>' + esc(r.LastSignIn) + '</td></tr>';}uTable+='</tbody></table></div>';sec.innerHTML+=card('sec-users','<i class="fa-solid fa-users"></i>','All Users',d.total + ' users','badge-neutral',uTable);const noMfaUsers=d.users.filter(function(u){return u.MfaStatus!=='Registered';});let noMfaBody='';if(noMfaUsers.length===0){noMfaBody='<p class="empty">Excellent! Every user has at least one MFA method registered.</p>';} else {noMfaBody='<div class="table-scroll"><table><thead><tr><th>Display Name</th><th>UPN</th><th>Role</th><th>Last Sign-In</th></tr></thead><tbody>';for(const r of noMfaUsers){noMfaBody+='<tr><td>' + esc(r.DisplayName) + '</td><td style="font-size:.72rem">' + esc(r.UPN) + '</td><td>' + adminPill(r.IsAdmin) + '</td><td>' + esc(r.LastSignIn) + '</td></tr>';}noMfaBody+='</tbody></table></div>';}sec.innerHTML+=card('sec-nomfa','<i class="fa-solid fa-triangle-exclamation"></i>','Users Without MFA',noMfaUsers.length===0?'All users protected':(noMfaUsers.length + ' user(s) at risk'),noMfaUsers.length===0?'badge-ok':'badge-crit',noMfaBody);const adminUsers=d.users.filter(function(u){return u.IsAdmin===true;});let adminBody='';if(adminUsers.length===0){adminBody='<p class="empty">No privileged admin accounts found in the scanned roles.</p>';} else {adminBody='<div class="table-scroll"><table><thead><tr><th>Display Name</th><th>UPN</th><th>MFA Status</th><th>Default Method</th><th>All Methods</th><th>Last Sign-In</th></tr></thead><tbody>';for(const r of adminUsers){adminBody+='<tr><td>' + esc(r.DisplayName) + '</td><td style="font-size:.72rem">' + esc(r.UPN) + '</td><td>' + mfaStatusPill(r.MfaStatus) + '</td><td>' + methodLabel(r.DefaultMethod) + '</td><td style="font-size:.7rem;color:var(--text2)">' + esc(r.Methods || '—') + '</td><td>' + esc(r.LastSignIn) + '</td></tr>';}adminBody+='</tbody></table></div>';}sec.innerHTML+=card('sec-admins','<i class="fa-solid fa-user-shield"></i>','Privileged Admin Accounts',d.adminNoMfa + ' without MFA',d.adminNoMfa>0?'badge-crit':d.admins>0?'badge-ok':'badge-neutral',adminBody);}
render();
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
