#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Checks MFA authentication methods for all users in a Microsoft 365 tenant.

.DESCRIPTION
    Connects to Microsoft Graph API and retrieves all users, then checks
    which authentication (MFA) methods each user has registered.
    Outputs a report to the console and optionally exports to CSV.

.PARAMETER ExportCsv
    If specified, exports the results to a CSV file at the given path.
    Example: -ExportCsv "C:\Reports\MFA_Report.csv"

.PARAMETER IncludeGuests
    If specified, includes guest accounts in the report.

.PARAMETER SkipDisabled
    If specified, skips disabled user accounts.

.EXAMPLE
    .\Check-MFAStatus.ps1
    .\Check-MFAStatus.ps1 -ExportCsv "C:\Reports\MFA_Report.csv"
    .\Check-MFAStatus.ps1 -ExportCsv "C:\Reports\MFA_Report.csv" -SkipDisabled
#>

[CmdletBinding()]
param (
    [string]$ExportCsv,
    [switch]$IncludeGuests,
    [switch]$SkipDisabled
)

function Connect-ToMicrosoftGraph {
    Write-Host "`n[*] Connecting to Microsoft Graph..." -ForegroundColor Cyan

    $requiredScopes = @(
        "User.Read.All",
        "UserAuthenticationMethod.Read.All"
    )

    try {
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        $context = Get-MgContext
        Write-Host "[+] Connected as: $($context.Account)" -ForegroundColor Green
        Write-Host "[+] Tenant ID:    $($context.TenantId)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

function Get-UserList {
    param (
        [bool]$IncludeGuests,
        [bool]$SkipDisabled
    )

    Write-Host "`n[*] Retrieving users from tenant..." -ForegroundColor Cyan

    $filter = $null

    if (-not $IncludeGuests) {
        $filter = "userType eq 'Member'"
    }

    $params = @{
        All         = $true
        Property    = "Id,DisplayName,UserPrincipalName,AccountEnabled,UserType,Department,JobTitle"
        ErrorAction = "Stop"
    }

    if ($filter) {
        $params["Filter"] = $filter
    }

    try {
        $users = Get-MgUser @params

        if ($SkipDisabled) {
            $users = $users | Where-Object { $_.AccountEnabled -eq $true }
        }

        Write-Host "[+] Found $($users.Count) user(s) to process." -ForegroundColor Green
        return $users
    }
    catch {
        Write-Error "Failed to retrieve users: $_"
        exit 1
    }
}

function Get-MFAMethodsForUser {
    param (
        [string]$UserId
    )

    $methods = @{
        HasMFA                  = $false
        MicrosoftAuthenticator  = $false
        PhoneAuthentication     = $false
        FIDO2SecurityKey        = $false
        SoftwareOath            = $false
        WindowsHelloForBusiness = $false
        EmailOtp                = $false
        TemporaryAccessPass     = $false
        MethodCount             = 0
        MethodNames             = @()
    }

    try {
        $authApp = Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($authApp) {
            $methods.MicrosoftAuthenticator = $true
            $methods.MethodNames += "MicrosoftAuthenticator"
        }

        $phone = Get-MgUserAuthenticationPhoneMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($phone) {
            $methods.PhoneAuthentication = $true
            $methods.MethodNames += "PhoneSMS/Voice"
        }

        $fido2 = Get-MgUserAuthenticationFido2Method -UserId $UserId -ErrorAction SilentlyContinue
        if ($fido2) {
            $methods.FIDO2SecurityKey = $true
            $methods.MethodNames += "FIDO2SecurityKey"
        }

        $softOath = Get-MgUserAuthenticationSoftwareOathMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($softOath) {
            $methods.SoftwareOath = $true
            $methods.MethodNames += "SoftwareOATH(TOTP)"
        }

        $wHello = Get-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($wHello) {
            $methods.WindowsHelloForBusiness = $true
            $methods.MethodNames += "WindowsHelloForBusiness"
        }

        $email = Get-MgUserAuthenticationEmailMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($email) {
            $methods.EmailOtp = $true
            $methods.MethodNames += "EmailOTP"
        }

        $tap = Get-MgUserAuthenticationTemporaryAccessPassMethod -UserId $UserId -ErrorAction SilentlyContinue
        if ($tap) {
            $methods.TemporaryAccessPass = $true
            $methods.MethodNames += "TemporaryAccessPass"
        }
    }
    catch {
        Write-Warning "  Could not retrieve methods for user $UserId : $_"
    }

    $methods.MethodCount = $methods.MethodNames.Count
    $methods.HasMFA = $methods.MethodCount -gt 0

    return $methods
}

function Show-Summary {
    param ($Results)

    $total = $Results.Count
    $mfaEnabled = ($Results | Where-Object { $_.MFAEnabled -eq $true }).Count
    $noMfa = $total - $mfaEnabled
    $pct = if ($total -gt 0) { [math]::Round(($mfaEnabled / $total) * 100, 1) } else { 0 }

    Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "           MFA STATUS SUMMARY           " -ForegroundColor Cyan
    Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Total Users Checked : $total"
    Write-Host "  MFA Enabled         : $mfaEnabled ($pct%)" -ForegroundColor Green
    Write-Host "  No MFA Registered   : $noMfa" -ForegroundColor $(if ($noMfa -gt 0) { "Red" } else { "Green" })
    Write-Host "════════════════════════════════════════`n" -ForegroundColor Cyan

    if ($noMfa -gt 0) {
        Write-Host "[!] Users WITHOUT any MFA method:" -ForegroundColor Yellow
        $Results | Where-Object { $_.MFAEnabled -eq $false } |
            Select-Object DisplayName, UserPrincipalName, AccountEnabled |
            Format-Table -AutoSize
    }
}

$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.SignIns"
)

Write-Host "`n[*] Checking required PowerShell modules..." -ForegroundColor Cyan

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "  [!] Module '$mod' not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "  [+] Installed $mod" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install module '$mod': $_"
            exit 1
        }
    }
    else {
        Write-Host "  [+] $mod is available." -ForegroundColor Green
    }
}

Connect-ToMicrosoftGraph

$users = Get-UserList -IncludeGuests $IncludeGuests.IsPresent -SkipDisabled $SkipDisabled.IsPresent

$report = [System.Collections.Generic.List[PSObject]]::new()
$counter = 0

Write-Host "`n[*] Checking MFA methods for each user...`n" -ForegroundColor Cyan

foreach ($user in $users) {
    $counter++
    $progress = [math]::Round(($counter / $users.Count) * 100)

    Write-Progress -Activity "Checking MFA Status" `
                   -Status "$counter / $($users.Count) - $($user.UserPrincipalName)" `
                   -PercentComplete $progress

    $mfa = Get-MFAMethodsForUser -UserId $user.Id

    $record = [PSCustomObject]@{
        DisplayName             = $user.DisplayName
        UserPrincipalName       = $user.UserPrincipalName
        AccountEnabled          = $user.AccountEnabled
        UserType                = $user.UserType
        Department              = $user.Department
        JobTitle                = $user.JobTitle
        MFAEnabled              = $mfa.HasMFA
        MethodCount             = $mfa.MethodCount
        RegisteredMethods       = ($mfa.MethodNames -join ", ")
        MicrosoftAuthenticator  = $mfa.MicrosoftAuthenticator
        PhoneSMS_Voice          = $mfa.PhoneAuthentication
        FIDO2SecurityKey        = $mfa.FIDO2SecurityKey
        SoftwareOATH_TOTP       = $mfa.SoftwareOath
        WindowsHelloForBusiness = $mfa.WindowsHelloForBusiness
        EmailOTP                = $mfa.EmailOtp
        TemporaryAccessPass     = $mfa.TemporaryAccessPass
    }

    $report.Add($record)
}

Write-Progress -Activity "Checking MFA Status" -Completed

$report | Format-Table DisplayName, UserPrincipalName, AccountEnabled, MFAEnabled, MethodCount, RegisteredMethods -AutoSize

Show-Summary -Results $report

if ($ExportCsv) {
    try {
        $report | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
        Write-Host "[+] Report exported to: $ExportCsv" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export CSV: $_"
    }
}

Disconnect-MgGraph | Out-Null
Write-Host "[*] Disconnected from Microsoft Graph.`n" -ForegroundColor Cyan
