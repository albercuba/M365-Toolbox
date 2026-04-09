<#
.SYNOPSIS
    Performs review and remediation actions for compromised Microsoft 365 user accounts.

.DESCRIPTION
    This script is intended for incident-response scenarios involving compromised
    Microsoft 365 accounts. It can connect to Microsoft Graph and Exchange Online,
    process one or more target users, and perform a controlled set of review and
    remediation actions such as disabling the user, revoking sessions, resetting
    the password, reviewing or removing MFA methods, disabling inbox rules,
    reviewing or removing mailbox forwarding, disabling Outlook on the web
    signature settings, reviewing mailbox delegates, and exporting unified audit,
    sign-in, geolocation, and message trace activity.

    The script is parameter-driven, supports WhatIf for destructive actions,
    uses stronger random password generation, and avoids writing plaintext
    passwords to disk unless explicitly requested in the result output.

.PARAMETER TenantId
    Microsoft Entra tenant ID used for app-based authentication.

.PARAMETER ClientId
    Application (client) ID used for app-based authentication.

.PARAMETER CertificateThumbprint
    Certificate thumbprint used for app-based authentication.

.PARAMETER CsvPath
    Path to a CSV file containing a UserPrincipalName column.

.PARAMETER UserPrincipalName
    One or more Microsoft 365 user principal names to review or remediate.

.PARAMETER Actions
    Actions to perform for the target users.

    Available values:
      - DisableUser
      - RevokeSessions
      - ResetPassword
      - ReviewMfaMethods
      - RemoveMfaMethods
      - DisableInboxRules
      - ReviewMailboxForwarding
      - RemoveMailboxForwarding
      - DisableSignature
      - ReviewMailboxDelegates
      - ExportAuditLog

.PARAMETER OutputPath
    Folder where status files, error logs, and audit log exports are written.

.PARAMETER AuditLogDays
    Number of days of unified audit log, sign-in, and message trace data to export.

.PARAMETER InstallMissingModules
    Installs missing Microsoft Graph and Exchange Online PowerShell modules for the current user.

.PARAMETER IncludeGeneratedPasswordsInResults
    Includes generated passwords in the result output. Use with care.

.EXAMPLE
    .\M365-CompromisedAccountRemediation.ps1 -UserPrincipalName user@domain.com -WhatIf

.EXAMPLE
    .\M365-CompromisedAccountRemediation.ps1 `
        -UserPrincipalName user@domain.com `
        -Actions DisableUser,RevokeSessions,ResetPassword,DisableInboxRules,ReviewMfaMethods `
        -OutputPath C:\IR

.EXAMPLE
    .\M365-CompromisedAccountRemediation.ps1 `
        -CsvPath C:\IR\users.csv `
        -Actions DisableUser,RevokeSessions,ResetPassword,RemoveMfaMethods,RemoveMailboxForwarding,DisableSignature,ExportAuditLog `
        -OutputPath C:\IR
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$TenantId,
    [Alias('CSVFilePath')]
    [string]$CsvPath,
    [Alias('UPNs')]
    [string[]]$UserPrincipalName,
    [ValidateSet(
        'DisableUser',
        'RevokeSessions',
        'ResetPassword',
        'ReviewMfaMethods',
        'RemoveMfaMethods',
        'ReviewInboxRules',
        'DisableInboxRules',
        'ReviewMailboxForwarding',
        'RemoveMailboxForwarding',
        'DisableSignature',
        'ReviewMailboxDelegates',
        'RemoveMailboxDelegates',
        'ReviewRecentSignIns',
        'DisableMailboxProtocols',
        'ExportAuditLog'
    )]
    [string[]]$Actions = @(
        'DisableUser',
        'RevokeSessions',
        'ResetPassword',
        'RemoveMfaMethods',
        'DisableInboxRules',
        'RemoveMailboxForwarding',
        'DisableMailboxProtocols',
        'DisableSignature',
        'ExportAuditLog'
    ),
    [int]$AuditLogDays = 10,
    [string]$OutputPath = (Get-Location).Path,
    [string]$ExportHtml,
    [switch]$IncludeGeneratedPasswordsInResults,
    [switch]$ExportIncidentPackage,
    [switch]$InstallMissingModules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Errors = [System.Collections.Generic.List[string]]::new()
$script:PasswordResults = [System.Collections.Generic.List[psobject]]::new()
$script:Timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$script:TenantLabel = ''
$script:RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Reports',
    'ExchangeOnlineManagement'
)
$script:ReviewActionSet = @(
    'ReviewMfaMethods',
    'ReviewInboxRules',
    'ReviewMailboxForwarding',
    'ReviewMailboxDelegates',
    'ReviewRecentSignIns',
    'ExportAuditLog'
)
$script:HighImpactActionSet = @(
    'DisableUser',
    'ResetPassword',
    'RemoveMfaMethods',
    'RemoveMailboxForwarding',
    'RemoveMailboxDelegates',
    'DisableMailboxProtocols'
)

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor Cyan
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('=' * 70) -ForegroundColor Cyan
}

function Add-ErrorEntry {
    param(
        [string]$User,
        [string]$Action,
        [string]$Message
    )

    $entry = '{0} [{1}] {2} - {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $User, $Action, $Message
    $script:Errors.Add($entry)
}

function New-ActionResult {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Success', 'WhatIf', 'Skipped', 'NotApplicable', 'Partial', 'Failed')]
        [string]$Status,
        [string]$Details
    )

    if ($Details) {
        return ('{0}: {1}' -f $Status, $Details)
    }

    return $Status
}

function Get-ActionStatus {
    param([string]$Value)

    if (-not $Value) {
        return ''
    }

    if ($Value -match '^(Success|WhatIf|Skipped|NotApplicable|Partial|Failed)(?::|\b)') {
        return $Matches[1]
    }

    return $Value
}

function Get-ActionDetails {
    param([string]$Value)

    if (-not $Value) {
        return ''
    }

    if ($Value -match '^(Success|WhatIf|Skipped|NotApplicable|Partial|Failed):\s*(.+)$') {
        return $Matches[2]
    }

    return $Value
}

function Initialize-OutputPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Add-TimestampToPath {
    param(
        [string]$Path,
        [string]$BaseName,
        [string]$Extension = '.html'
    )

    $stamp = Get-Date -Format 'dd.MM.yy-HH.mm.ss'
    $tenant = if ($script:TenantLabel) { $script:TenantLabel.Split('.')[0] } else { 'M365' }
    $fileName = '{0}_{1}_{2}{3}' -f $BaseName, $tenant, $stamp, $Extension

    if (-not $Path) {
        return Join-Path $OutputPath $fileName
    }

    if (Test-Path $Path -PathType Container) {
        return Join-Path $Path $fileName
    }

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if ($directory) {
        return Join-Path $directory $fileName
    }

    return $fileName
}

function Ensure-Modules {
    param([switch]$AllowInstall)

    Write-Host ''
    Write-Host '[*] Checking required PowerShell modules...' -ForegroundColor Cyan

    $graphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Reports'
    )

    $commonVersions = $null
    foreach ($moduleName in $graphModules) {
        $installed = @(Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending)
        if (-not $installed -or $installed.Count -eq 0) {
            throw "Required module '$moduleName' is not installed."
        }

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
        throw "No common Microsoft Graph module version is installed across: $($graphModules -join ', ')."
    }

    $selectedGraphVersion = $commonVersions |
        ForEach-Object { [version]$_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    foreach ($moduleName in $graphModules) {
        Import-Module $moduleName -RequiredVersion $selectedGraphVersion -Force -WarningAction SilentlyContinue
        Write-Host "  [+] $moduleName v$selectedGraphVersion" -ForegroundColor Green
    }

    $exoModule = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $exoModule) {
        throw "Required module 'ExchangeOnlineManagement' is not installed."
    }

    Import-Module ExchangeOnlineManagement -RequiredVersion $exoModule.Version -Force -WarningAction SilentlyContinue
    Write-Host "  [+] ExchangeOnlineManagement v$($exoModule.Version)" -ForegroundColor Green
}

function Connect-Services {
    Write-Section 'CONNECTING TO MICROSOFT 365'

    try {
        if ($TenantId) {
            Write-Host "[*] Requested tenant: $TenantId"
        }
        else {
            Write-Host '[*] Requested tenant: auto-detect'
        }

        Write-Host '[*] Starting device code authentication...' -ForegroundColor Yellow
        Write-Host '[*] When the code appears, open https://login.microsoft.com/device' -ForegroundColor Yellow

        $scopes = @(
            'Directory.AccessAsUser.All',
            'Directory.Read.All',
            'User.ReadWrite.All',
            'UserAuthenticationMethod.ReadWrite.All',
            'AuditLog.Read.All',
            'Reports.Read.All'
        )

        $graphConnectParams = @{
            Scopes       = $scopes
            NoWelcome    = $true
            ContextScope = 'Process'
            ErrorAction  = 'Stop'
        }

        if ($TenantId) {
            $graphConnectParams.TenantId = $TenantId
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph @graphConnectParams -UseDeviceAuthentication

        $context = Get-MgContext
        if ($context -and $context.TenantId) {
            $script:TenantLabel = [string]$context.TenantId
        }

        try {
            $initialDomain = Get-MgDomain -All -ErrorAction Stop | Where-Object { $_.IsInitial } | Select-Object -First 1
            if ($initialDomain -and $initialDomain.Id) {
                $script:TenantLabel = [string]$initialDomain.Id
            }
        }
        catch {
        }

        Write-Host '[*] Connecting to Exchange Online with device authentication...' -ForegroundColor Cyan
        if ($TenantId) {
            Connect-ExchangeOnline -Device -Organization $TenantId -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
        }
    }
    catch {
        throw "Failed to connect to Microsoft Graph and Exchange Online. $($_.Exception.Message)"
    }

    Write-Host '[+] Microsoft Graph connected successfully.' -ForegroundColor Green
    Write-Host '[+] Exchange Online connected successfully.' -ForegroundColor Green
}

function Disconnect-Services {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

function Resolve-TargetUsers {
    $targets = [System.Collections.Generic.List[string]]::new()

    if ($CsvPath) {
        $rows = Import-Csv -Path $CsvPath
        foreach ($row in $rows) {
            if ($row.UserPrincipalName) {
                $targets.Add($row.UserPrincipalName.Trim())
            }
        }
    }

    foreach ($upn in @($UserPrincipalName)) {
        if ($upn -and $upn.Trim()) {
            $targets.Add($upn.Trim())
        }
    }

    $resolved = @($targets | Where-Object { $_ } | Select-Object -Unique)
    if ($resolved.Count -eq 0) {
        throw 'Provide at least one target via -UserPrincipalName or -CsvPath.'
    }

    return $resolved
}

function New-StrongPassword {
    param([int]$Length = 20)

    if ($Length -lt 12) {
        throw 'Password length must be at least 12 characters.'
    }

    $lower   = 'abcdefghijkmnopqrstuvwxyz'.ToCharArray()
    $upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $digits  = '23456789'.ToCharArray()
    $special = '!@#$%^&*()-_=+[]{}:,.?'.ToCharArray()
    $all     = $lower + $upper + $digits + $special

    $chars = [System.Collections.Generic.List[char]]::new()
    $chars.Add(($lower | Get-Random))
    $chars.Add(($upper | Get-Random))
    $chars.Add(($digits | Get-Random))
    $chars.Add(($special | Get-Random))

    for ($i = $chars.Count; $i -lt $Length; $i++) {
        $chars.Add(($all | Get-Random))
    }

    return (-join ($chars | Sort-Object { Get-Random }))
}

function Get-AuthenticationMethodSegment {
    param([string]$ODataType)

    switch ($ODataType) {
        '#microsoft.graph.emailAuthenticationMethod' { return 'emailMethods' }
        '#microsoft.graph.fido2AuthenticationMethod' { return 'fido2Methods' }
        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { return 'microsoftAuthenticatorMethods' }
        '#microsoft.graph.phoneAuthenticationMethod' { return 'phoneMethods' }
        '#microsoft.graph.softwareOathAuthenticationMethod' { return 'softwareOathMethods' }
        '#microsoft.graph.temporaryAccessPassAuthenticationMethod' { return 'temporaryAccessPassMethods' }
        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { return 'windowsHelloForBusinessMethods' }
        default { return $null }
    }
}

function Get-UserContext {
    param([string]$Upn)

    $user = Get-MgUser -UserId $Upn -ErrorAction SilentlyContinue
    if (-not $user) {
        return [pscustomobject]@{
            Exists     = $false
            Upn        = $Upn
            User       = $null
            Mailbox    = $null
            HasMailbox = $false
            IsGuest    = $false
            IsSynced   = $false
        }
    }

    $mailbox = Get-Mailbox -Identity $Upn -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue

    return [pscustomobject]@{
        Exists     = $true
        Upn        = $Upn
        User       = $user
        Mailbox    = $mailbox
        HasMailbox = ($null -ne $mailbox)
        IsGuest    = ($user.UserType -eq 'Guest')
        IsSynced   = [bool]$user.OnPremisesSyncEnabled
    }
}

function Invoke-DisableUser {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    try {
        if ($PSCmdlet.ShouldProcess($Context.Upn, 'Disable Azure AD account')) {
            Update-MgUser -UserId $Context.Upn -AccountEnabled:$false
            return (New-ActionResult -Status 'Success')
        }

        return (New-ActionResult -Status 'WhatIf')
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'DisableUser' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-RevokeSessions {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    try {
        if ($PSCmdlet.ShouldProcess($Context.Upn, 'Revoke all sign-in sessions')) {
            Revoke-MgUserSignInSession -UserId $Context.Upn | Out-Null
            return (New-ActionResult -Status 'Success')
        }

        return (New-ActionResult -Status 'WhatIf')
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'RevokeSessions' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-ResetPassword {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [psobject]$Context,
        [switch]$IncludePassword
    )

    try {
        $password = New-StrongPassword
        $passwordProfile = @{
            forceChangePasswordNextSignIn = $true
            password = $password
        }

        if ($PSCmdlet.ShouldProcess($Context.Upn, 'Reset password to a strong random value')) {
            Update-MgUser -UserId $Context.Upn -PasswordProfile $passwordProfile
            $script:PasswordResults.Add([pscustomobject]@{
                UserPrincipalName = $Context.Upn
                Password          = if ($IncludePassword) { $password } else { '[suppressed]' }
                ChangedAt         = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            })

            return [pscustomobject]@{
                Status   = (New-ActionResult -Status 'Success')
                Password = if ($IncludePassword) { $password } else { '[suppressed]' }
            }
        }

        return [pscustomobject]@{
            Status   = (New-ActionResult -Status 'WhatIf')
            Password = '[not generated]'
        }
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ResetPassword' -Message $_.Exception.Message
        return [pscustomobject]@{
            Status   = (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
            Password = $null
        }
    }
}

function Get-MfaReview {
    param([psobject]$Context)

    try {
        $methods = @(Get-MgUserAuthenticationMethod -UserId $Context.Upn)
        $nonPasswordMethods = @(
            $methods |
                Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.passwordAuthenticationMethod' } |
                ForEach-Object { $_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', '' }
        )

        if ($nonPasswordMethods.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No non-password authentication methods are registered.')
        }

        return (New-ActionResult -Status 'Success' -Details ('Registered methods: ' + ($nonPasswordMethods -join ', ')))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ReviewMfaMethods' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-RemoveMfaMethods {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    try {
        $methods = @(
            Get-MgUserAuthenticationMethod -UserId $Context.Upn |
                Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.passwordAuthenticationMethod' }
        )

        if ($methods.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No removable non-password methods found.')
        }

        $results = [System.Collections.Generic.List[string]]::new()
        foreach ($method in $methods) {
            $odataType = [string]$method.AdditionalProperties['@odata.type']
            $segment = Get-AuthenticationMethodSegment -ODataType $odataType
            if (-not $segment) {
                $results.Add("Skipped unsupported method type $odataType")
                continue
            }

            $description = '{0} ({1})' -f $segment, $method.Id
            if ($PSCmdlet.ShouldProcess($Context.Upn, "Remove authentication method $description")) {
                Invoke-MgGraphRequest -Method DELETE -Uri ("https://graph.microsoft.com/v1.0/users/{0}/authentication/{1}/{2}" -f $Context.Upn, $segment, $method.Id) | Out-Null
                $results.Add("Removed $segment")
            }
            else {
                $results.Add("WhatIf $segment")
            }
        }

        $status = if ($results -match '^Skipped ') { 'Partial' } else { 'Success' }
        return (New-ActionResult -Status $status -Details ($results -join '; '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'RemoveMfaMethods' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Get-InboxRuleReview {
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        $rules = @(Get-InboxRule -Mailbox $Context.Upn -ErrorAction Stop | Where-Object { $_.Enabled })
        if ($rules.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No enabled inbox rules found.')
        }

        $descriptions = foreach ($rule in $rules) {
            $targets = @($rule.ForwardTo + $rule.ForwardAsAttachmentTo + $rule.RedirectTo | Where-Object { $_ })
            if ($targets.Count -gt 0) {
                '{0} -> {1}' -f $rule.Name, ($targets -join ', ')
            }
            else {
                [string]$rule.Name
            }
        }

        return (New-ActionResult -Status 'Success' -Details ($descriptions -join '; '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ReviewInboxRules' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-DisableInboxRules {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        $rules = @(Get-InboxRule -Mailbox $Context.Upn | Where-Object { $_.Enabled })
        if ($rules.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No enabled inbox rules found.')
        }

        $ruleNames = @($rules.Name)
        $processed = 0
        foreach ($rule in $rules) {
            if ($PSCmdlet.ShouldProcess($Context.Upn, "Disable inbox rule '$($rule.Name)'")) {
                Disable-InboxRule -Identity $rule.Identity -Confirm:$false
                $processed++
            }
        }

        if ($processed -eq 0) {
            return (New-ActionResult -Status 'WhatIf' -Details ($ruleNames -join ', '))
        }

        $status = if ($processed -lt $rules.Count) { 'Partial' } else { 'Success' }
        return (New-ActionResult -Status $status -Details ($ruleNames -join ', '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'DisableInboxRules' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Get-MailboxForwardingReview {
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        $mailbox = $Context.Mailbox
        $details = [System.Collections.Generic.List[string]]::new()

        if ($mailbox.ForwardingSmtpAddress) {
            $details.Add("ForwardingSmtpAddress=$($mailbox.ForwardingSmtpAddress)")
        }
        if ($mailbox.ForwardingAddress) {
            $details.Add("ForwardingAddress=$($mailbox.ForwardingAddress)")
        }
        if ($mailbox.DeliverToMailboxAndForward -ne $null) {
            $details.Add("DeliverToMailboxAndForward=$($mailbox.DeliverToMailboxAndForward)")
        }

        if ($details.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No mailbox forwarding configuration found.')
        }

        return (New-ActionResult -Status 'Success' -Details ($details -join '; '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ReviewMailboxForwarding' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-RemoveMailboxForwarding {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        if ($PSCmdlet.ShouldProcess($Context.Upn, 'Remove mailbox forwarding configuration')) {
            Set-Mailbox -Identity $Context.Upn -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward:$false
            return (New-ActionResult -Status 'Success')
        }

        return (New-ActionResult -Status 'WhatIf')
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'RemoveMailboxForwarding' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-DisableSignature {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        if ($PSCmdlet.ShouldProcess($Context.Upn, 'Disable mailbox signature settings')) {
            Set-MailboxMessageConfiguration `
                -Identity $Context.Upn `
                -AutoAddSignature $false `
                -AutoAddSignatureOnMobile $false `
                -AutoAddSignatureOnReply $false `
                -SignatureHtml '' `
                -SignatureText '' `
                -SignatureTextOnMobile '' `
                -DefaultSignature '' `
                -DefaultSignatureOnReply '' `
                -ErrorAction Stop

            return (New-ActionResult -Status 'Success')
        }

        return (New-ActionResult -Status 'WhatIf')
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'DisableSignature' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Get-MailboxDelegateReview {
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        $details = [System.Collections.Generic.List[string]]::new()

        $fullAccess = @(
            Get-MailboxPermission -Identity $Context.Upn |
                Where-Object {
                    -not $_.IsInherited -and
                    $_.User -notmatch 'NT AUTHORITY|S-1-5-|Discovery Management|Organization Management|Exchange Servers|Exchange Trusted Subsystem|SELF'
                }
        )
        if ($fullAccess.Count -gt 0) {
            $details.Add('FullAccess=' + (($fullAccess | Select-Object -ExpandProperty User) -join ', '))
        }

        $sendAs = @(
            Get-RecipientPermission -Identity $Context.Upn -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Trustee -and
                    $_.Trustee -notmatch 'NT AUTHORITY|S-1-5-|Discovery Management|Organization Management'
                }
        )
        if ($sendAs.Count -gt 0) {
            $details.Add('SendAs=' + (($sendAs | Select-Object -ExpandProperty Trustee) -join ', '))
        }

        $sendOnBehalf = @($Context.Mailbox.GrantSendOnBehalfTo)
        if ($sendOnBehalf.Count -gt 0) {
            $details.Add('SendOnBehalf=' + ($sendOnBehalf -join ', '))
        }

        if ($details.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No mailbox delegates found beyond default/system entries.')
        }

        return (New-ActionResult -Status 'Success' -Details ($details -join '; '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ReviewMailboxDelegates' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-RemoveMailboxDelegates {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        $removed = [System.Collections.Generic.List[string]]::new()

        $fullAccess = @(
            Get-MailboxPermission -Identity $Context.Upn |
                Where-Object {
                    -not $_.IsInherited -and
                    $_.User -notmatch 'NT AUTHORITY|S-1-5-|Discovery Management|Organization Management|Exchange Servers|Exchange Trusted Subsystem|SELF'
                }
        )

        foreach ($permission in $fullAccess) {
            if ($PSCmdlet.ShouldProcess($Context.Upn, "Remove FullAccess delegate '$($permission.User)'")) {
                Remove-MailboxPermission -Identity $Context.Upn -User $permission.User -AccessRights FullAccess -InheritanceType All -Confirm:$false
                $removed.Add("FullAccess=$($permission.User)")
            }
        }

        $sendAs = @(
            Get-RecipientPermission -Identity $Context.Upn -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Trustee -and
                    $_.Trustee -notmatch 'NT AUTHORITY|S-1-5-|Discovery Management|Organization Management'
                }
        )

        foreach ($permission in $sendAs) {
            if ($PSCmdlet.ShouldProcess($Context.Upn, "Remove SendAs delegate '$($permission.Trustee)'")) {
                Remove-RecipientPermission -Identity $Context.Upn -Trustee $permission.Trustee -AccessRights SendAs -Confirm:$false
                $removed.Add("SendAs=$($permission.Trustee)")
            }
        }

        foreach ($delegate in @($Context.Mailbox.GrantSendOnBehalfTo)) {
            if ($PSCmdlet.ShouldProcess($Context.Upn, "Remove SendOnBehalf delegate '$delegate'")) {
                Set-Mailbox -Identity $Context.Upn -GrantSendOnBehalfTo @{ remove = $delegate } -ErrorAction Stop
                $removed.Add("SendOnBehalf=$delegate")
            }
        }

        if ($removed.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No removable mailbox delegates found.')
        }

        $status = if ($WhatIfPreference) { 'WhatIf' } else { 'Success' }
        return (New-ActionResult -Status $status -Details ($removed -join '; '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'RemoveMailboxDelegates' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Get-RecentSignInReview {
    param([psobject]$Context)

    try {
        $escapedUpn = $Context.Upn.Replace("'", "''")
        $filter = "userPrincipalName eq '$escapedUpn'"
        $signIns = @(
            Get-MgAuditLogSignIn -Filter $filter -Top 12 -ErrorAction Stop |
                Sort-Object CreatedDateTime -Descending |
                Select-Object -First 5
        )

        if ($signIns.Count -eq 0) {
            return (New-ActionResult -Status 'NotApplicable' -Details 'No recent sign-ins found.')
        }

        $summary = foreach ($signIn in $signIns) {
            $location = Get-GeoLocationString -Location $signIn.Location
            $ipAddress = if ($signIn.IpAddress) { [string]$signIn.IpAddress } else { 'No IP' }
            $locationLabel = if ($location) { $location } else { 'No location' }
            $clientLabel = if ($signIn.ClientAppUsed) { [string]$signIn.ClientAppUsed } else { 'Unknown client' }
            $appLabel = if ($signIn.AppDisplayName) { [string]$signIn.AppDisplayName } else { 'Unknown app' }
            '{0} | {1} | {2} | {3} | {4}' -f (
                [string]$signIn.CreatedDateTime,
                $ipAddress,
                $locationLabel,
                $clientLabel,
                $appLabel
            )
        }

        return (New-ActionResult -Status 'Success' -Details ($summary -join '; '))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ReviewRecentSignIns' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Invoke-DisableMailboxProtocols {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([psobject]$Context)

    if (-not $Context.HasMailbox) {
        return (New-ActionResult -Status 'NotApplicable' -Details 'No Exchange mailbox found.')
    }

    try {
        if ($PSCmdlet.ShouldProcess($Context.Upn, 'Disable OWA, ActiveSync, POP, IMAP, and MAPI mailbox protocols')) {
            Set-CASMailbox `
                -Identity $Context.Upn `
                -OWAEnabled:$false `
                -ActiveSyncEnabled:$false `
                -PopEnabled:$false `
                -ImapEnabled:$false `
                -MAPIEnabled:$false `
                -OWAforDevicesEnabled:$false `
                -ErrorAction Stop

            return (New-ActionResult -Status 'Success' -Details 'OWA, ActiveSync, POP, IMAP, MAPI, and OWA for Devices disabled.')
        }

        return (New-ActionResult -Status 'WhatIf' -Details 'OWA, ActiveSync, POP, IMAP, MAPI, and OWA for Devices would be disabled.')
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'DisableMailboxProtocols' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function Get-SafeFileBaseName {
    param([string]$Value)

    return ($Value -replace '[^a-zA-Z0-9._-]', '_')
}

function Get-IncidentPackagePath {
    param(
        [string]$DestinationPath,
        [string]$Upn
    )

    $safeName = Get-SafeFileBaseName -Value $Upn
    $packagePath = Join-Path $DestinationPath ("IncidentPackage_{0}" -f $safeName)
    if (-not (Test-Path -LiteralPath $packagePath)) {
        New-Item -ItemType Directory -Path $packagePath -Force | Out-Null
    }

    return $packagePath
}

function Export-IncidentPackageSummary {
    param(
        [psobject]$Context,
        [hashtable]$Row,
        [string]$DestinationPath
    )

    $packagePath = Get-IncidentPackagePath -DestinationPath $DestinationPath -Upn $Context.Upn
    $summaryPath = Join-Path $packagePath ("Summary_{0}.json" -f $script:Timestamp)
    $summary = [pscustomobject]@{
        UserPrincipalName = $Context.Upn
        UserFound = $Row.UserFound
        MailboxFound = $Row.MailboxFound
        IsGuest = $Row.IsGuest
        IsSynced = $Row.IsSynced
        ReviewActions = @{
            ReviewMfaMethods = $Row.ReviewMfaMethods
            ReviewInboxRules = $Row.ReviewInboxRules
            ReviewMailboxForwarding = $Row.ReviewMailboxForwarding
            ReviewMailboxDelegates = $Row.ReviewMailboxDelegates
            ReviewRecentSignIns = $Row.ReviewRecentSignIns
        }
        RemediationActions = @{
            DisableUser = $Row.DisableUser
            RevokeSessions = $Row.RevokeSessions
            ResetPassword = $Row.ResetPassword
            RemoveMfaMethods = $Row.RemoveMfaMethods
            DisableInboxRules = $Row.DisableInboxRules
            RemoveMailboxForwarding = $Row.RemoveMailboxForwarding
            RemoveMailboxDelegates = $Row.RemoveMailboxDelegates
            DisableSignature = $Row.DisableSignature
            DisableMailboxProtocols = $Row.DisableMailboxProtocols
        }
        ExportAuditLog = $Row.ExportAuditLog
    }

    $summary | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $summaryPath -Encoding utf8
    return $packagePath
}

function Get-TargetDetailsString {
    param($AuditInfo)

    $targets = @()
    foreach ($name in @('Target', 'TargetUserOrGroupName', 'ObjectId', 'ItemName', 'DestFolder', 'MailboxOwnerUPN')) {
        if ($AuditInfo.PSObject.Properties.Name -contains $name -and $AuditInfo.$name) {
            $targets += [string]$AuditInfo.$name
        }
    }

    if ($AuditInfo.PSObject.Properties.Name -contains 'Targets' -and $AuditInfo.Targets) {
        foreach ($target in @($AuditInfo.Targets)) {
            if ($target.PSObject.Properties.Name -contains 'Id' -and $target.Id) { $targets += [string]$target.Id }
            if ($target.PSObject.Properties.Name -contains 'DisplayName' -and $target.DisplayName) { $targets += [string]$target.DisplayName }
            if ($target.PSObject.Properties.Name -contains 'Type' -and $target.Type) { $targets += [string]$target.Type }
        }
    }

    return (($targets | Where-Object { $_ } | Select-Object -Unique) -join '; ')
}

function Get-GeoLocationString {
    param($Location)

    if (-not $Location) {
        return ''
    }

    $parts = @()
    foreach ($name in @('City', 'State', 'CountryOrRegion')) {
        if ($Location.PSObject.Properties.Name -contains $name -and $Location.$name) {
            $parts += [string]$Location.$name
        }
    }

    return ($parts -join ', ')
}

function Export-UnifiedAuditLogReport {
    param(
        [psobject]$Context,
        [string]$DestinationPath,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    $safeName = Get-SafeFileBaseName -Value $Context.Upn
    $outputCsv = Join-Path $DestinationPath ("{0}_UnifiedAudit_{1}.csv" -f $safeName, $script:Timestamp)

    $records = Search-UnifiedAuditLog -UserIds $Context.Upn -ResultSize 5000 -SessionCommand ReturnLargeSet -StartDate $StartDate -EndDate $EndDate
    $normalized = foreach ($record in $records) {
        $auditInfo = $record.AuditData | ConvertFrom-Json
        [pscustomobject]@{
            ActivityTime   = $record.CreationDate
            UserName       = $record.UserIds
            Operation      = $record.Operations
            Result         = $auditInfo.ResultStatus
            Workload       = $auditInfo.Workload
            ClientIP       = if ($auditInfo.ClientIP) { $auditInfo.ClientIP } else { $auditInfo.ClientIPAddress }
            UserAgent      = $auditInfo.UserAgent
            AppId          = $auditInfo.AppId
            ObjectId       = $auditInfo.ObjectId
            OperationCount = $auditInfo.OperationCount
            LogonType      = $auditInfo.LogonType
            ExternalAccess = $auditInfo.ExternalAccess
            TargetDetails  = Get-TargetDetailsString -AuditInfo $auditInfo
            MoreInfo       = $record.AuditData
        }
    }

    if (@($normalized).Count -gt 0) {
        $normalized | Export-Csv -Path $outputCsv -NoTypeInformation
    }

    return [pscustomobject]@{
        Count = @($normalized).Count
        Path  = if (@($normalized).Count -gt 0) { $outputCsv } else { $null }
    }
}

function Export-SignInReports {
    param(
        [psobject]$Context,
        [string]$DestinationPath,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    $safeName = Get-SafeFileBaseName -Value $Context.Upn
    $escapedUpn = $Context.Upn.Replace("'", "''")
    $startUtc = $StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $endUtc = $EndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $filter = "userPrincipalName eq '$escapedUpn' and createdDateTime ge $startUtc and createdDateTime le $endUtc"
    $signIns = @(Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop)

    $normalized = foreach ($signIn in $signIns) {
        [pscustomobject]@{
            CreatedDateTime   = $signIn.CreatedDateTime
            UserPrincipalName = $signIn.UserPrincipalName
            UserDisplayName   = $signIn.UserDisplayName
            AppDisplayName    = $signIn.AppDisplayName
            AppId             = $signIn.AppId
            ClientAppUsed     = $signIn.ClientAppUsed
            UserAgent         = $signIn.UserAgent
            IPAddress         = $signIn.IpAddress
            InteractiveSignIn = $signIn.IsInteractive
            StatusCode        = if ($signIn.Status) { $signIn.Status.ErrorCode } else { $null }
            FailureReason     = if ($signIn.Status) { $signIn.Status.FailureReason } else { $null }
            ConditionalAccess = $signIn.ConditionalAccessStatus
            RiskLevel         = $signIn.RiskLevelDuringSignIn
            RiskState         = $signIn.RiskState
            Geolocation       = Get-GeoLocationString -Location $signIn.Location
            Country           = if ($signIn.Location) { $signIn.Location.CountryOrRegion } else { $null }
            State             = if ($signIn.Location) { $signIn.Location.State } else { $null }
            City              = if ($signIn.Location) { $signIn.Location.City } else { $null }
            Latitude          = if ($signIn.Location -and $signIn.Location.GeoCoordinates) { $signIn.Location.GeoCoordinates.Latitude } else { $null }
            Longitude         = if ($signIn.Location -and $signIn.Location.GeoCoordinates) { $signIn.Location.GeoCoordinates.Longitude } else { $null }
            ResourceDisplay   = $signIn.ResourceDisplayName
            CorrelationId     = $signIn.CorrelationId
        }
    }

    $interactive = @($normalized | Where-Object { $_.InteractiveSignIn -eq $true })
    $nonInteractive = @($normalized | Where-Object { $_.InteractiveSignIn -ne $true })
    $interactivePath = Join-Path $DestinationPath ("{0}_InteractiveSignIns_{1}.csv" -f $safeName, $script:Timestamp)
    $nonInteractivePath = Join-Path $DestinationPath ("{0}_NonInteractiveSignIns_{1}.csv" -f $safeName, $script:Timestamp)

    if ($interactive.Count -gt 0) {
        $interactive | Export-Csv -Path $interactivePath -NoTypeInformation
    }
    if ($nonInteractive.Count -gt 0) {
        $nonInteractive | Export-Csv -Path $nonInteractivePath -NoTypeInformation
    }

    return [pscustomobject]@{
        InteractiveCount    = $interactive.Count
        InteractivePath     = if ($interactive.Count -gt 0) { $interactivePath } else { $null }
        NonInteractiveCount = $nonInteractive.Count
        NonInteractivePath  = if ($nonInteractive.Count -gt 0) { $nonInteractivePath } else { $null }
    }
}

function Export-MessageTraceReports {
    param(
        [psobject]$Context,
        [string]$DestinationPath,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    $safeName = Get-SafeFileBaseName -Value $Context.Upn
    $outgoing = @(Get-MessageTrace -SenderAddress $Context.Upn -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -ErrorAction Stop)
    $incoming = @(Get-MessageTrace -RecipientAddress $Context.Upn -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -ErrorAction Stop)

    $outgoingPath = Join-Path $DestinationPath ("{0}_OutgoingMessageTrace_{1}.csv" -f $safeName, $script:Timestamp)
    $incomingPath = Join-Path $DestinationPath ("{0}_IncomingMessageTrace_{1}.csv" -f $safeName, $script:Timestamp)

    if ($outgoing.Count -gt 0) {
        $outgoing | Select-Object Received,SenderAddress,RecipientAddress,Subject,Status,MessageTraceId,Size,FromIP,ToIP |
            Export-Csv -Path $outgoingPath -NoTypeInformation
    }
    if ($incoming.Count -gt 0) {
        $incoming | Select-Object Received,SenderAddress,RecipientAddress,Subject,Status,MessageTraceId,Size,FromIP,ToIP |
            Export-Csv -Path $incomingPath -NoTypeInformation
    }

    return [pscustomobject]@{
        OutgoingCount = $outgoing.Count
        OutgoingPath  = if ($outgoing.Count -gt 0) { $outgoingPath } else { $null }
        IncomingCount = $incoming.Count
        IncomingPath  = if ($incoming.Count -gt 0) { $incomingPath } else { $null }
    }
}

function Export-ActivityLog {
    param(
        [psobject]$Context,
        [string]$DestinationPath,
        [int]$Days
    )

    try {
        $endDate = (Get-Date).AddSeconds(-1)
        $startDate = (Get-Date).AddDays(-1 * $Days)

        $auditResult = Export-UnifiedAuditLogReport -Context $Context -DestinationPath $DestinationPath -StartDate $startDate -EndDate $endDate
        $signInResult = Export-SignInReports -Context $Context -DestinationPath $DestinationPath -StartDate $startDate -EndDate $endDate
        $messageTraceResult = Export-MessageTraceReports -Context $Context -DestinationPath $DestinationPath -StartDate $startDate -EndDate $endDate

        return (New-ActionResult -Status 'Success' -Details ('UnifiedAudit={0}; InteractiveSignIns={1}; NonInteractiveSignIns={2}; OutgoingMail={3}; IncomingMail={4}' -f `
            $auditResult.Count,
            $signInResult.InteractiveCount,
            $signInResult.NonInteractiveCount,
            $messageTraceResult.OutgoingCount,
            $messageTraceResult.IncomingCount))
    }
    catch {
        Add-ErrorEntry -User $Context.Upn -Action 'ExportAuditLog' -Message $_.Exception.Message
        return (New-ActionResult -Status 'Failed' -Details $_.Exception.Message)
    }
}

function New-ResultRow {
    param([string]$Upn)

    return [ordered]@{
        UserPrincipalName       = $Upn
        UserFound               = $false
        MailboxFound            = $false
        IsGuest                 = $false
        IsSynced                = $false
        DisableUser             = ''
        RevokeSessions          = ''
        ResetPassword           = ''
        GeneratedPassword       = ''
        ReviewMfaMethods        = ''
        RemoveMfaMethods        = ''
        ReviewInboxRules        = ''
        DisableInboxRules       = ''
        ReviewMailboxForwarding = ''
        RemoveMailboxForwarding = ''
        DisableSignature        = ''
        ReviewMailboxDelegates  = ''
        RemoveMailboxDelegates  = ''
        ReviewRecentSignIns     = ''
        DisableMailboxProtocols = ''
        ExportAuditLog          = ''
        IncidentPackagePath     = ''
    }
}

function Export-HtmlReport {
    param(
        [Parameter(Mandatory)]
        [object[]]$Results,
        [Parameter(Mandatory)]
        [string]$Path
    )

    Write-Host '[*] Building HTML dashboard...' -ForegroundColor Cyan

    $totalTargets = $Results.Count
    $usersFound = @($Results | Where-Object { $_.UserFound }).Count
    $disabled = @($Results | Where-Object { (Get-ActionStatus $_.DisableUser) -eq 'Success' }).Count
    $sessionsRevoked = @($Results | Where-Object { (Get-ActionStatus $_.RevokeSessions) -eq 'Success' }).Count
    $passwordReset = @($Results | Where-Object { (Get-ActionStatus $_.ResetPassword) -eq 'Success' }).Count
    $mfaRemoved = @($Results | Where-Object { (Get-ActionStatus $_.RemoveMfaMethods) -in @('Success', 'Partial') }).Count
    $auditExported = @($Results | Where-Object { (Get-ActionStatus $_.ExportAuditLog) -in @('Success', 'Partial') }).Count
    $mailboxNotFound = @($Results | Where-Object { -not $_.MailboxFound }).Count
    $passwordSuppressed = @($Results | Where-Object {
        (Get-ActionStatus $_.ResetPassword) -eq 'Success' -and $_.GeneratedPassword -eq '[suppressed]'
    }).Count
    $auditFailed = @($Results | Where-Object { (Get-ActionStatus $_.ExportAuditLog) -eq 'Failed' }).Count
    $contained = @($Results | Where-Object {
        ((Get-ActionStatus $_.DisableUser) -eq 'Success' -or (Get-ActionStatus $_.RevokeSessions) -eq 'Success') -and
        @(
            Get-ActionStatus $_.ResetPassword,
            Get-ActionStatus $_.RemoveMfaMethods,
            Get-ActionStatus $_.DisableInboxRules,
            Get-ActionStatus $_.RemoveMailboxForwarding,
            Get-ActionStatus $_.DisableMailboxProtocols
        ) -match 'Success|Partial'
    }).Count
    $needsManualFollowUp = @($Results | Where-Object {
        -not $_.UserFound -or
        @(
            $_.DisableUser,
            $_.RevokeSessions,
            $_.ResetPassword,
            $_.RemoveMfaMethods,
            $_.DisableInboxRules,
            $_.RemoveMailboxForwarding,
            $_.RemoveMailboxDelegates,
            $_.DisableMailboxProtocols,
            $_.ExportAuditLog
        ) | ForEach-Object { Get-ActionStatus $_ } | Where-Object { $_ -in @('Failed', 'Partial') }
    }).Count

    $htmlData = [PSCustomObject]@{
        tenant          = if ($script:TenantLabel) { $script:TenantLabel } else { $TenantId }
        reportDate      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        totalTargets    = $totalTargets
        usersFound      = $usersFound
        disabled        = $disabled
        sessionsRevoked = $sessionsRevoked
        passwordReset   = $passwordReset
        mfaRemoved      = $mfaRemoved
        auditExported   = $auditExported
        mailboxNotFound = $mailboxNotFound
        passwordSuppressed = $passwordSuppressed
        auditFailed = $auditFailed
        contained = $contained
        needsManualFollowUp = $needsManualFollowUp
        selectedActions = @($Actions)
        reviewActions = @($script:ReviewActionSet)
        highImpactActions = @($script:HighImpactActionSet)
        errors          = @($script:Errors)
        rows            = @($Results)
    } | ConvertTo-Json -Depth 6 -Compress

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Compromised Account Remediation</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  :root{--bg:#f5f7fb;--bg2:#ffffff;--bg3:#f7f9fc;--border:#d6dfec;--text:#17324d;--text2:#4b6887;--text3:#6f88a3;--accent:#0f7cc0;--ok:#15803d;--warn:#b45309;--crit:#dc2626;--mono:'JetBrains Mono',monospace;--sans:'DM Sans',sans-serif;--r:8px;--r2:12px}
  body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.6}
  .topbar{position:sticky;top:0;z-index:20;background:rgba(245,247,251,.88);backdrop-filter:blur(16px);border-bottom:1px solid var(--border);padding:0 2rem;height:52px;display:flex;align-items:center;gap:1rem}
  .topbar-logo,.topbar-server,.topbar-time,.kpi-value,.strip-value,.pill,th,td{font-family:var(--mono)}
  .topbar-logo{font-size:.72rem;font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:var(--accent)}
  .topbar-server{color:var(--text2);font-size:.78rem}
  .topbar-time{margin-left:auto;color:var(--text3);font-size:.7rem}
  .page{max-width:1440px;margin:0 auto;padding:1.75rem 2rem 2rem}
  .server-strip,.card,.kpi{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2)}
  .server-strip{padding:1rem 1.5rem;display:flex;flex-wrap:wrap;gap:2rem;margin-bottom:1.25rem}
  .strip-item{display:flex;flex-direction:column;gap:.2rem}
  .strip-label,.kpi-label{font-size:.63rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text3)}
  .strip-value{font-size:.82rem;color:var(--accent)}
  .hero{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:.85rem;margin-bottom:1.25rem}
  .kpi{padding:1rem 1.1rem;position:relative;overflow:hidden}
  .kpi::after{content:'';position:absolute;top:0;left:0;right:0;height:2px}
  .kpi.neutral::after{background:var(--accent)}
  .kpi.ok::after{background:var(--ok)}
  .kpi.warn::after{background:var(--warn)}
  .kpi.crit::after{background:var(--crit)}
  .kpi-value{font-size:1.6rem;font-weight:700;line-height:1;color:var(--accent)}
  .kpi-sub{font-size:.68rem;color:var(--text3);margin-top:.35rem}
  .sections{display:grid;gap:1rem}
  .card-header{display:flex;align-items:center;gap:.75rem;padding:.85rem 1.25rem;border-bottom:1px solid var(--border)}
  .card-title{font-size:.72rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text2);flex:1}
  .card-badge{font-size:.7rem;padding:.18rem .55rem;border-radius:4px;background:rgba(15,124,192,.1);color:var(--accent)}
  .card-body{padding:1.25rem}
  .table-scroll{max-height:620px;overflow:auto;border:1px solid var(--border);border-radius:var(--r)}
  table{width:100%;border-collapse:collapse;font-size:.77rem;table-layout:fixed}
  thead{background:var(--bg3);position:sticky;top:0;z-index:1}
  th{position:relative;padding:.55rem .9rem;text-align:left;font-size:.63rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);border-bottom:1px solid var(--border);white-space:nowrap}
  td{padding:.5rem .9rem;border-bottom:1px solid var(--border);color:var(--text);font-size:.76rem;vertical-align:top;word-break:break-word}
  tr:last-child td{border-bottom:none}
  tbody tr:hover td{background:rgba(15,124,192,.05)}
  .pill{display:inline-block;padding:.12rem .45rem;border-radius:4px;font-size:.7rem}
  .ok{background:rgba(21,128,61,.1);color:var(--ok)}
  .warn{background:rgba(180,83,9,.1);color:var(--warn)}
  .crit{background:rgba(220,38,38,.1);color:var(--crit)}
  .neutral{background:rgba(15,124,192,.1);color:var(--accent)}
  .errors{display:grid;gap:.45rem}
  .error-item{background:var(--bg3);border:1px solid var(--border);border-radius:var(--r);padding:.7rem .8rem;font-family:var(--mono);font-size:.72rem;color:#b91c1c}
  .col-resizer{position:absolute;top:0;right:0;width:10px;height:100%;cursor:col-resize;user-select:none;touch-action:none}
  .col-resizer::after{content:'';position:absolute;top:20%;bottom:20%;right:4px;width:2px;border-radius:999px;background:transparent;transition:background .15s ease}
  th:hover .col-resizer::after,.col-resizer.active::after{background:var(--accent)}
</style>
</head>
<body>
<div class="topbar"><div class="topbar-logo">Compromised Account Remediation</div><div class="topbar-server"></div><div class="topbar-time"></div></div>
<div class="page"><div id="server-strip" class="server-strip"></div><div id="hero" class="hero"></div><div id="sections" class="sections"></div></div>
<script>
const DATA = $htmlData;
document.querySelector('.topbar-server').textContent = DATA.tenant;
document.querySelector('.topbar-time').textContent = 'Generated: ' + DATA.reportDate;
function esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function kpi(value,label,sub,cls){return '<div class="kpi ' + cls + '"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>';}
function actionStatus(value){
  const text=String(value==null?'':value);
  const match=text.match(/^(Success|WhatIf|Skipped|NotApplicable|Partial|Failed)(?::|\b)/);
  return match?match[1]:text;
}
function actionDetails(value){
  const text=String(value==null?'':value);
  const match=text.match(/^(Success|WhatIf|Skipped|NotApplicable|Partial|Failed):\s*(.+)$/);
  return match?match[2]:text;
}
function pill(v){
  const raw=String(v==null?'':v);
  if(raw==='true' || raw==='false'){
    return '<span class="pill ' + (raw==='true'?'ok':'crit') + '">' + raw + '</span>';
  }
  const status=actionStatus(raw);
  let cls='neutral';
  if(status==='Success')cls='ok';
  else if(status==='Failed')cls='crit';
  else if(status==='WhatIf' || status==='Skipped' || status==='Partial')cls='warn';
  else if(status==='NotApplicable')cls='neutral';
  return '<span class="pill ' + cls + '">' + esc(status || raw || '—') + '</span>';
}
function detailCell(v){
  const raw=String(v==null?'':v);
  if(!raw){return '—';}
  const status=actionStatus(raw);
  const details=actionDetails(raw);
  if(!status || status===raw){return esc(details || raw || '—');}
  if(!details || details===status){return pill(status);}
  return pill(status) + '<div style="margin-top:.3rem">' + esc(details) + '</div>';
}
function enableResizableColumns(){
  document.querySelectorAll('.table-scroll table').forEach(function(table){
    const headers=table.querySelectorAll('thead th');
    if(!headers.length){return;}
    table.style.width='max-content';
    table.style.minWidth='100%';
    headers.forEach(function(header){
      if(header.querySelector('.col-resizer')){return;}
      const resizer=document.createElement('span');
      resizer.className='col-resizer';
      header.appendChild(resizer);
      let startX=0;
      let startWidth=0;
      const onMove=function(event){
        const nextWidth=Math.max(90,startWidth + (event.clientX - startX));
        header.style.width=nextWidth + 'px';
      };
      const onUp=function(){
        resizer.classList.remove('active');
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
      };
      resizer.addEventListener('mousedown', function(event){
        event.preventDefault();
        startX=event.clientX;
        startWidth=header.getBoundingClientRect().width;
        resizer.classList.add('active');
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
    });
  });
}
document.getElementById('server-strip').innerHTML=[
  ['Tenant',DATA.tenant],
  ['Targets',String(DATA.totalTargets)],
  ['Selected Actions',DATA.selectedActions.join(', ') || 'Default'],
  ['High Impact',DATA.highImpactActions.filter(action=>DATA.selectedActions.includes(action)).join(', ') || 'None'],
  ['Generated',DATA.reportDate]
].map(function(pair){return '<div class="strip-item"><span class="strip-label">'+pair[0]+'</span><span class="strip-value">'+esc(pair[1])+'</span></div>';}).join('');
document.getElementById('hero').innerHTML=[
  kpi(DATA.contained,'Contained','core containment applied','ok'),
  kpi(DATA.needsManualFollowUp,'Follow-up','manual review still needed','warn'),
  kpi(DATA.mailboxNotFound,'Mailbox Missing','targets without a mailbox','crit'),
  kpi(DATA.passwordSuppressed,'Passwords Hidden','reset values suppressed','neutral'),
  kpi(DATA.auditFailed,'Audit Failed','activity export failures','crit'),
  kpi(DATA.auditExported,'Audit Exports','activity packages generated','neutral')
].join('');
const selectedReviewActions=DATA.selectedActions.filter(action=>DATA.reviewActions.includes(action));
const selectedRemediationActions=DATA.selectedActions.filter(action=>!DATA.reviewActions.includes(action));
const selectedHighImpactActions=DATA.selectedActions.filter(action=>DATA.highImpactActions.includes(action));
let actionCard='<div class="card"><div class="card-header"><span class="card-title">Workflow Summary</span><span class="card-badge">'+DATA.selectedActions.length+' action(s)</span></div><div class="card-body"><div class="errors"><div class="error-item" style="color:var(--text)"><strong>Review actions:</strong> '+esc(selectedReviewActions.join(', ') || 'None')+'</div><div class="error-item" style="color:var(--text)"><strong>Remediation actions:</strong> '+esc(selectedRemediationActions.join(', ') || 'None')+'</div><div class="error-item" style="color:var(--text)"><strong>High-impact actions:</strong> '+esc(selectedHighImpactActions.join(', ') || 'None')+'</div></div></div></div>';
let resultTable='<div class="card"><div class="card-header"><span class="card-title">Per-User Results</span><span class="card-badge">'+DATA.rows.length+' row(s)</span></div><div class="card-body"><div class="table-scroll"><table><thead><tr><th>UPN</th><th>User</th><th>Mailbox</th><th>Guest</th><th>Synced</th><th>Disable</th><th>Revoke</th><th>Reset</th><th>Review MFA</th><th>Remove MFA</th><th>Review Rules</th><th>Disable Rules</th><th>Forwarding</th><th>Delegates</th><th>Recent Sign-Ins</th><th>Protocols</th><th>Audit</th><th>Package</th></tr></thead><tbody>';
for(const row of DATA.rows){resultTable+='<tr><td>'+esc(row.UserPrincipalName || row.UPN || '')+'</td><td>'+pill(String(Boolean(row.UserFound)))+'</td><td>'+pill(String(Boolean(row.MailboxFound)))+'</td><td>'+pill(String(Boolean(row.IsGuest)))+'</td><td>'+pill(String(Boolean(row.IsSynced)))+'</td><td>'+pill(row.DisableUser)+'</td><td>'+pill(row.RevokeSessions)+'</td><td>'+pill(row.ResetPassword)+'</td><td>'+detailCell(row.ReviewMfaMethods)+'</td><td>'+detailCell(row.RemoveMfaMethods)+'</td><td>'+detailCell(row.ReviewInboxRules)+'</td><td>'+detailCell(row.DisableInboxRules)+'</td><td>'+detailCell(row.ReviewMailboxForwarding || row.RemoveMailboxForwarding)+'</td><td>'+detailCell(row.ReviewMailboxDelegates || row.RemoveMailboxDelegates)+'</td><td>'+detailCell(row.ReviewRecentSignIns)+'</td><td>'+detailCell(row.DisableMailboxProtocols)+'</td><td>'+detailCell(row.ExportAuditLog)+'</td><td>'+esc(row.IncidentPackagePath || '—')+'</td></tr>';}
resultTable+='</tbody></table></div></div></div>';
let errorsCard='';
if(DATA.errors.length){errorsCard='<div class="card"><div class="card-header"><span class="card-title">Errors</span><span class="card-badge">'+DATA.errors.length+'</span></div><div class="card-body"><div class="errors">'+DATA.errors.map(function(error){return '<div class="error-item">'+esc(error)+'</div>';}).join('')+'</div></div></div>';}
document.getElementById('sections').innerHTML=actionCard+resultTable+errorsCard;
enableResizableColumns();
</script>
</body>
</html>
"@

    $html | Out-File -LiteralPath $Path -Encoding utf8
    Write-Host "[+] HTML dashboard exported to: $Path" -ForegroundColor Green
}

function Save-Outputs {
    param(
        [object[]]$Results,
        [string]$DestinationPath
    )

    $statusFile = Join-Path $DestinationPath ("CompromisedAccountRemediation_Status_{0}.csv" -f $script:Timestamp)
    $errorFile = Join-Path $DestinationPath ("CompromisedAccountRemediation_Errors_{0}.log" -f $script:Timestamp)
    $htmlFile = Add-TimestampToPath -Path $(if ($ExportHtml) { $ExportHtml } else { $DestinationPath }) -BaseName 'CompromisedAccountRemediation' -Extension '.html'

    $Results | Export-Csv -Path $statusFile -NoTypeInformation
    if ($script:Errors.Count -gt 0) {
        $script:Errors | Out-File -FilePath $errorFile -Encoding utf8
    }
    Export-HtmlReport -Results $Results -Path $htmlFile

    Write-Section 'OUTPUT'
    Write-Host "Status file    : $statusFile" -ForegroundColor Yellow
    if ($script:Errors.Count -gt 0) {
        Write-Host "Error log      : $errorFile" -ForegroundColor Yellow
    }
    Write-Host "HTML dashboard : $htmlFile" -ForegroundColor Yellow

    return [pscustomobject]@{
        StatusFile = $statusFile
        ErrorFile  = if ($script:Errors.Count -gt 0) { $errorFile } else { $null }
        HtmlFile   = $htmlFile
    }
}

function Invoke-CompromisedAccountRemediation {
    $resolvedOutputPath = Initialize-OutputPath -Path $OutputPath
    Ensure-Modules -AllowInstall:$InstallMissingModules
    Connect-Services

    try {
        $targets = Resolve-TargetUsers
        $results = [System.Collections.Generic.List[psobject]]::new()

        Write-Section 'Remediation'
        Write-Host ('Target users: ' + ($targets -join ', ')) -ForegroundColor Cyan
        Write-Host ('Actions     : ' + ($Actions -join ', ')) -ForegroundColor Cyan

        foreach ($upn in $targets) {
            Write-Progress -Activity 'Processing compromised account' -Status $upn
            $context = Get-UserContext -Upn $upn
            $row = New-ResultRow -Upn $upn

            if (-not $context.Exists) {
                $row.UserFound = $false
                $results.Add([pscustomobject]$row)
                Add-ErrorEntry -User $upn -Action 'Lookup' -Message 'User not found.'
                continue
            }

            $row.UserFound = $true
            $row.MailboxFound = $context.HasMailbox
            $row.IsGuest = $context.IsGuest
            $row.IsSynced = $context.IsSynced

            foreach ($action in $Actions) {
                switch ($action) {
                    'DisableUser' {
                        $row.DisableUser = Invoke-DisableUser -Context $context -WhatIf:$WhatIfPreference
                    }
                    'RevokeSessions' {
                        $row.RevokeSessions = Invoke-RevokeSessions -Context $context -WhatIf:$WhatIfPreference
                    }
                    'ResetPassword' {
                        $passwordResult = Invoke-ResetPassword -Context $context -IncludePassword:$IncludeGeneratedPasswordsInResults -WhatIf:$WhatIfPreference
                        $row.ResetPassword = $passwordResult.Status
                        $row.GeneratedPassword = if ($IncludeGeneratedPasswordsInResults) { $passwordResult.Password } else { '[suppressed]' }
                    }
                    'ReviewMfaMethods' {
                        $row.ReviewMfaMethods = Get-MfaReview -Context $context
                    }
                    'RemoveMfaMethods' {
                        $row.RemoveMfaMethods = Invoke-RemoveMfaMethods -Context $context -WhatIf:$WhatIfPreference
                    }
                    'ReviewInboxRules' {
                        $row.ReviewInboxRules = Get-InboxRuleReview -Context $context
                    }
                    'DisableInboxRules' {
                        $row.DisableInboxRules = Invoke-DisableInboxRules -Context $context -WhatIf:$WhatIfPreference
                    }
                    'ReviewMailboxForwarding' {
                        $row.ReviewMailboxForwarding = Get-MailboxForwardingReview -Context $context
                    }
                    'RemoveMailboxForwarding' {
                        $row.RemoveMailboxForwarding = Invoke-RemoveMailboxForwarding -Context $context -WhatIf:$WhatIfPreference
                    }
                    'DisableSignature' {
                        $row.DisableSignature = Invoke-DisableSignature -Context $context -WhatIf:$WhatIfPreference
                    }
                    'ReviewMailboxDelegates' {
                        $row.ReviewMailboxDelegates = Get-MailboxDelegateReview -Context $context
                    }
                    'RemoveMailboxDelegates' {
                        $row.RemoveMailboxDelegates = Invoke-RemoveMailboxDelegates -Context $context -WhatIf:$WhatIfPreference
                    }
                    'ReviewRecentSignIns' {
                        $row.ReviewRecentSignIns = Get-RecentSignInReview -Context $context
                    }
                    'DisableMailboxProtocols' {
                        $row.DisableMailboxProtocols = Invoke-DisableMailboxProtocols -Context $context -WhatIf:$WhatIfPreference
                    }
                    'ExportAuditLog' {
                        $auditDestination = if ($ExportIncidentPackage) {
                            Get-IncidentPackagePath -DestinationPath $resolvedOutputPath -Upn $context.Upn
                        }
                        else {
                            $resolvedOutputPath
                        }
                        $row.ExportAuditLog = Export-ActivityLog -Context $context -DestinationPath $auditDestination -Days $AuditLogDays
                    }
                }
            }

            if ($ExportIncidentPackage) {
                $row.IncidentPackagePath = Export-IncidentPackageSummary -Context $context -Row $row -DestinationPath $resolvedOutputPath
            }

            $results.Add([pscustomobject]$row)
        }

        Write-Progress -Activity 'Processing compromised account' -Completed
        $outputInfo = Save-Outputs -Results @($results) -DestinationPath $resolvedOutputPath

        Write-Section 'Summary'
        @($results) | Format-Table -AutoSize | Out-Host

        return [pscustomobject]@{
            Results    = @($results)
            OutputInfo = $outputInfo
        }
    }
    finally {
        Disconnect-Services
    }
}

Invoke-CompromisedAccountRemediation
