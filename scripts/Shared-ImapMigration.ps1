[CmdletBinding()]
param()

$script:ImapMigrationRequiredHeaders = @("EmailAddress", "UserName", "Password")
$script:ImapMigrationOptionalHeaders = @("UserRoot")
$script:ImapSensitiveValues = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

function Import-ExchangeOnlineForImapMigration {
    Write-Host ""
    Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan
    Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
    Write-Host "  [+] ExchangeOnlineManagement" -ForegroundColor Green
}

function Connect-ImapMigrationExchangeOnline {
    param(
        [string]$TenantId,
        [string]$TenantAdminUpn
    )

    Import-ExchangeOnlineForImapMigration

    Write-SectionHeader "CONNECTING TO EXCHANGE ONLINE"
    if ($TenantId) {
        Write-Host "[*] Requested tenant: $TenantId"
    }
    else {
        Write-Host "[*] Requested tenant: auto-detect"
    }

    if ($TenantAdminUpn) {
        Write-Host "[*] Sign-in hint: $TenantAdminUpn"
    }

    Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
    Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    $exoParams = @{
        ShowBanner   = $false
        ShowProgress = $false
        Device       = $true
        ErrorAction  = "Stop"
    }

    if ($TenantId) {
        $exoParams.Organization = $TenantId
    }

    if ($TenantAdminUpn) {
        $exoParams.UserPrincipalName = $TenantAdminUpn
    }

    Connect-ExchangeOnline @exoParams | Out-Null
    $script:ToolboxTenantLabel = if ($TenantId) { $TenantId } else { "ExchangeOnline" }
    Write-Host "[+] Connected to Exchange Online" -ForegroundColor Green
}

function Disconnect-ImapMigrationExchangeOnline {
    $disconnectCommand = Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue
    if ($disconnectCommand) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}

function Resolve-ImapMigrationCsvPath {
    param([string]$CsvPath)

    $candidate = $CsvPath
    if (-not $candidate) {
        $candidate = $env:M365_TOOLBOX_IMAP_MIGRATION_CSV_PATH
    }

    if (-not $candidate) {
        throw "CSV path is required. Provide CsvPath or set M365_TOOLBOX_IMAP_MIGRATION_CSV_PATH in the backend container."
    }

    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        throw "CSV file was not found at the provided container path. Mount the migration CSV into the backend container and provide that path."
    }

    return (Resolve-Path -LiteralPath $candidate).Path
}

function Register-ImapSensitiveValues {
    param([array]$Rows)

    foreach ($row in @($Rows)) {
        foreach ($propertyName in @("UserName", "Password")) {
            $property = $row.PSObject.Properties |
                Where-Object { $_.Name -ieq $propertyName } |
                Select-Object -First 1
            if ($property -and $property.Value) {
                [void]$script:ImapSensitiveValues.Add([string]$property.Value)
            }
        }
    }
}

function Protect-ImapMigrationText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    $safeText = [string]$Text
    foreach ($secret in @($script:ImapSensitiveValues)) {
        if ($secret -and $secret.Length -ge 3) {
            $safeText = $safeText.Replace($secret, "[redacted]")
        }
    }

    return $safeText
}

function Import-ImapMigrationCsv {
    param([string]$CsvPath)

    Write-ProgressStep "Validating IMAP migration CSV"
    $resolvedPath = Resolve-ImapMigrationCsvPath -CsvPath $CsvPath
    $rows = @(Import-Csv -LiteralPath $resolvedPath)

    if ($rows.Count -eq 0) {
        throw "CSV file contains no mailbox rows."
    }

    $headers = @($rows[0].PSObject.Properties.Name)
    foreach ($requiredHeader in $script:ImapMigrationRequiredHeaders) {
        if ($requiredHeader -notin $headers) {
            throw "CSV schema is missing required header '$requiredHeader'. Expected headers: $($script:ImapMigrationRequiredHeaders -join ', '). Optional: $($script:ImapMigrationOptionalHeaders -join ', ')."
        }
    }

    $redactedRows = [System.Collections.Generic.List[object]]::new()
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        foreach ($requiredHeader in $script:ImapMigrationRequiredHeaders) {
            if (-not ([string]$row.$requiredHeader).Trim()) {
                throw "CSV row $rowNumber has an empty required value for '$requiredHeader'."
            }
        }

        if ([string]$row.EmailAddress -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
            throw "CSV row $rowNumber has an invalid EmailAddress value."
        }

        [void]$redactedRows.Add([pscustomobject]@{
            RowNumber       = $rowNumber
            EmailAddress    = [string]$row.EmailAddress
            UserNamePresent = if ([string]$row.UserName) { "Yes" } else { "No" }
            PasswordPresent = if ([string]$row.Password) { "Yes" } else { "No" }
            UserRootPresent = if ($row.PSObject.Properties.Name -contains "UserRoot" -and [string]$row.UserRoot) { "Yes" } else { "No" }
        })
    }

    Register-ImapSensitiveValues -Rows $rows

    return [pscustomobject]@{
        Path            = $resolvedPath
        Rows            = $rows
        RedactedRows    = @($redactedRows)
        MailboxCount    = $rows.Count
        RequiredHeaders = $script:ImapMigrationRequiredHeaders
        OptionalHeaders = $script:ImapMigrationOptionalHeaders
    }
}

function New-ImapMigrationArtifactSet {
    param(
        [string]$ExportHtml,
        [string]$OutputPath,
        [string]$BaseName
    )

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName $BaseName -OutputPath $OutputPath
    $directory = [System.IO.Path]::GetDirectoryName($htmlPath)
    if (-not $directory) {
        $directory = (Get-Location).Path
        $htmlPath = Join-Path $directory $htmlPath
    }
    if ($directory) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $basePath = Join-Path $directory ([System.IO.Path]::GetFileNameWithoutExtension($htmlPath))

    return [pscustomobject]@{
        HtmlPath        = $htmlPath
        BasePath        = $basePath
        SummaryCsvPath  = "$basePath-summary.csv"
        UsersCsvPath    = "$basePath-users.csv"
        DiagnosticsPath = "$basePath-diagnostics.json"
        NotesPath       = "$basePath-notes.txt"
    }
}

function Export-ImapMigrationDiagnostics {
    param(
        [Parameter(Mandatory)]
        $Artifacts,

        [Parameter(Mandatory)]
        [object]$Data,

        [string[]]$Notes = @()
    )

    $Data | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Artifacts.DiagnosticsPath -Encoding UTF8
    if ($Notes.Count -gt 0) {
        $Notes | Set-Content -LiteralPath $Artifacts.NotesPath -Encoding UTF8
    }

    Write-Host "[+] Diagnostics exported to: $($Artifacts.DiagnosticsPath)" -ForegroundColor Green
    if ($Notes.Count -gt 0) {
        Write-Host "[+] Notes exported to: $($Artifacts.NotesPath)" -ForegroundColor Green
    }
}

function Get-ImapMigrationEndpointSnapshot {
    param($Endpoint)

    if (-not $Endpoint) {
        return $null
    }

    return [pscustomobject]@{
        Identity                = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "Identity")
        Name                    = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "Name")
        EndpointType            = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "EndpointType")
        RemoteServer            = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "RemoteServer")
        Port                    = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "Port")
        Security                = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "Security")
        MaxConcurrentMigrations = [string](Get-DirectoryObjectValue -DirectoryObject $Endpoint -Name "MaxConcurrentMigrations")
    }
}

function Test-ImapEndpointAvailabilitySafe {
    param(
        [Parameter(Mandatory)]
        [string]$ImapServer,

        [Parameter(Mandatory)]
        [int]$Port,

        [ValidateSet("Ssl", "Tls", "None")]
        [string]$Security = "Ssl",

        [switch]$SkipCertificateValidation
    )

    Write-ProgressStep "Testing IMAP endpoint availability from Exchange Online"
    $testParams = @{
        Imap        = $true
        RemoteServer = $ImapServer
        Port        = $Port
        Security    = $Security
        ErrorAction = "Stop"
    }

    if ($SkipCertificateValidation) {
        $testParams.AcceptUntrustedCertificates = $true
    }

    return Test-MigrationServerAvailability @testParams
}

function Get-ImapMigrationBatchSnapshot {
    param($Batch)

    if (-not $Batch) {
        return $null
    }

    return [pscustomobject]@{
        Identity        = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "Identity")
        Name            = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "Name")
        Status          = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "Status")
        State           = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "State")
        TotalCount      = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "TotalCount")
        SyncedCount     = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "SyncedCount")
        FailedCount     = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "FailedCount")
        SourceEndpoint  = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "SourceEndpoint")
        CreationDate    = [string](Get-DirectoryObjectValue -DirectoryObject $Batch -Name "CreationDate")
    }
}

function Get-ImapMigrationUserRows {
    param([Parameter(Mandatory)][string]$BatchName)

    $rows = [System.Collections.Generic.List[object]]::new()
    $users = @(Get-MigrationUser -BatchId $BatchName -ErrorAction SilentlyContinue)

    foreach ($user in $users) {
        $stats = $null
        try {
            $stats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction Stop
        }
        catch {}

        [void]$rows.Add([pscustomobject]@{
            Identity           = Protect-ImapMigrationText -Text ([string](Get-DirectoryObjectValue -DirectoryObject $user -Name "Identity"))
            EmailAddress       = Protect-ImapMigrationText -Text ([string](Get-DirectoryObjectValue -DirectoryObject $user -Name "EmailAddress"))
            Recipient          = Protect-ImapMigrationText -Text ([string](Get-DirectoryObjectValue -DirectoryObject $user -Name "Recipient"))
            Status             = [string](Get-DirectoryObjectValue -DirectoryObject $user -Name "Status")
            SyncedItemCount    = if ($stats) { [string](Get-DirectoryObjectValue -DirectoryObject $stats -Name "SyncedItemCount") } else { "" }
            SkippedItemCount   = if ($stats) { [string](Get-DirectoryObjectValue -DirectoryObject $stats -Name "SkippedItemCount") } else { "" }
            BytesTransferred   = if ($stats) { [string](Get-DirectoryObjectValue -DirectoryObject $stats -Name "BytesTransferred") } else { "" }
            PercentageComplete = if ($stats) { [string](Get-DirectoryObjectValue -DirectoryObject $stats -Name "PercentageComplete") } else { "" }
            ErrorSummary       = if ($stats) { Protect-ImapMigrationText -Text ([string](Get-DirectoryObjectValue -DirectoryObject $stats -Name "ErrorSummary")) } else { "" }
        })
    }

    return @($rows)
}
