[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$TenantAdminUpn,

    [Parameter(Mandatory)]
    [string]$EndpointName,

    [Parameter(Mandatory)]
    [string]$BatchName,

    [Parameter(Mandatory)]
    [string]$ImapServer,

    [ValidateRange(1, 65535)]
    [int]$Port = 993,

    [ValidateSet("Ssl", "Tls", "None")]
    [string]$Security = "Ssl",

    [string]$CsvPath,
    [string]$TargetDeliveryDomain,
    [string[]]$NotificationEmailAddresses,

    [ValidateRange(1, 100)]
    [int]$MaxConcurrentMigrations = 20,

    [switch]$AutoStart,
    [switch]$AutoComplete,
    [switch]$SkipCertificateValidation,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")
. (Join-Path $PSScriptRoot "Shared-ImapMigration.ps1")

$operations = [System.Collections.Generic.List[object]]::new()
$errors = [System.Collections.Generic.List[string]]::new()
$csvInfo = $null
$endpointSnapshot = $null
$batchSnapshot = $null
$userRows = @()
$createdBatch = $false
$startedBatch = $false

try {
    Write-SectionHeader "IMAP MIGRATION BATCH"
    Write-Host "[*] Workflow: Exchange Online IMAP migration batch creation"
    Write-Host "[*] This workflow uploads CSVData to Exchange Online migration services."
    Write-Host "[*] Mailbox passwords are read from the mounted CSV file only."
    if ($SkipCertificateValidation) {
        Write-Host "[!] Skip certificate validation is enabled. Use only for temporary testing." -ForegroundColor Yellow
    }

    $csvInfo = Import-ImapMigrationCsv -CsvPath $CsvPath
    [void]$operations.Add([pscustomobject]@{
        Step   = "CSV validation"
        Result = "Passed"
        Detail = "$($csvInfo.MailboxCount) mailbox row(s); no secret values emitted"
    })

    Connect-ImapMigrationExchangeOnline -TenantId $TenantId -TenantAdminUpn $TenantAdminUpn
    [void]$operations.Add([pscustomobject]@{
        Step   = "Exchange Online connection"
        Result = "Passed"
        Detail = "Device-code authentication completed"
    })

    try {
        Test-ImapEndpointAvailabilitySafe -ImapServer $ImapServer -Port $Port -Security $Security -SkipCertificateValidation:$SkipCertificateValidation | Out-Null
        [void]$operations.Add([pscustomobject]@{
            Step   = "IMAP endpoint test"
            Result = "Passed"
            Detail = "Exchange Online can reach ${ImapServer}:$Port using $Security"
        })
    }
    catch {
        $message = Get-ToolboxExceptionMessage -Exception $_.Exception
        [void]$errors.Add($message)
        [void]$operations.Add([pscustomobject]@{
            Step   = "IMAP endpoint test"
            Result = "Failed"
            Detail = Protect-ImapMigrationText -Text $message
        })
        throw
    }

    $existingEndpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
    if ($existingEndpoint) {
        $endpointSnapshot = Get-ImapMigrationEndpointSnapshot -Endpoint $existingEndpoint
        $mismatches = [System.Collections.Generic.List[string]]::new()
        if ($endpointSnapshot.RemoteServer -and $endpointSnapshot.RemoteServer -ine $ImapServer) {
            [void]$mismatches.Add("server '$($endpointSnapshot.RemoteServer)'")
        }
        if ($endpointSnapshot.Port -and [string]$endpointSnapshot.Port -ne [string]$Port) {
            [void]$mismatches.Add("port '$($endpointSnapshot.Port)'")
        }
        if ($endpointSnapshot.Security -and $endpointSnapshot.Security -ine $Security) {
            [void]$mismatches.Add("security '$($endpointSnapshot.Security)'")
        }

        if ($mismatches.Count -gt 0) {
            throw "Migration endpoint '$EndpointName' already exists but does not match requested settings: $($mismatches -join ', '). Use a different endpoint name or match the existing endpoint."
        }

        [void]$operations.Add([pscustomobject]@{
            Step   = "Migration endpoint"
            Result = "Reused"
            Detail = "Existing endpoint '$EndpointName' matched requested server, port, and security settings"
        })
    }
    else {
        $endpointParams = @{
            IMAP                    = $true
            Name                    = $EndpointName
            RemoteServer            = $ImapServer
            Port                    = $Port
            Security                = $Security
            MaxConcurrentMigrations = $MaxConcurrentMigrations
            ErrorAction             = "Stop"
        }

        if ($SkipCertificateValidation) {
            $endpointParams.AcceptUntrustedCertificates = $true
        }

        $createdEndpoint = New-MigrationEndpoint @endpointParams
        $endpointSnapshot = Get-ImapMigrationEndpointSnapshot -Endpoint $createdEndpoint
        [void]$operations.Add([pscustomobject]@{
            Step   = "Migration endpoint"
            Result = "Created"
            Detail = "Created endpoint '$EndpointName'"
        })
    }

    $existingBatch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
    if ($existingBatch) {
        throw "Migration batch '$BatchName' already exists. Use the status workflow to monitor it, or choose a new batch name."
    }

    Write-ProgressStep "Creating migration batch"
    $csvBytes = [System.IO.File]::ReadAllBytes($csvInfo.Path)
    $batchParams = @{
        Name           = $BatchName
        CSVData        = $csvBytes
        SourceEndpoint = $EndpointName
        ErrorAction    = "Stop"
    }

    if ($TargetDeliveryDomain) {
        $batchParams.TargetDeliveryDomain = $TargetDeliveryDomain
    }

    if ($AutoComplete) {
        $batchParams.AutoComplete = $true
    }

    $notificationEmails = @(Normalize-DelimitedValue -Value $NotificationEmailAddresses)
    if ($notificationEmails.Count -gt 0) {
        $batchParams.NotificationEmails = $notificationEmails
    }

    $newBatch = New-MigrationBatch @batchParams
    $createdBatch = $true
    $batchSnapshot = Get-ImapMigrationBatchSnapshot -Batch $newBatch
    [void]$operations.Add([pscustomobject]@{
        Step   = "Migration batch"
        Result = "Created"
        Detail = "Created batch '$BatchName' with $($csvInfo.MailboxCount) mailbox row(s); auto-complete: $([bool]$AutoComplete)"
    })

    if ($AutoStart) {
        Write-ProgressStep "Starting migration batch"
        Start-MigrationBatch -Identity $BatchName -ErrorAction Stop | Out-Null
        $startedBatch = $true
        [void]$operations.Add([pscustomobject]@{
            Step   = "Batch start"
            Result = "Started"
            Detail = "Batch '$BatchName' was started"
        })
    }
    else {
        [void]$operations.Add([pscustomobject]@{
            Step   = "Batch start"
            Result = "Not started"
            Detail = "AutoStart was not selected"
        })
    }

    $latestBatch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
    if ($latestBatch) {
        $batchSnapshot = Get-ImapMigrationBatchSnapshot -Batch $latestBatch
    }
    $userRows = @(Get-ImapMigrationUserRows -BatchName $BatchName)
}
catch {
    $message = Get-ToolboxExceptionMessage -Exception $_.Exception
    if ($message -and $message -notin $errors) {
        [void]$errors.Add($message)
    }
    [void]$operations.Add([pscustomobject]@{
        Step   = "Workflow"
        Result = "Failed"
        Detail = Protect-ImapMigrationText -Text $message
    })
}
finally {
    Disconnect-ImapMigrationExchangeOnline
}

$artifacts = New-ImapMigrationArtifactSet -ExportHtml $ExportHtml -OutputPath $OutputPath -BaseName "ImapMigrationBatch"
$mailboxCount = if ($csvInfo) { $csvInfo.MailboxCount } else { 0 }
$statusValue = if ($errors.Count -eq 0) { "Created" } else { "Failed" }

$operations | Export-Csv -LiteralPath $artifacts.SummaryCsvPath -NoTypeInformation -Encoding UTF8
$userRows | Export-Csv -LiteralPath $artifacts.UsersCsvPath -NoTypeInformation -Encoding UTF8

$notes = @(
    "CSVData was passed to New-MigrationBatch; this script does not copy mail locally.",
    "No mailbox passwords are printed, written to artifact names, or exposed as UI fields.",
    "IMAP migration moves mail only. Contacts, calendar items, tasks, and rules are not migrated.",
    "Target Microsoft 365 mailboxes must already exist and be licensed."
)

Export-ImapMigrationDiagnostics -Artifacts $artifacts -Data ([pscustomobject]@{
    Workflow                  = "Batch"
    EndpointName              = $EndpointName
    BatchName                 = $BatchName
    ImapServer                = $ImapServer
    Port                      = $Port
    Security                  = $Security
    TargetDeliveryDomain      = $TargetDeliveryDomain
    MaxConcurrentMigrations   = $MaxConcurrentMigrations
    SkipCertificateValidation = [bool]$SkipCertificateValidation
    AutoStartRequested        = [bool]$AutoStart
    AutoCompleteRequested     = [bool]$AutoComplete
    CreatedBatch              = $createdBatch
    StartedBatch              = $startedBatch
    MailboxCount              = $mailboxCount
    Endpoint                  = $endpointSnapshot
    Batch                     = $batchSnapshot
    Operations                = @($operations)
    Errors                    = @($errors | ForEach-Object { Protect-ImapMigrationText -Text $_ })
}) -Notes $notes

Export-ToolboxHtmlReport -Path $artifacts.HtmlPath -Title "M365 IMAP Migration Batch" -Tenant $script:ToolboxTenantLabel -Subtitle "Exchange Online IMAP migration batch orchestration" -Kpis @(
    @{ label = "Workflow"; value = $statusValue; sub = "Batch creation"; cls = if ($errors.Count -eq 0) { "ok" } else { "crit" } },
    @{ label = "Mailboxes"; value = $mailboxCount; sub = "CSV rows"; cls = "neutral" },
    @{ label = "Batch"; value = $BatchName; sub = if ($batchSnapshot) { $batchSnapshot.Status } else { "Unavailable" }; cls = "neutral" },
    @{ label = "Autostart"; value = if ($AutoStart) { "Yes" } else { "No" }; sub = if ($startedBatch) { "Started" } else { "Manual start" }; cls = if ($startedBatch) { "ok" } else { "neutral" } }
) -StripItems @(
    @{ label = "Endpoint"; value = $EndpointName },
    @{ label = "IMAP Server"; value = "${ImapServer}:$Port ($Security)" },
    @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        id = "batch-operations"
        title = "Batch Operations"
        badge = "$($operations.Count) step(s)"
        columns = @(
            @{ key = "Step"; header = "Step" },
            @{ key = "Result"; header = "Result"; type = "pill" },
            @{ key = "Detail"; header = "Detail" }
        )
        rows = @($operations)
    },
    @{
        id = "migration-users"
        title = "Migration Users"
        badge = "$($userRows.Count) user(s)"
        filterKey = "EmailAddress"
        filterLabel = "Mailbox Filter"
        filterPlaceholder = "Type a mailbox address"
        columns = @(
            @{ key = "EmailAddress"; header = "Mailbox" },
            @{ key = "Status"; header = "Status"; type = "pill" },
            @{ key = "SyncedItemCount"; header = "Synced Items" },
            @{ key = "SkippedItemCount"; header = "Skipped Items" },
            @{ key = "BytesTransferred"; header = "Bytes Transferred" },
            @{ key = "PercentageComplete"; header = "Complete" },
            @{ key = "ErrorSummary"; header = "Error Summary" }
        )
        rows = @($userRows)
    },
    @{
        id = "security-notes"
        title = "Security Notes"
        badge = "Secrets protected"
        columns = @(
            @{ key = "Note"; header = "Note" }
        )
        rows = @($notes | ForEach-Object { [pscustomobject]@{ Note = $_ } })
    }
)

Write-Host "[+] HTML dashboard exported to: $($artifacts.HtmlPath)" -ForegroundColor Green
Write-Host "[+] Summary CSV exported to: $($artifacts.SummaryCsvPath)" -ForegroundColor Green
Write-Host "[+] User status CSV exported to: $($artifacts.UsersCsvPath)" -ForegroundColor Green

if ($errors.Count -gt 0) {
    throw "IMAP migration batch workflow failed. Review the generated report and diagnostics artifact."
}
