[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$TenantAdminUpn,

    [Parameter(Mandatory)]
    [string]$EndpointName,

    [Parameter(Mandatory)]
    [string]$ImapServer,

    [ValidateRange(1, 65535)]
    [int]$Port = 993,

    [ValidateSet("Ssl", "Tls", "None")]
    [string]$Security = "Ssl",

    [string]$CsvPath,
    [switch]$SkipCertificateValidation,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")
. (Join-Path $PSScriptRoot "Shared-ImapMigration.ps1")

$checks = [System.Collections.Generic.List[object]]::new()
$errors = [System.Collections.Generic.List[string]]::new()
$csvInfo = $null
$availability = $null
$exchangeConnected = $false

try {
    Write-SectionHeader "IMAP MIGRATION PREFLIGHT"
    Write-Host "[*] Workflow: Exchange Online IMAP migration orchestration"
    Write-Host "[*] This is not a local mailbox copy. Microsoft 365 migration services perform the mailbox sync."
    Write-Host "[*] Mailbox passwords must live only in the mounted CSV file, never in browser fields."
    if ($SkipCertificateValidation) {
        Write-Host "[!] Skip certificate validation is enabled. Use only for temporary testing." -ForegroundColor Yellow
    }

    try {
        $csvInfo = Import-ImapMigrationCsv -CsvPath $CsvPath
        [void]$checks.Add([pscustomobject]@{
            Check  = "CSV file"
            Result = "Passed"
            Detail = "$($csvInfo.MailboxCount) mailbox row(s); required headers present"
        })
    }
    catch {
        $message = Get-ToolboxExceptionMessage -Exception $_.Exception
        [void]$errors.Add($message)
        [void]$checks.Add([pscustomobject]@{
            Check  = "CSV file"
            Result = "Failed"
            Detail = Protect-ImapMigrationText -Text $message
        })
    }

    if ($errors.Count -eq 0) {
        try {
            Connect-ImapMigrationExchangeOnline -TenantId $TenantId -TenantAdminUpn $TenantAdminUpn
            $exchangeConnected = $true
            [void]$checks.Add([pscustomobject]@{
                Check  = "Exchange Online"
                Result = "Passed"
                Detail = "Device-code authentication completed"
            })
        }
        catch {
            $message = Get-ToolboxExceptionMessage -Exception $_.Exception
            [void]$errors.Add($message)
            [void]$checks.Add([pscustomobject]@{
                Check  = "Exchange Online"
                Result = "Failed"
                Detail = Protect-ImapMigrationText -Text $message
            })
        }
    }
    else {
        [void]$checks.Add([pscustomobject]@{
            Check  = "Exchange Online"
            Result = "Skipped"
            Detail = "CSV validation did not pass"
        })
    }

    if ($exchangeConnected) {
        try {
            $availability = Test-ImapEndpointAvailabilitySafe -ImapServer $ImapServer -Port $Port -Security $Security -SkipCertificateValidation:$SkipCertificateValidation
            [void]$checks.Add([pscustomobject]@{
                Check  = "IMAP endpoint"
                Result = "Passed"
                Detail = "Exchange Online can reach ${ImapServer}:$Port using $Security"
            })
        }
        catch {
            $message = Get-ToolboxExceptionMessage -Exception $_.Exception
            [void]$errors.Add($message)
            [void]$checks.Add([pscustomobject]@{
                Check  = "IMAP endpoint"
                Result = "Failed"
                Detail = Protect-ImapMigrationText -Text $message
            })
        }
    }
    else {
        [void]$checks.Add([pscustomobject]@{
            Check  = "IMAP endpoint"
            Result = "Skipped"
            Detail = "Exchange Online connection did not complete"
        })
    }
}
finally {
    Disconnect-ImapMigrationExchangeOnline
}

$artifacts = New-ImapMigrationArtifactSet -ExportHtml $ExportHtml -OutputPath $OutputPath -BaseName "ImapMigrationPreflight"
$readiness = if ($errors.Count -eq 0) { "Ready" } else { "Blocked" }
$mailboxCount = if ($csvInfo) { $csvInfo.MailboxCount } else { 0 }
$mailboxRows = @()
if ($csvInfo) {
    $mailboxRows = @($csvInfo.RedactedRows)
}

$checks | Export-Csv -LiteralPath $artifacts.SummaryCsvPath -NoTypeInformation -Encoding UTF8
$mailboxRows | Export-Csv -LiteralPath $artifacts.UsersCsvPath -NoTypeInformation -Encoding UTF8

$notes = @(
    "Expected CSV headers: EmailAddress, UserName, Password. Optional: UserRoot.",
    "Mailbox passwords are read from the mounted CSV file only and are not emitted into reports.",
    "Microsoft 365 target mailboxes must already exist and be licensed before creating the migration batch.",
    "IMAP migration moves mail only. Contacts, calendar items, tasks, and rules are not migrated by this workflow."
)

Export-ImapMigrationDiagnostics -Artifacts $artifacts -Data ([pscustomobject]@{
    Workflow                  = "Preflight"
    EndpointName              = $EndpointName
    ImapServer                = $ImapServer
    Port                      = $Port
    Security                  = $Security
    SkipCertificateValidation = [bool]$SkipCertificateValidation
    MailboxCount              = $mailboxCount
    Readiness                 = $readiness
    RequiredCsvHeaders        = $script:ImapMigrationRequiredHeaders
    OptionalCsvHeaders        = $script:ImapMigrationOptionalHeaders
    Checks                    = @($checks)
    Availability              = if ($availability) { Protect-ImapMigrationText -Text ($availability | ConvertTo-Json -Depth 6 -Compress) } else { "" }
}) -Notes $notes

Export-ToolboxHtmlReport -Path $artifacts.HtmlPath -Title "M365 IMAP Migration Preflight" -Tenant $script:ToolboxTenantLabel -Subtitle "Readiness checks for Exchange Online IMAP migration" -Kpis @(
    @{ label = "Readiness"; value = $readiness; sub = "Preflight result"; cls = if ($readiness -eq "Ready") { "ok" } else { "crit" } },
    @{ label = "Mailboxes"; value = $mailboxCount; sub = "CSV rows"; cls = "neutral" },
    @{ label = "IMAP Server"; value = "${ImapServer}:$Port"; sub = $Security; cls = "neutral" },
    @{ label = "Secrets"; value = "Protected"; sub = "Passwords not reported"; cls = "ok" }
) -StripItems @(
    @{ label = "Endpoint"; value = $EndpointName },
    @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        id = "readiness-checks"
        title = "Readiness Checks"
        badge = "$($checks.Count) check(s)"
        columns = @(
            @{ key = "Check"; header = "Check" },
            @{ key = "Result"; header = "Result"; type = "pill" },
            @{ key = "Detail"; header = "Detail" }
        )
        rows = @($checks)
    },
    @{
        id = "mailbox-rows"
        title = "Mailbox Rows"
        badge = "$($mailboxRows.Count) mailbox(es)"
        filterKey = "EmailAddress"
        filterLabel = "Mailbox Filter"
        filterPlaceholder = "Type a mailbox address"
        columns = @(
            @{ key = "RowNumber"; header = "CSV Row" },
            @{ key = "EmailAddress"; header = "Target Mailbox" },
            @{ key = "UserNamePresent"; header = "User Name Present"; type = "pill" },
            @{ key = "PasswordPresent"; header = "Password Present"; type = "pill" },
            @{ key = "UserRootPresent"; header = "UserRoot Present"; type = "pill" }
        )
        rows = @($mailboxRows)
    },
    @{
        id = "operator-notes"
        title = "Operator Notes"
        badge = "Safe usage"
        columns = @(
            @{ key = "Note"; header = "Note" }
        )
        rows = @($notes | ForEach-Object { [pscustomobject]@{ Note = $_ } })
    }
)

Write-Host "[+] HTML dashboard exported to: $($artifacts.HtmlPath)" -ForegroundColor Green
Write-Host "[+] Summary CSV exported to: $($artifacts.SummaryCsvPath)" -ForegroundColor Green
Write-Host "[+] Mailbox CSV exported to: $($artifacts.UsersCsvPath)" -ForegroundColor Green

if ($errors.Count -gt 0) {
    throw "IMAP migration preflight failed. Review the readiness report and correct the failed check(s)."
}
