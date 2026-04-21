[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$TenantAdminUpn,

    [Parameter(Mandatory)]
    [string]$BatchName,

    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")
. (Join-Path $PSScriptRoot "Shared-ImapMigration.ps1")

$summaryRows = [System.Collections.Generic.List[object]]::new()
$errors = [System.Collections.Generic.List[string]]::new()
$batchSnapshot = $null
$userRows = @()

try {
    Write-SectionHeader "IMAP MIGRATION STATUS"
    Write-Host "[*] Querying Exchange Online migration batch '$BatchName'."
    Connect-ImapMigrationExchangeOnline -TenantId $TenantId -TenantAdminUpn $TenantAdminUpn

    $batch = Get-MigrationBatch -Identity $BatchName -ErrorAction Stop
    $batchSnapshot = Get-ImapMigrationBatchSnapshot -Batch $batch
    $userRows = @(Get-ImapMigrationUserRows -BatchName $BatchName)

    foreach ($propertyName in @("Name", "Status", "State", "TotalCount", "SyncedCount", "FailedCount", "SourceEndpoint", "CreationDate")) {
        [void]$summaryRows.Add([pscustomobject]@{
            Property = $propertyName
            Value    = [string](Get-DirectoryObjectValue -DirectoryObject $batchSnapshot -Name $propertyName)
        })
    }
}
catch {
    $message = Get-ToolboxExceptionMessage -Exception $_.Exception
    [void]$errors.Add($message)
    [void]$summaryRows.Add([pscustomobject]@{
        Property = "Workflow"
        Value    = Protect-ImapMigrationText -Text $message
    })
}
finally {
    Disconnect-ImapMigrationExchangeOnline
}

$artifacts = New-ImapMigrationArtifactSet -ExportHtml $ExportHtml -OutputPath $OutputPath -BaseName "ImapMigrationStatus"
$statusValue = if ($errors.Count -eq 0) { "Collected" } else { "Failed" }

$summaryRows | Export-Csv -LiteralPath $artifacts.SummaryCsvPath -NoTypeInformation -Encoding UTF8
$userRows | Export-Csv -LiteralPath $artifacts.UsersCsvPath -NoTypeInformation -Encoding UTF8

$notes = @(
    "This workflow only reads Exchange Online migration status.",
    "Per-user error summaries are redacted before export.",
    "Use the Exchange admin center or Exchange Online PowerShell for operational actions outside this toolbox workflow."
)

Export-ImapMigrationDiagnostics -Artifacts $artifacts -Data ([pscustomobject]@{
    Workflow   = "Status"
    BatchName  = $BatchName
    Status     = $statusValue
    Batch      = $batchSnapshot
    Users      = @($userRows)
    Errors     = @($errors | ForEach-Object { Protect-ImapMigrationText -Text $_ })
}) -Notes $notes

Export-ToolboxHtmlReport -Path $artifacts.HtmlPath -Title "M365 IMAP Migration Status" -Tenant $script:ToolboxTenantLabel -Subtitle "Current Exchange Online migration batch status" -Kpis @(
    @{ label = "Workflow"; value = $statusValue; sub = "Status collection"; cls = if ($errors.Count -eq 0) { "ok" } else { "crit" } },
    @{ label = "Batch"; value = $BatchName; sub = if ($batchSnapshot) { $batchSnapshot.Status } else { "Unavailable" }; cls = "neutral" },
    @{ label = "Users"; value = $userRows.Count; sub = "Migration users"; cls = "neutral" },
    @{ label = "Failed"; value = if ($batchSnapshot) { $batchSnapshot.FailedCount } else { "n/a" }; sub = "Batch failures"; cls = if ($batchSnapshot -and [int]($batchSnapshot.FailedCount -as [int]) -gt 0) { "warn" } else { "neutral" } }
) -StripItems @(
    @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        id = "batch-summary"
        title = "Batch Summary"
        badge = "$($summaryRows.Count) field(s)"
        columns = @(
            @{ key = "Property"; header = "Property" },
            @{ key = "Value"; header = "Value" }
        )
        rows = @($summaryRows)
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
        id = "operator-notes"
        title = "Operator Notes"
        badge = "Read-only"
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
    throw "IMAP migration status workflow failed. Review the generated report and diagnostics artifact."
}
