[CmdletBinding()]
param(
    [string]$TenantId,
    [int]$LookbackDays = 14,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("AuditLog.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING MAILBOX SIGN-IN DATA"

$startUtc = (Get-Date).AddDays(-1 * $LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$signIns = @(Get-MgAuditLogSignIn -Filter "createdDateTime ge $startUtc" -All -ErrorAction Stop)
$mailboxRows = [System.Collections.Generic.List[object]]::new()

foreach ($signIn in $signIns) {
    $appName = [string]$signIn.AppDisplayName
    $resourceName = [string]$signIn.ResourceDisplayName
    if ($appName -notmatch 'Exchange|Outlook|Office 365 Exchange Online' -and $resourceName -notmatch 'Exchange|Outlook|Office 365 Exchange Online') {
        continue
    }

    $signals = [System.Collections.Generic.List[string]]::new()
    $clientApp = [string]$signIn.ClientAppUsed
    if ($clientApp -match 'IMAP|POP|SMTP|Other clients') { [void]$signals.Add("Legacy or uncommon client: $clientApp") }
    if ($signIn.Status -and $signIn.Status.ErrorCode -and [int]$signIn.Status.ErrorCode -ne 0) { [void]$signals.Add("Failure code $($signIn.Status.ErrorCode)") }
    if ($signIn.ConditionalAccessStatus -and [string]$signIn.ConditionalAccessStatus -eq 'failure') { [void]$signals.Add("Conditional Access failure") }
    if ($signIn.IsInteractive -eq $false) { [void]$signals.Add("Non-interactive mailbox access") }
    $country = if ($signIn.Location) { [string]$signIn.Location.CountryOrRegion } else { "" }
    if ($country) { [void]$signals.Add("Country $country") }

    if ($signals.Count -eq 0) {
        continue
    }

    [void]$mailboxRows.Add([pscustomobject]@{
        UserPrincipalName = [string]$signIn.UserPrincipalName
        CreatedDateTime   = [string]$signIn.CreatedDateTime
        AppDisplayName    = $appName
        ClientAppUsed     = $clientApp
        IPAddress         = [string]$signIn.IpAddress
        Location          = Get-GeoLocationString -Location $signIn.Location
        ConditionalAccess = [string]$signIn.ConditionalAccessStatus
        Signals           = ($signals -join '; ')
    })
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailboxLoginAnomalyReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mailbox Login Anomaly Review" -Tenant $tenantName -Subtitle "Mailbox-related sign-ins with unusual client, failure, or access patterns" -Kpis @(
    @{ label = "Mailbox Events"; value = $mailboxRows.Count; sub = "Mailbox sign-ins with anomaly signals"; cls = if ($mailboxRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Users"; value = @($mailboxRows | Select-Object -ExpandProperty UserPrincipalName -Unique).Count; sub = "Affected mailbox identities"; cls = "neutral" },
    @{ label = "Legacy Client"; value = @($mailboxRows | Where-Object { $_.ClientAppUsed -match 'IMAP|POP|SMTP|Other clients' }).Count; sub = "Legacy protocol activity"; cls = "crit" },
    @{ label = "Lookback"; value = "$LookbackDays d"; sub = "Review window"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Mailbox Sign-In Anomalies"
        badge = "$($mailboxRows.Count) event(s)"
        columns = @(
            @{ key = "UserPrincipalName"; header = "User" },
            @{ key = "CreatedDateTime"; header = "Time" },
            @{ key = "AppDisplayName"; header = "Application" },
            @{ key = "ClientAppUsed"; header = "Client" },
            @{ key = "IPAddress"; header = "IP" },
            @{ key = "Location"; header = "Location" },
            @{ key = "ConditionalAccess"; header = "CA Status"; type = "pill" },
            @{ key = "Signals"; header = "Signals" }
        )
        rows = @($mailboxRows | Sort-Object CreatedDateTime -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
