[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$LookbackDays = 30
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("AuditLog.Read.All", "User.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING LEGACY AUTH EXPOSURE"

$cutoff = (Get-Date).ToUniversalTime().AddDays(-1 * $LookbackDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
$signIns = @(Get-MgAuditLogSignIn -Filter "createdDateTime ge $cutoff" -All -Property UserPrincipalName,AppDisplayName,ClientAppUsed,CreatedDateTime,ConditionalAccessStatus,IpAddress -ErrorAction Stop)
$legacyClients = @("imap","pop","smtp","mapi","exchange active sync","other clients","autodiscover","exchange web services")

$rows = foreach ($entry in $signIns) {
    $client = [string]$entry.clientAppUsed
    if (-not $client) { continue }
    if (($legacyClients | Where-Object { $client.ToLower().Contains($_) }).Count -eq 0) { continue }

    [pscustomobject]@{
        UserPrincipalName       = [string]$entry.userPrincipalName
        ClientApp               = $client
        Application             = [string]$entry.appDisplayName
        ConditionalAccessStatus = [string]$entry.conditionalAccessStatus
        IPAddress               = [string]$entry.ipAddress
        Activity                = if ($entry.createdDateTime) { (Get-Date $entry.createdDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "LegacyAuthExposure" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Legacy Auth Exposure Report" -Tenant $tenantName -Subtitle "Legacy authentication sign-ins within the selected lookback window" -Kpis @(
    @{ label = "Sign-ins"; value = $rows.Count; sub = "Legacy auth activity"; cls = if ($rows.Count -gt 0) { "crit" } else { "ok" } },
    @{ label = "Users"; value = @($rows | Select-Object -ExpandProperty UserPrincipalName -Unique).Count; sub = "Affected identities"; cls = if ($rows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Clients"; value = @($rows | Select-Object -ExpandProperty ClientApp -Unique).Count; sub = "Legacy protocols"; cls = "warn" },
    @{ label = "Window"; value = $LookbackDays; sub = "Days"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Lookback"; value = "$LookbackDays days" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Legacy Authentication Sign-Ins"
        badge = "$($rows.Count) entries"
        columns = @(
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "ClientApp"; header = "Client" },
            @{ key = "Application"; header = "Application" },
            @{ key = "ConditionalAccessStatus"; header = "CA Status"; type = "pill" },
            @{ key = "IPAddress"; header = "IP Address" },
            @{ key = "Activity"; header = "Activity" }
        )
        rows = @($rows | Sort-Object Activity -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
