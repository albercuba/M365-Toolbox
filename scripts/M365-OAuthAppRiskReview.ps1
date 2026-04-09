[CmdletBinding()]
param(
    [string]$TenantId,
    [int]$DaysAhead = 60,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Application.Read.All", "DelegatedPermissionGrant.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING OAUTH APP RISK DATA"

$servicePrincipals = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,displayName,appId,publisherName,appOwnerOrganizationId,accountEnabled")
$grants = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants")
$applications = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials")
$applicationByAppId = @{}
foreach ($application in $applications) {
    $applicationByAppId[[string]$application.appId] = $application
}

$rows = foreach ($sp in $servicePrincipals) {
    $spGrants = @($grants | Where-Object { $_.clientId -eq $sp.id })
    if ($spGrants.Count -eq 0) { continue }

    $app = $applicationByAppId[[string]$sp.appId]
    $riskSignals = [System.Collections.Generic.List[string]]::new()
    $scopes = (($spGrants | Select-Object -ExpandProperty scope -Unique) -join " | ")
    if ($scopes -match 'Mail\.ReadWrite|Mail\.Send|Files\.ReadWrite\.All|Sites\.FullControl\.All|Directory\.ReadWrite\.All|User\.ReadWrite\.All') {
        [void]$riskSignals.Add("High-privilege delegated scopes")
    }
    if (-not $sp.publisherName) { [void]$riskSignals.Add("No publisher name") }
    if ($app) {
        $expiringCreds = @($app.passwordCredentials + $app.keyCredentials | Where-Object { $_.endDateTime -and (Get-Date $_.endDateTime) -lt (Get-Date).AddDays($DaysAhead) })
        if ($expiringCreds.Count -gt 0) { [void]$riskSignals.Add("Credential expires within $DaysAhead days") }
    }

    [pscustomobject]@{
        Application   = [string]$sp.displayName
        AppId         = [string]$sp.appId
        Publisher     = if ($sp.publisherName) { [string]$sp.publisherName } else { "Unknown" }
        Enabled       = if ($sp.accountEnabled) { "Enabled" } else { "Disabled" }
        GrantCount    = $spGrants.Count
        Scopes        = $scopes
        RiskSignals   = if ($riskSignals.Count -gt 0) { $riskSignals -join "; " } else { "Review ownership and usage" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "OAuthAppRiskReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 OAuth App Risk Review" -Tenant $tenantName -Subtitle "OAuth apps with delegated grants, elevated scopes, and credential risk signals" -Kpis @(
    @{ label = "Apps"; value = $rows.Count; sub = "Enterprise apps with delegated grants"; cls = "neutral" },
    @{ label = "High Privilege"; value = @($rows | Where-Object { $_.RiskSignals -match 'High-privilege' }).Count; sub = "Apps with elevated scopes"; cls = "crit" },
    @{ label = "Expiring Creds"; value = @($rows | Where-Object { $_.RiskSignals -match 'Credential expires' }).Count; sub = "Apps nearing credential expiry"; cls = "warn" },
    @{ label = "Disabled"; value = @($rows | Where-Object { $_.Enabled -eq 'Disabled' }).Count; sub = "Disabled apps with grants"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Days Ahead"; value = "$DaysAhead day(s)" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "OAuth App Risk Signals"
        badge = "$($rows.Count) app(s)"
        columns = @(
            @{ key = "Application"; header = "Application" },
            @{ key = "Publisher"; header = "Publisher" },
            @{ key = "Enabled"; header = "State"; type = "pill" },
            @{ key = "GrantCount"; header = "Grant Count" },
            @{ key = "Scopes"; header = "Scopes" },
            @{ key = "RiskSignals"; header = "Risk Signals" }
        )
        rows = @($rows | Sort-Object Application)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
