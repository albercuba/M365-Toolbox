[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$DaysAhead = 60
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Application.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING APP CREDENTIAL EXPIRY DATA"

$threshold = (Get-Date).AddDays($DaysAhead)
$applications = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,displayName,appId,passwordCredentials,keyCredentials&$top=999')
$rows = [System.Collections.Generic.List[object]]::new()

foreach ($app in $applications) {
    foreach ($credential in @($app.passwordCredentials)) {
        $endDate = if ($credential.endDateTime) { Get-Date $credential.endDateTime } else { $null }
        [void]$rows.Add([pscustomobject]@{
            Application   = [string]$app.displayName
            AppId         = [string]$app.appId
            Credential    = "Secret"
            DisplayName   = [string]$credential.displayName
            EndDate       = if ($endDate) { $endDate.ToString("yyyy-MM-dd") } else { "" }
            DaysRemaining = if ($endDate) { [int][math]::Floor(($endDate - (Get-Date)).TotalDays) } else { 99999 }
            Status        = if (-not $endDate) { "Unknown" } elseif ($endDate -lt (Get-Date)) { "Expired" } elseif ($endDate -le $threshold) { "Expiring Soon" } else { "Healthy" }
        })
    }
    foreach ($credential in @($app.keyCredentials)) {
        $endDate = if ($credential.endDateTime) { Get-Date $credential.endDateTime } else { $null }
        [void]$rows.Add([pscustomobject]@{
            Application   = [string]$app.displayName
            AppId         = [string]$app.appId
            Credential    = "Certificate"
            DisplayName   = [string]$credential.displayName
            EndDate       = if ($endDate) { $endDate.ToString("yyyy-MM-dd") } else { "" }
            DaysRemaining = if ($endDate) { [int][math]::Floor(($endDate - (Get-Date)).TotalDays) } else { 99999 }
            Status        = if (-not $endDate) { "Unknown" } elseif ($endDate -lt (Get-Date)) { "Expired" } elseif ($endDate -le $threshold) { "Expiring Soon" } else { "Healthy" }
        })
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "AppCredentialExpiry" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 App Credential Expiry Report" -Tenant $tenantName -Subtitle "Secrets and certificates nearing expiry across app registrations" -Kpis @(
    @{ label = "Credentials"; value = $rows.Count; sub = "Inspected app creds"; cls = "neutral" },
    @{ label = "Expired"; value = @($rows | Where-Object { $_.Status -eq "Expired" }).Count; sub = "Past due"; cls = "crit" },
    @{ label = "Expiring Soon"; value = @($rows | Where-Object { $_.Status -eq "Expiring Soon" }).Count; sub = "Within $DaysAhead days"; cls = "warn" },
    @{ label = "Apps"; value = @($rows | Select-Object -ExpandProperty Application -Unique).Count; sub = "Affected apps"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Window"; value = "$DaysAhead days" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Credential Expiry"
        badge = "$($rows.Count) credentials"
        columns = @(
            @{ key = "Application"; header = "Application" },
            @{ key = "Credential"; header = "Credential Type"; type = "pill" },
            @{ key = "DisplayName"; header = "Credential Name" },
            @{ key = "EndDate"; header = "End Date" },
            @{ key = "DaysRemaining"; header = "Days Remaining" },
            @{ key = "Status"; header = "Status"; type = "pill" }
        )
        rows = @($rows | Sort-Object DaysRemaining, Application)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
