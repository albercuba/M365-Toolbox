[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("ServiceHealth.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING SERVICE HEALTH DATA"

$health = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews')
$issues = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues')

$healthRows = foreach ($item in $health) {
    [pscustomobject]@{
        Service     = [string]$item.service
        Status      = [string]$item.status
        Feature     = [string]$item.feature
        LastUpdated = if ($item.lastModifiedDateTime) { (Get-Date $item.lastModifiedDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}

$issueRows = foreach ($issue in $issues) {
    [pscustomobject]@{
        Id             = [string]$issue.id
        Service        = [string]$issue.service
        Classification = [string]$issue.classification
        Status         = [string]$issue.status
        Title          = [string]$issue.title
        StartTime      = if ($issue.startDateTime) { (Get-Date $issue.startDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ServiceHealth" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Service Health Snapshot" -Tenant $tenantName -Subtitle "Current Microsoft 365 service status and active advisories" -Kpis @(
    @{ label = "Services"; value = $healthRows.Count; sub = "Health overviews"; cls = "neutral" },
    @{ label = "Issues"; value = $issueRows.Count; sub = "Active advisories/incidents"; cls = if ($issueRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Healthy"; value = @($healthRows | Where-Object { $_.Status -eq "serviceOperational" }).Count; sub = "Operational services"; cls = "ok" },
    @{ label = "Impacted"; value = @($healthRows | Where-Object { $_.Status -ne "serviceOperational" }).Count; sub = "Non-operational"; cls = if (@($healthRows | Where-Object { $_.Status -ne "serviceOperational" }).Count -gt 0) { "crit" } else { "ok" } }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Service Health"
        badge = "$($healthRows.Count) services"
        columns = @(
            @{ key = "Service"; header = "Service" },
            @{ key = "Status"; header = "Status"; type = "pill" },
            @{ key = "Feature"; header = "Feature" },
            @{ key = "LastUpdated"; header = "Last Updated" }
        )
        rows = @($healthRows | Sort-Object Service)
    },
    @{
        title = "Active Issues"
        badge = "$($issueRows.Count) issues"
        columns = @(
            @{ key = "Id"; header = "Issue ID" },
            @{ key = "Service"; header = "Service" },
            @{ key = "Classification"; header = "Classification"; type = "pill" },
            @{ key = "Status"; header = "Status"; type = "pill" },
            @{ key = "Title"; header = "Title" },
            @{ key = "StartTime"; header = "Start Time" }
        )
        rows = @($issueRows | Sort-Object StartTime -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
