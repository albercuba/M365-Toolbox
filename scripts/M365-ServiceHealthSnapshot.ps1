[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Devices.ServiceAnnouncement")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("ServiceHealth.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING SERVICE HEALTH DATA"

$health = @(Get-MgServiceAnnouncementHealthOverview -All -ErrorAction Stop)
$issues = @(Get-MgServiceAnnouncementIssue -All -ErrorAction Stop)

$healthRows = foreach ($item in $health) {
    $lastModifiedDateTime = Get-DirectoryObjectValue -DirectoryObject $item -Name 'lastModifiedDateTime'
    [pscustomobject]@{
        Service     = [string](Get-DirectoryObjectValue -DirectoryObject $item -Name 'service')
        Status      = [string](Get-DirectoryObjectValue -DirectoryObject $item -Name 'status')
        Feature     = [string](Get-DirectoryObjectValue -DirectoryObject $item -Name 'feature')
        LastUpdated = if ($lastModifiedDateTime) { (Get-Date $lastModifiedDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}

$issueRows = foreach ($issue in $issues) {
    $startDateTime = Get-DirectoryObjectValue -DirectoryObject $issue -Name 'startDateTime'
    [pscustomobject]@{
        Id             = [string](Get-DirectoryObjectValue -DirectoryObject $issue -Name 'id')
        Service        = [string](Get-DirectoryObjectValue -DirectoryObject $issue -Name 'service')
        Classification = [string](Get-DirectoryObjectValue -DirectoryObject $issue -Name 'classification')
        Status         = [string](Get-DirectoryObjectValue -DirectoryObject $issue -Name 'status')
        Title          = [string](Get-DirectoryObjectValue -DirectoryObject $issue -Name 'title')
        StartTime      = if ($startDateTime) { (Get-Date $startDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
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
