[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("SecurityIncident.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING DEFENDER INCIDENTS"

$incidents = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/security/incidents?$top=200')
$rows = foreach ($incident in $incidents) {
    [pscustomobject]@{
        IncidentId = [string]$incident.incidentNumber
        DisplayName = [string]$incident.displayName
        Severity   = [string]$incident.severity
        Status     = [string]$incident.status
        Classification = [string]$incident.classification
        Created    = if ($incident.createdDateTime) { (Get-Date $incident.createdDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "DefenderIncidents" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Defender Incident Snapshot" -Tenant $tenantName -Subtitle "Current incidents from Microsoft Defender / Graph Security" -Kpis @(
    @{ label = "Incidents"; value = $rows.Count; sub = "Returned incidents"; cls = if ($rows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "High"; value = @($rows | Where-Object { $_.Severity -eq "high" }).Count; sub = "High severity"; cls = "crit" },
    @{ label = "Active"; value = @($rows | Where-Object { $_.Status -ne "resolved" }).Count; sub = "Not resolved"; cls = "warn" },
    @{ label = "Resolved"; value = @($rows | Where-Object { $_.Status -eq "resolved" }).Count; sub = "Resolved"; cls = "ok" }
    ) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Incident Snapshot"
        badge = "$($rows.Count) incidents"
        columns = @(
            @{ key = "IncidentId"; header = "Incident" },
            @{ key = "DisplayName"; header = "Title" },
            @{ key = "Severity"; header = "Severity"; type = "pill" },
            @{ key = "Status"; header = "Status"; type = "pill" },
            @{ key = "Classification"; header = "Classification" },
            @{ key = "Created"; header = "Created" }
        )
        rows = @($rows | Sort-Object Created -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
