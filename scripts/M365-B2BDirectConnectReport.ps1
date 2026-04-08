[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING CROSS-TENANT ACCESS DATA"

$policy = $null
$policyWarning = $null
try {
    $policy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy' -ErrorAction Stop
}
catch {
    $policy = $null
    $policyWarning = $_.Exception.Message
    Write-Warning "  [!] Cross-tenant access policy could not be retrieved. $policyWarning"
}

$rows = @(
    [pscustomobject]@{ Setting = "Default Inbound B2B"; Value = if ($policy.default?.b2bCollaborationInbound?.applicationsAndServices?.accessType) { [string]$policy.default.b2bCollaborationInbound.applicationsAndServices.accessType } else { "Unknown" } },
    [pscustomobject]@{ Setting = "Default Outbound B2B"; Value = if ($policy.default?.b2bCollaborationOutbound?.applicationsAndServices?.accessType) { [string]$policy.default.b2bCollaborationOutbound.applicationsAndServices.accessType } else { "Unknown" } },
    [pscustomobject]@{ Setting = "Tenant Restrictions"; Value = if ($policy.tenantRestrictions?.devices?.trustType) { [string]$policy.tenantRestrictions.devices.trustType } else { "Unknown" } },
    [pscustomobject]@{ Setting = "Partner Count"; Value = if ($policy.partners) { @($policy.partners.PSObject.Properties).Count } else { 0 } }
)

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "B2BDirectConnect" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 B2B Direct Connect Report" -Tenant $tenantName -Subtitle "Cross-tenant access policy defaults and partner posture" -Kpis @(
    @{ label = "Partners"; value = if ($policy.partners) { @($policy.partners.PSObject.Properties).Count } else { 0 }; sub = "Configured partners"; cls = "neutral" },
    @{ label = "Inbound"; value = if ($policy.default?.b2bCollaborationInbound?.applicationsAndServices?.accessType) { [string]$policy.default.b2bCollaborationInbound.applicationsAndServices.accessType } else { "Unknown" }; sub = "Default inbound"; cls = "warn" },
    @{ label = "Outbound"; value = if ($policy.default?.b2bCollaborationOutbound?.applicationsAndServices?.accessType) { [string]$policy.default.b2bCollaborationOutbound.applicationsAndServices.accessType } else { "Unknown" }; sub = "Default outbound"; cls = "warn" },
    @{ label = "Policy"; value = if ($policy) { "Available" } else { "Unavailable" }; sub = "Graph response"; cls = if ($policy) { "ok" } else { "warn" } }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Collection Status"
        badge = if ($policyWarning) { "Warning" } else { "Healthy" }
        text = if ($policyWarning) { "Cross-tenant access policy could not be retrieved: $policyWarning" } else { "Cross-tenant access policy data was collected successfully." }
    },
    @{
        title = "Cross-Tenant Access Defaults"
        badge = "Policy"
        columns = @(
            @{ key = "Setting"; header = "Setting" },
            @{ key = "Value"; header = "Value"; type = "pill" }
        )
        rows = $rows
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
