[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Application.Read.All", "DelegatedPermissionGrant.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING APP CONSENT DATA"

$servicePrincipals = @(Get-MgServicePrincipal -All -Property Id,DisplayName,AppId,PublisherName,AppOwnerOrganizationId -ErrorAction Stop)
$grants = @(Get-MgOauth2PermissionGrant -All -ErrorAction Stop)
$assignmentRows = [System.Collections.Generic.List[object]]::new()

foreach ($sp in $servicePrincipals) {
    $spGrants = @($grants | Where-Object { $_.clientId -eq $sp.id })
    if ($spGrants.Count -eq 0) { continue }

    [void]$assignmentRows.Add([pscustomobject]@{
        DisplayName  = [string]$sp.displayName
        AppId        = [string]$sp.appId
        Publisher    = [string]$sp.publisherName
        GrantCount   = $spGrants.Count
        ConsentType  = (($spGrants | Select-Object -ExpandProperty consentType -Unique) -join ", ")
        ScopePreview = (($spGrants | Select-Object -ExpandProperty scope -Unique) -join " | ")
    })
}

$highPrivilegeRows = @($assignmentRows | Where-Object { $_.ScopePreview -match 'Mail\.Read|Mail\.ReadWrite|Files\.ReadWrite|Directory\.ReadWrite|User\.ReadWrite|offline_access' })
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "AppConsentAudit" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 App Consent Audit" -Tenant $tenantName -Subtitle "Enterprise app consents and delegated permission exposure" -Kpis @(
    @{ label = "Apps"; value = $assignmentRows.Count; sub = "Apps with grants"; cls = "neutral" },
    @{ label = "Grants"; value = $grants.Count; sub = "OAuth permission grants"; cls = "neutral" },
    @{ label = "High Privilege"; value = $highPrivilegeRows.Count; sub = "Apps with elevated scopes"; cls = if ($highPrivilegeRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Service Principals"; value = $servicePrincipals.Count; sub = "Total enterprise apps"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Apps With Delegated Consent"
        badge = "$($assignmentRows.Count) apps"
        columns = @(
            @{ key = "DisplayName"; header = "Application" },
            @{ key = "Publisher"; header = "Publisher" },
            @{ key = "GrantCount"; header = "Grant Count" },
            @{ key = "ConsentType"; header = "Consent Type" },
            @{ key = "ScopePreview"; header = "Scopes" }
        )
        rows = @($assignmentRows | Sort-Object GrantCount -Descending)
    },
    @{
        title = "High Privilege Consents"
        badge = "$($highPrivilegeRows.Count) apps"
        columns = @(
            @{ key = "DisplayName"; header = "Application" },
            @{ key = "Publisher"; header = "Publisher" },
            @{ key = "GrantCount"; header = "Grant Count" },
            @{ key = "ScopePreview"; header = "Scopes" }
        )
        rows = $highPrivilegeRows
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
