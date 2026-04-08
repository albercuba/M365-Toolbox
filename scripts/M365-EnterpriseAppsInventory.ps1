[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Application.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING ENTERPRISE APP INVENTORY"

$servicePrincipals = @(Invoke-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,displayName,appId,servicePrincipalType,publisherName,signInAudience,appRoleAssignmentRequired,accountEnabled,passwordCredentials,keyCredentials,tags&$top=999')
$rows = foreach ($sp in $servicePrincipals) {
    [pscustomobject]@{
        DisplayName        = [string]$sp.displayName
        Publisher          = [string]$sp.publisherName
        Type               = [string]$sp.servicePrincipalType
        SignInAudience     = [string]$sp.signInAudience
        AssignmentRequired = if ($sp.appRoleAssignmentRequired) { "Yes" } else { "No" }
        Enabled            = if ($sp.accountEnabled) { "Enabled" } else { "Disabled" }
        Secrets            = @($sp.passwordCredentials).Count
        Certificates       = @($sp.keyCredentials).Count
        Tags               = @($sp.tags) -join ", "
    }
}

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "EnterpriseAppsInventory" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Enterprise Apps Inventory" -Tenant $tenantName -Subtitle "Service principal inventory with publisher, assignment, and credential visibility" -Kpis @(
    @{ label = "Apps"; value = $rows.Count; sub = "Service principals"; cls = "neutral" },
    @{ label = "Enabled"; value = @($rows | Where-Object { $_.Enabled -eq "Enabled" }).Count; sub = "Enabled apps"; cls = "ok" },
    @{ label = "Assignment Required"; value = @($rows | Where-Object { $_.AssignmentRequired -eq "Yes" }).Count; sub = "Access gated"; cls = "neutral" },
    @{ label = "With Credentials"; value = @($rows | Where-Object { $_.Secrets -gt 0 -or $_.Certificates -gt 0 }).Count; sub = "Secret/cert present"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Enterprise Applications"
        badge = "$($rows.Count) apps"
        columns = @(
            @{ key = "DisplayName"; header = "Application" },
            @{ key = "Publisher"; header = "Publisher" },
            @{ key = "Type"; header = "Type" },
            @{ key = "SignInAudience"; header = "Audience" },
            @{ key = "AssignmentRequired"; header = "Assignment Required"; type = "pill" },
            @{ key = "Enabled"; header = "Enabled"; type = "pill" },
            @{ key = "Secrets"; header = "Secrets" },
            @{ key = "Certificates"; header = "Certificates" }
        )
        rows = @($rows | Sort-Object DisplayName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
