[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Application.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING PRIVILEGED APP DATA"

$servicePrincipals = @(Get-MgServicePrincipal -All -Property Id,DisplayName,AppId,PublisherName,PasswordCredentials,KeyCredentials -ErrorAction Stop)
$rows = foreach ($sp in $servicePrincipals) {
    $passwordCount = @($sp.passwordCredentials).Count
    $keyCount = @($sp.keyCredentials).Count
    [pscustomobject]@{
        DisplayName    = [string]$sp.displayName
        AppId          = [string]$sp.appId
        Publisher      = [string]$sp.publisherName
        Secrets        = $passwordCount
        Certificates   = $keyCount
        CredentialRisk = if (($passwordCount + $keyCount) -gt 3) { "Elevated" } elseif (($passwordCount + $keyCount) -gt 0) { "Present" } else { "None" }
    }
}

$withSecrets = @($rows | Where-Object { $_.Secrets -gt 0 -or $_.Certificates -gt 0 })
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "PrivilegedApps" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Privileged App Audit" -Tenant $tenantName -Subtitle "Enterprise app credentials and non-human identity visibility" -Kpis @(
    @{ label = "Apps"; value = $rows.Count; sub = "Service principals"; cls = "neutral" },
    @{ label = "With Credentials"; value = $withSecrets.Count; sub = "Apps with secrets/certs"; cls = if ($withSecrets.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Secrets"; value = (($rows | Measure-Object Secrets -Sum).Sum); sub = "Password credentials"; cls = "warn" },
    @{ label = "Certificates"; value = (($rows | Measure-Object Certificates -Sum).Sum); sub = "Key credentials"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Application Credential Inventory"
        badge = "$($rows.Count) apps"
        columns = @(
            @{ key = "DisplayName"; header = "Application" },
            @{ key = "Publisher"; header = "Publisher" },
            @{ key = "Secrets"; header = "Secrets" },
            @{ key = "Certificates"; header = "Certificates" },
            @{ key = "CredentialRisk"; header = "Risk"; type = "pill" }
        )
        rows = @($rows | Sort-Object Secrets -Descending, Certificates -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
