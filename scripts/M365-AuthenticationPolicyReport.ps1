[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING AUTHENTICATION POLICY DATA"

$securityDefaults = $null
$authMethodsPolicy = $null
$microsoftAuthenticator = $null

try { $securityDefaults = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -ErrorAction Stop } catch {}
try { $authMethodsPolicy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy' -ErrorAction Stop } catch {}
try { $microsoftAuthenticator = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/MicrosoftAuthenticator' -ErrorAction Stop } catch {}

$rows = @(
    [pscustomobject]@{ Setting = "Security Defaults"; Value = if ($securityDefaults.isEnabled) { "Enabled" } else { "Disabled" } },
    [pscustomobject]@{ Setting = "Authentication Methods Policy"; Value = if ($authMethodsPolicy.policyVersion) { "Configured" } else { "Unknown" } },
    [pscustomobject]@{ Setting = "Microsoft Authenticator State"; Value = if ($microsoftAuthenticator.state) { [string]$microsoftAuthenticator.state } else { "Unknown" } },
    [pscustomobject]@{ Setting = "Policy Version"; Value = if ($authMethodsPolicy.policyVersion) { [string]$authMethodsPolicy.policyVersion } else { "Unknown" } }
)

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "AuthenticationPolicies" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Authentication Policy Report" -Tenant $tenantName -Subtitle "Security defaults and authentication methods policy posture" -Kpis @(
    @{ label = "Security Defaults"; value = if ($securityDefaults.isEnabled) { "Enabled" } else { "Disabled" }; sub = "Baseline protection"; cls = if ($securityDefaults.isEnabled) { "ok" } else { "warn" } },
    @{ label = "Policy Version"; value = if ($authMethodsPolicy.policyVersion) { [string]$authMethodsPolicy.policyVersion } else { "Unknown" }; sub = "Auth methods"; cls = "neutral" },
    @{ label = "Configurations"; value = if ($authMethodsPolicy.authenticationMethodConfigurations) { @($authMethodsPolicy.authenticationMethodConfigurations).Count } else { 0 }; sub = "Method configs"; cls = "neutral" },
    @{ label = "Review Needed"; value = if (-not $securityDefaults.isEnabled) { "Yes" } else { "No" }; sub = "Baseline controls"; cls = if (-not $securityDefaults.isEnabled) { "warn" } else { "ok" } }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Authentication Policy Settings"
        badge = "Tenant posture"
        columns = @(
            @{ key = "Setting"; header = "Setting" },
            @{ key = "Value"; header = "Value"; type = "pill" }
        )
        rows = $rows
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
