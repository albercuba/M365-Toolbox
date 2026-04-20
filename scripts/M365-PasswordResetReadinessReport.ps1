[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("User.Read.All", "UserAuthenticationMethod.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING PASSWORD RESET READINESS"

$users = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled -ErrorAction Stop)
$rows = foreach ($user in $users) {
    $methods = @()
    try {
        $methods = @(Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop)
    }
    catch {
        $methods = @()
    }

    $methodNames = @($methods | ForEach-Object { Get-GraphAuthMethodLabel -Method $_ } | Where-Object { $_ -and $_ -ne 'Password' })
    $resetReady = @($methodNames | Where-Object { $_ -match 'Phone|Email OTP|Software OATH|Authenticator App' }).Count -gt 0

    [pscustomobject]@{
        DisplayName       = [string]$user.displayName
        UserPrincipalName = [string]$user.userPrincipalName
        Enabled           = if ($user.accountEnabled) { "Enabled" } else { "Disabled" }
        ResetReady        = if ($resetReady) { "Ready" } else { "Not Ready" }
        Methods           = if ($methodNames.Count -gt 0) { ($methodNames -join ", ") } else { "None" }
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "PasswordResetReadiness" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Password Reset Readiness Report" -Tenant $tenantName -Subtitle "SSPR readiness based on registered authentication methods" -Kpis @(
    @{ label = "Users"; value = $rows.Count; sub = "Users reviewed"; cls = "neutral" },
    @{ label = "Ready"; value = @($rows | Where-Object { $_.ResetReady -eq "Ready" }).Count; sub = "Can self-reset"; cls = "ok" },
    @{ label = "Not Ready"; value = @($rows | Where-Object { $_.ResetReady -eq "Not Ready" }).Count; sub = "Needs registration"; cls = "warn" },
    @{ label = "Disabled"; value = @($rows | Where-Object { $_.Enabled -eq "Disabled" }).Count; sub = "Disabled users"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Password Reset Readiness"
        badge = "$($rows.Count) users"
        columns = @(
            @{ key = "DisplayName"; header = "Name" },
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "Enabled"; header = "Enabled"; type = "pill" },
            @{ key = "ResetReady"; header = "Readiness"; type = "pill" },
            @{ key = "Methods"; header = "Methods" }
        )
        rows = @($rows | Sort-Object ResetReady, DisplayName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
