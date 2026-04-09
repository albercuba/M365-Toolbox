[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Policy.Read.All", "Directory.Read.All", "User.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING CROSS-TENANT TRUST DATA"

$partnerPolicies = @(Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners")
$guests = @(Get-MgUser -Filter "userType eq 'Guest'" -All -Property Id,DisplayName,UserPrincipalName,Mail,ExternalUserState)

$guestDomains = $guests |
    ForEach-Object {
        $source = if ($_.Mail) { [string]$_.Mail } else { [string]$_.UserPrincipalName }
        if ($source -match '@') {
            [pscustomobject]@{ Domain = ($source.Split('@')[-1]).ToLowerInvariant() }
        }
    } |
    Where-Object { $_ } |
    Group-Object Domain |
    ForEach-Object {
        [pscustomobject]@{
            Domain      = $_.Name
            GuestCount  = $_.Count
        }
    } |
    Sort-Object GuestCount -Descending

$partnerRows = foreach ($partner in $partnerPolicies) {
    [pscustomobject]@{
        TenantId          = [string]$partner.tenantId
        Name              = [string]$partner.displayName
        InboundDefaults   = [string]$partner.inboundTrust?.isMfaAccepted
        OutboundDefaults  = [string]$partner.outboundTrust?.isMfaAccepted
        B2BCollaboration  = [string]$partner.b2bCollaborationInbound?.applicationsAndUsers?.accessType
        B2BDirectConnect  = [string]$partner.b2bDirectConnectInbound?.applicationsAndUsers?.accessType
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ExternalTenantTrustReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 External Tenant Trust Review" -Tenant $tenantName -Subtitle "Cross-tenant access partners and guest domain concentration" -Kpis @(
    @{ label = "Partners"; value = $partnerRows.Count; sub = "Configured cross-tenant partners"; cls = "neutral" },
    @{ label = "Guests"; value = $guests.Count; sub = "Guest accounts in tenant"; cls = "neutral" },
    @{ label = "Guest Domains"; value = @($guestDomains).Count; sub = "Distinct external domains"; cls = "neutral" },
    @{ label = "MFA Trust"; value = @($partnerRows | Where-Object { $_.InboundDefaults -eq 'True' -or $_.OutboundDefaults -eq 'True' }).Count; sub = "Partners trusting MFA claims"; cls = "warn" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Cross-Tenant Partners"
        badge = "$($partnerRows.Count) partner(s)"
        columns = @(
            @{ key = "TenantId"; header = "Tenant ID" },
            @{ key = "Name"; header = "Display Name" },
            @{ key = "InboundDefaults"; header = "Inbound MFA Trust"; type = "pill" },
            @{ key = "OutboundDefaults"; header = "Outbound MFA Trust"; type = "pill" },
            @{ key = "B2BCollaboration"; header = "B2B Collaboration" },
            @{ key = "B2BDirectConnect"; header = "B2B Direct Connect" }
        )
        rows = @($partnerRows | Sort-Object Name)
    },
    @{
        title = "Guest Domain Concentration"
        badge = "$(@($guestDomains).Count) domain(s)"
        columns = @(
            @{ key = "Domain"; header = "Domain" },
            @{ key = "GuestCount"; header = "Guest Count" }
        )
        rows = @($guestDomains | Select-Object -First 50)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
