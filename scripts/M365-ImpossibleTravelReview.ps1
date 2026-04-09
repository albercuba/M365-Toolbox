[CmdletBinding()]
param(
    [string]$TenantId,
    [int]$LookbackDays = 14,
    [int]$MaxHoursBetween = 12,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("AuditLog.Read.All", "Directory.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING SIGN-IN DATA"

$startUtc = (Get-Date).AddDays(-1 * $LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$signIns = @(Get-MgAuditLogSignIn -Filter "createdDateTime ge $startUtc" -All -ErrorAction Stop)
$travelRows = [System.Collections.Generic.List[object]]::new()

$byUser = $signIns |
    Where-Object { $_.UserPrincipalName -and $_.Location -and $_.Location.CountryOrRegion } |
    Group-Object UserPrincipalName

foreach ($group in $byUser) {
    $ordered = @($group.Group | Sort-Object CreatedDateTime)
    for ($index = 1; $index -lt $ordered.Count; $index++) {
        $previous = $ordered[$index - 1]
        $current = $ordered[$index]
        $previousCountry = [string]$previous.Location.CountryOrRegion
        $currentCountry = [string]$current.Location.CountryOrRegion
        if (-not $previousCountry -or -not $currentCountry -or $previousCountry -eq $currentCountry) {
            continue
        }

        $hours = [math]::Round((New-TimeSpan -Start $previous.CreatedDateTime -End $current.CreatedDateTime).TotalHours, 1)
        if ($hours -gt $MaxHoursBetween -or $hours -lt 0) {
            continue
        }

        [void]$travelRows.Add([pscustomobject]@{
            UserPrincipalName = [string]$group.Name
            PreviousSignIn    = [string]$previous.CreatedDateTime
            PreviousCountry   = $previousCountry
            PreviousIP        = [string]$previous.IpAddress
            CurrentSignIn     = [string]$current.CreatedDateTime
            CurrentCountry    = $currentCountry
            CurrentIP         = [string]$current.IpAddress
            ClientApp         = [string]$current.ClientAppUsed
            AppDisplayName    = [string]$current.AppDisplayName
            HoursBetween      = $hours
        })
    }
}

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "ImpossibleTravelReview" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$affectedUsers = @($travelRows | Select-Object -ExpandProperty UserPrincipalName -Unique).Count

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Impossible Travel Review" -Tenant $tenantName -Subtitle "Rapid cross-country sign-in changes across the selected lookback window" -Kpis @(
    @{ label = "Sign-Ins"; value = $signIns.Count; sub = "Audited sign-ins"; cls = "neutral" },
    @{ label = "Travel Pairs"; value = $travelRows.Count; sub = "Country changes within threshold"; cls = if ($travelRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Affected Users"; value = $affectedUsers; sub = "Users with suspicious travel"; cls = if ($affectedUsers -gt 0) { "warn" } else { "ok" } },
    @{ label = "Threshold"; value = "$MaxHoursBetween h"; sub = "Maximum allowed time gap"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Lookback"; value = "$LookbackDays day(s)" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Suspicious Travel Pairs"
        badge = "$($travelRows.Count) pair(s)"
        columns = @(
            @{ key = "UserPrincipalName"; header = "User" },
            @{ key = "PreviousSignIn"; header = "Previous Sign-In" },
            @{ key = "PreviousCountry"; header = "Previous Country" },
            @{ key = "CurrentSignIn"; header = "Current Sign-In" },
            @{ key = "CurrentCountry"; header = "Current Country" },
            @{ key = "HoursBetween"; header = "Hours Between" },
            @{ key = "ClientApp"; header = "Client App" },
            @{ key = "AppDisplayName"; header = "Application" }
        )
        rows = @($travelRows | Sort-Object HoursBetween)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
