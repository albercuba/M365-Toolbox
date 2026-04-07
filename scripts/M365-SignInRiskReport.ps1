[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$LookbackDays = 30
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("IdentityRiskyUser.Read.All", "IdentityRiskEvent.Read.All", "AuditLog.Read.All", "User.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING SIGN-IN RISK DATA"

$riskFeatureUnavailable = $false
$riskFeatureMessage = ""

try {
    $riskyUsers = @(Get-MgRiskyUser -All -ErrorAction Stop)
}
catch {
    $message = $_.Exception.Message
    if ($message -match "not licensed for this feature" -or $message -match "403 \(Forbidden\)") {
        $riskFeatureUnavailable = $true
        $riskFeatureMessage = "Microsoft Entra ID Identity Protection data is not available for this tenant or license level."
        Write-Warning "  [!] $riskFeatureMessage"
        $riskyUsers = @()
    }
    else {
        throw
    }
}

try {
    $riskDetections = if ($riskFeatureUnavailable) { @() } else { @(Get-MgRiskDetection -All -ErrorAction Stop) }
}
catch {
    $message = $_.Exception.Message
    if ($message -match "not licensed for this feature" -or $message -match "403 \(Forbidden\)") {
        $riskFeatureUnavailable = $true
        if (-not $riskFeatureMessage) {
            $riskFeatureMessage = "Microsoft Entra ID Identity Protection data is not available for this tenant or license level."
            Write-Warning "  [!] $riskFeatureMessage"
        }
        $riskDetections = @()
    }
    else {
        throw
    }
}

$cutoff = (Get-Date).AddDays(-1 * $LookbackDays)

$filteredDetections = @(
    $riskDetections | Where-Object {
        if (-not $_.activityDateTime) { return $true }
        (Get-Date $_.activityDateTime) -ge $cutoff
    }
)

$userRows = @(
foreach ($user in $riskyUsers) {
    [pscustomobject]@{
        UserPrincipalName = [string]$user.userPrincipalName
        RiskLevel         = [string]$user.riskLevel
        RiskState         = [string]$user.riskState
        Detail            = [string]$user.riskDetail
        Updated           = if ($user.riskLastUpdatedDateTime) { (Get-Date $user.riskLastUpdatedDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}
)

$detectionRows = @(
foreach ($detection in $filteredDetections) {
    [pscustomobject]@{
        UserPrincipalName = [string]$detection.userPrincipalName
        RiskType          = [string]$detection.riskType
        RiskLevel         = [string]$detection.riskLevel
        DetectionTiming   = [string]$detection.detectionTimingType
        IPAddress         = [string]$detection.ipAddress
        Activity          = if ($detection.activityDateTime) { (Get-Date $detection.activityDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
    }
}
)

$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$highRiskUsers = @($userRows | Where-Object { $_.RiskLevel -eq "high" }).Count
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "SignInRisk" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Sign-In Risk Report" -Tenant $tenantName -Subtitle "Risky users and detections from Microsoft Entra ID" -Kpis @(
    @{ label = "Risky Users"; value = $userRows.Count; sub = "Current risky identities"; cls = if ($userRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "High Risk"; value = $highRiskUsers; sub = "Users at high risk"; cls = if ($highRiskUsers -gt 0) { "crit" } else { "ok" } },
    @{ label = "Detections"; value = $detectionRows.Count; sub = "Within $LookbackDays days"; cls = if ($detectionRows.Count -gt 0) { "warn" } else { "ok" } },
    @{ label = "Window"; value = $LookbackDays; sub = "Days reviewed"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Lookback"; value = "$LookbackDays days" },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "License Status"
        badge = if ($riskFeatureUnavailable) { "Unavailable" } else { "Available" }
        text = if ($riskFeatureUnavailable) {
            $riskFeatureMessage
        }
        else {
            "Identity Protection data was available and collected successfully."
        }
    },
    @{
        title = "Risky Users"
        badge = "$($userRows.Count) users"
        columns = @(
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "RiskLevel"; header = "Risk Level"; type = "pill" },
            @{ key = "RiskState"; header = "Risk State"; type = "pill" },
            @{ key = "Detail"; header = "Detail" },
            @{ key = "Updated"; header = "Updated" }
        )
        rows = @($userRows | Sort-Object RiskLevel, UserPrincipalName)
    },
    @{
        title = "Risk Detections"
        badge = "$($detectionRows.Count) detections"
        columns = @(
            @{ key = "UserPrincipalName"; header = "UPN" },
            @{ key = "RiskType"; header = "Risk Type" },
            @{ key = "RiskLevel"; header = "Level"; type = "pill" },
            @{ key = "DetectionTiming"; header = "Timing" },
            @{ key = "IPAddress"; header = "IP Address" },
            @{ key = "Activity"; header = "Activity" }
        )
        rows = @($detectionRows | Sort-Object Activity -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
