[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [int]$TopActions = 15
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("SecurityEvents.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING SECURE SCORE DATA"

$scoreResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1' -ErrorAction Stop
$score = @($scoreResponse.value | Sort-Object createdDateTime -Descending | Select-Object -First 1)[0]
$controlResponse = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?`$top={0}" -f $TopActions) -ErrorAction Stop
$controls = @($controlResponse.value)

$controlRows = foreach ($control in $controls) {
    [pscustomobject]@{
        Title         = [string]$control.title
        Category      = [string]$control.controlCategory
        MaxScore      = [string]$control.maxScore
        UserImpact    = [string]$control.userImpact
        Implementation = [string]$control.implementationCost
    }
}

$currentScore = [double]$score.currentScore
$maxScore = [double]$score.maxScore
$percent = if ($maxScore -gt 0) { [math]::Round(($currentScore / $maxScore) * 100, 1) } else { 0 }
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }
$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "SecureScore" -OutputPath $OutputPath

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Secure Score Snapshot" -Tenant $tenantName -Subtitle "Current Secure Score and top control opportunities" -Kpis @(
    @{ label = "Current"; value = $currentScore; sub = "Current secure score"; cls = "neutral" },
    @{ label = "Max"; value = $maxScore; sub = "Maximum score"; cls = "neutral" },
    @{ label = "Percent"; value = "$percent%"; sub = "Score attainment"; cls = if ($percent -lt 50) { "warn" } elseif ($percent -lt 75) { "neutral" } else { "ok" } },
    @{ label = "Controls"; value = $controlRows.Count; sub = "Improvement actions"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Snapshot"; value = if ($score.createdDateTime) { (Get-Date $score.createdDateTime).ToString("yyyy-MM-dd HH:mm") } else { "Latest" } },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Score Snapshot"
        badge = "$percent%"
        text = if ($score.averageComparativeScores) { "Secure Score snapshot captured successfully." } else { "Latest secure score snapshot captured successfully." }
    },
    @{
        title = "Control Opportunities"
        badge = "$($controlRows.Count) controls"
        columns = @(
            @{ key = "Title"; header = "Control" },
            @{ key = "Category"; header = "Category" },
            @{ key = "MaxScore"; header = "Max Score" },
            @{ key = "UserImpact"; header = "User Impact" },
            @{ key = "Implementation"; header = "Implementation Cost" }
        )
        rows = @($controlRows | Sort-Object MaxScore -Descending)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
