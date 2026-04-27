[CmdletBinding()]
param(
    [string]$TenantId,

    [string[]]$Reports = @("OneDrive", "SharePoint", "Mailbox"),

    [string]$OutputPath,

    [string]$ExportHtml
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$script:TenantLabel = ""

function Normalize-ReportSelection {
    param([string[]]$SelectedReports)

    $allowedReports = @("OneDrive", "SharePoint", "Mailbox")
    $normalized = @(
        $SelectedReports |
            Where-Object { $_ } |
            ForEach-Object { $_ -split "," } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ } |
            Select-Object -Unique
    )

    if (-not $normalized -or $normalized.Count -eq 0) {
        return @($allowedReports)
    }

    $invalid = @($normalized | Where-Object { $_ -notin $allowedReports })
    if ($invalid.Count -gt 0) {
        throw "Invalid report selection: $($invalid -join ', '). Allowed values: $($allowedReports -join ', ')."
    }

    return $normalized
}

function Write-SectionHeader {
    param([string]$Title)

    $line = "=" * 50
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host (" " * [Math]::Floor((50 - $Title.Length) / 2) + $Title) -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Add-TimestampToPath {
    param(
        [string]$Path,
        [string]$BaseName,
        [string]$Extension = ".html"
    )

    $stamp = Get-Date -Format "dd.MM.yy-HH.mm.ss"
    $tenant = if ($script:TenantLabel) { $script:TenantLabel.Split(".")[0] } else { "M365" }
    $fileName = "{0}_{1}_{2}{3}" -f $BaseName, $tenant, $stamp, $Extension

    if (-not $Path) {
        if ($OutputPath) {
            return Join-Path $OutputPath $fileName
        }
        return Join-Path (Get-Location) $fileName
    }

    if (Test-Path $Path -PathType Container) {
        return Join-Path $Path $fileName
    }

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if ($directory) {
        return Join-Path $directory $fileName
    }

    return $fileName
}

function Assert-RequiredModules {
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Reports",
        "Microsoft.Graph.Sites"
    )

    Write-Host ""
    Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan

    $commonVersions = $null
    foreach ($moduleName in $requiredModules) {
        $installed = @(Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending)
        if (-not $installed -or $installed.Count -eq 0) {
            throw "Required module '$moduleName' is not installed."
        }

        $versionStrings = @($installed | ForEach-Object { $_.Version.ToString() })
        if ($null -eq $commonVersions) {
            $commonVersions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($versionString in $versionStrings) {
                [void]$commonVersions.Add($versionString)
            }
        }
        else {
            $nextCommonVersions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($versionString in $versionStrings) {
                if ($commonVersions.Contains($versionString)) {
                    [void]$nextCommonVersions.Add($versionString)
                }
            }
            $commonVersions = $nextCommonVersions
        }
    }

    if (-not $commonVersions -or $commonVersions.Count -eq 0) {
        throw "No common Microsoft Graph module version is installed across: $($requiredModules -join ', ')."
    }

    $selectedVersion = $commonVersions |
        ForEach-Object { [version]$_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    foreach ($moduleName in $requiredModules) {
        Import-Module $moduleName -RequiredVersion $selectedVersion -Force -WarningAction SilentlyContinue
        Write-Host "  [+] $moduleName v$selectedVersion" -ForegroundColor Green
    }
}

function Connect-ToGraph {
    Write-SectionHeader "CONNECTING TO MICROSOFT GRAPH"

    if ($TenantId) {
        Write-Host "[*] Requested tenant: $TenantId"
    }
    else {
        Write-Host "[*] Requested tenant: auto-detect"
    }

    Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
    Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

    Disconnect-MgGraph -ErrorAction SilentlyContinue

    $connectParams = @{
        Scopes       = @("Reports.Read.All", "Sites.Read.All", "Directory.Read.All")
        NoWelcome    = $true
        ContextScope = "Process"
        ErrorAction  = "Stop"
    }

    if ($TenantId) {
        $connectParams.TenantId = $TenantId
    }

    try {
        Connect-MgGraph @connectParams -UseDeviceAuthentication
        $context = Get-MgContext
        if ($context.TenantId) {
            $script:TenantLabel = [string]$context.TenantId
        }
        Write-Host "[+] Connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        throw "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    }
}

function Resolve-TenantLabel {
    try {
        $initialDomain = Get-MgDomain -All -ErrorAction Stop | Where-Object { $_.IsInitial } | Select-Object -First 1
        if ($initialDomain -and $initialDomain.Id) {
            $script:TenantLabel = [string]$initialDomain.Id
        }
    }
    catch {
        if (-not $script:TenantLabel) {
            $context = Get-MgContext
            if ($context -and $context.TenantId) {
                $script:TenantLabel = [string]$context.TenantId
            }
        }
    }
}

function ConvertTo-GB {
    param([double]$Bytes)

    return [math]::Round($Bytes / 1GB, 2)
}

function Import-ReportCsv {
    param(
        [Parameter(Mandatory)]
        [string]$RequestUri
    )

    $tempFile = [System.IO.Path]::GetTempFileName()

    try {
        # The generated Get-MgReport*Detail cmdlets can emit invalid progress values on some
        # PowerShell / Graph module combinations. Using Invoke-MgGraphRequest avoids that bug
        # while still downloading the same CSV report payload.
        Invoke-MgGraphRequest -Method GET -Uri $RequestUri -OutputFilePath $tempFile -ErrorAction Stop | Out-Null
        $rawContent = [System.IO.File]::ReadAllText($tempFile, [System.Text.Encoding]::UTF8)
        $rawContent = $rawContent -replace "^[^R]*Report Refresh Date", "Report Refresh Date"
        return @($rawContent | ConvertFrom-Csv)
    }
    finally {
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }
}

function Get-OneDriveReport {
    Write-SectionHeader "ONEDRIVE USAGE REPORT"
    Write-Host "[*] Fetching OneDrive usage data from the Reports API..." -ForegroundColor Cyan

    $rows = Import-ReportCsv -RequestUri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')"
    $report = [System.Collections.Generic.List[object]]::new()

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  [!] No OneDrive usage rows were returned."
        return $report
    }

    $counter = 0
    foreach ($row in $rows) {
        $counter++
        if ($counter -eq 1 -or $counter % 50 -eq 0 -or $counter -eq $rows.Count) {
            Write-Host "  [*] Processing row $counter / $($rows.Count) ..." -ForegroundColor Cyan
        }

        $usedGb = 0
        if ([string]$row.'Storage Used (Byte)' -match '^\d+$') {
            $usedGb = ConvertTo-GB -Bytes ([double]$row.'Storage Used (Byte)')
        }

        $report.Add([PSCustomObject]@{
            DisplayName = [string]$row.'Owner Display Name'
            Principal   = [string]$row.'Owner Principal Name'
            UsedGB      = $usedGb
        }) | Out-Null
    }

    Write-Host "  [+] OneDrive report complete. $($report.Count) row(s)." -ForegroundColor Green
    return $report
}

function Get-SharePointReport {
    Write-SectionHeader "SHAREPOINT USAGE REPORT"
    Write-Host "[*] Fetching SharePoint usage data from the Reports API..." -ForegroundColor Cyan

    $rows = Import-ReportCsv -RequestUri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')"
    $report = [System.Collections.Generic.List[object]]::new()

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  [!] No SharePoint usage rows were returned."
        return $report
    }

    $counter = 0
    foreach ($row in $rows) {
        $counter++
        if ($counter -eq 1 -or $counter % 50 -eq 0 -or $counter -eq $rows.Count) {
            Write-Host "  [*] Processing row $counter / $($rows.Count) ..." -ForegroundColor Cyan
        }

        $siteUrl = [string]$row.'Site URL'
        if ([string]::IsNullOrWhiteSpace($siteUrl) -and -not [string]::IsNullOrWhiteSpace([string]$row.'Site Id')) {
            try {
                $site = Get-MgSite -SiteId ([string]$row.'Site Id') -Property WebUrl -ErrorAction Stop
                if ($site.WebUrl) {
                    $siteUrl = [string]$site.WebUrl
                }
            }
            catch {
            }
        }

        $usedGb = 0
        if ([string]$row.'Storage Used (Byte)' -match '^\d+$') {
            $usedGb = ConvertTo-GB -Bytes ([double]$row.'Storage Used (Byte)')
        }

        $report.Add([PSCustomObject]@{
            DisplayName = if ($siteUrl) { $siteUrl } else { [string]$row.'Site Id' }
            Principal   = [string]$row.'Owner Principal Name'
            UsedGB      = $usedGb
            Url         = $siteUrl
        }) | Out-Null
    }

    Write-Host "  [+] SharePoint report complete. $($report.Count) row(s)." -ForegroundColor Green
    return $report
}

function Get-MailboxReport {
    Write-SectionHeader "MAILBOX USAGE REPORT"
    Write-Host "[*] Fetching mailbox usage data from the Reports API..." -ForegroundColor Cyan

    $rows = Import-ReportCsv -RequestUri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')"
    $report = [System.Collections.Generic.List[object]]::new()

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  [!] No mailbox usage rows were returned."
        return $report
    }

    $counter = 0
    foreach ($row in $rows) {
        $counter++
        if ($counter -eq 1 -or $counter % 50 -eq 0 -or $counter -eq $rows.Count) {
            Write-Host "  [*] Processing row $counter / $($rows.Count) ..." -ForegroundColor Cyan
        }

        $usedGb = 0
        if ([string]$row.'Storage Used (Byte)' -match '^\d+$') {
            $usedGb = ConvertTo-GB -Bytes ([double]$row.'Storage Used (Byte)')
        }

        $report.Add([PSCustomObject]@{
            DisplayName = [string]$row.'Display Name'
            Principal   = [string]$row.'User Principal Name'
            UsedGB      = $usedGb
        }) | Out-Null
    }

    Write-Host "  [+] Mailbox report complete. $($report.Count) row(s)." -ForegroundColor Green
    return $report
}

function Get-ServiceSummary {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    return [PSCustomObject]@{
        Name        = $Name
        TotalItems  = $Rows.Count
        ActiveItems = @($Rows | Where-Object { $_.UsedGB -gt 0 }).Count
        TotalGB     = [math]::Round((($Rows | Measure-Object -Property UsedGB -Sum).Sum), 2)
        Rows        = @($Rows | Sort-Object -Property @(
            @{ Expression = "UsedGB"; Descending = $true },
            @{ Expression = "DisplayName"; Descending = $false }
        ))
    }
}

function Export-UsageHtmlReport {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$TenantName,
        [Parameter(Mandatory)]
        [object[]]$ServiceSummaries,
        [Parameter(Mandatory)]
        [string]$ReportDate
    )

    Write-Host "[*] Building HTML dashboard..." -ForegroundColor Cyan

    $overallItems = @($ServiceSummaries | Measure-Object -Property TotalItems -Sum).Sum
    $overallActive = @($ServiceSummaries | Measure-Object -Property ActiveItems -Sum).Sum
    $overallGb = [math]::Round((($ServiceSummaries | Measure-Object -Property TotalGB -Sum).Sum), 2)

    $htmlData = [PSCustomObject]@{
        tenant      = $TenantName
        reportDate  = $ReportDate
        totalItems  = $overallItems
        activeItems = $overallActive
        totalGB     = $overallGb
        services    = $ServiceSummaries
    } | ConvertTo-Json -Depth 6 -Compress

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M365 Usage Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root { --bg:#f5f7fb; --bg2:#ffffff; --bg3:#f7f9fc; --border:#d6dfec; --border2:#b8c8dc; --text:#17324d; --text2:#4b6887; --text3:#6f88a3; --accent:#0f7cc0; --accent2:#1091d5; --ok:#15803d; --warn:#b45309; --mono:'JetBrains Mono',monospace; --sans:'DM Sans',sans-serif; --r:8px; --r2:12px; }
  body { font-family:var(--sans); background:var(--bg); color:var(--text); min-height:100vh; line-height:1.6; }
  .topbar { position:sticky; top:0; z-index:20; background:rgba(245,247,251,.88); backdrop-filter:blur(16px); border-bottom:1px solid var(--border); padding:0 2rem; height:52px; display:flex; align-items:center; gap:1rem; }
  .topbar-logo,.topbar-server,.topbar-time,.strip-value,.card-badge,.summary-value,.kpi-value,th,td { font-family:var(--mono); }
  .topbar-logo { font-size:.72rem; font-weight:700; letter-spacing:.16em; text-transform:uppercase; color:var(--accent); }
  .topbar-server { color:var(--text2); font-size:.78rem; }
  .topbar-time { margin-left:auto; color:var(--text3); font-size:.7rem; }
  .page { max-width:1440px; margin:0 auto; padding:1.75rem 2rem 2rem; }
  .server-strip,.card,.kpi { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r2); }
  .server-strip { padding:1rem 1.5rem; display:flex; flex-wrap:wrap; gap:2rem; margin-bottom:1.25rem; }
  .strip-item { display:flex; flex-direction:column; gap:.2rem; }
  .strip-label,.kpi-label,.summary-label { font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); }
  .strip-value { font-size:.82rem; color:var(--accent); }
  .hero { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:.85rem; margin-bottom:1.25rem; }
  .kpi { padding:1rem 1.1rem; position:relative; overflow:hidden; }
  .kpi::after { content:''; position:absolute; top:0; left:0; right:0; height:2px; }
  .kpi.neutral::after { background:var(--accent2); }
  .kpi.ok::after { background:var(--ok); }
  .kpi.warn::after { background:var(--warn); }
  .kpi-value { font-size:1.6rem; font-weight:700; line-height:1; color:var(--accent); }
  .kpi-sub { font-size:.68rem; color:var(--text3); margin-top:.35rem; }
  .sections { display:grid; gap:1rem; }
  .card-header { display:flex; align-items:center; gap:.75rem; padding:.85rem 1.25rem; border-bottom:1px solid var(--border); }
  .card-title { font-size:.72rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text2); flex:1; }
  .card-badge { font-size:.7rem; padding:.18rem .55rem; border-radius:4px; background:rgba(15,124,192,.1); color:var(--accent); }
  .card-body { padding:1.25rem; }
  .summary-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:.85rem; }
  .summary-item { background:var(--bg3); border:1px solid var(--border); border-radius:var(--r); padding:.85rem 1rem; }
  .summary-value { font-size:1.05rem; font-weight:700; color:var(--accent); margin-top:.3rem; }
  .table-scroll { max-height:520px; overflow:auto; border:1px solid var(--border); border-radius:var(--r); margin-top:1rem; }
  table { width:100%; border-collapse:collapse; font-size:.77rem; }
  thead { background:var(--bg3); position:sticky; top:0; z-index:1; }
  th { padding:.55rem .9rem; text-align:left; font-size:.63rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--text3); border-bottom:1px solid var(--border); }
  td { padding:.5rem .9rem; border-bottom:1px solid var(--border); color:var(--text); font-size:.76rem; vertical-align:top; word-break:break-word; }
  tr:last-child td { border-bottom:none; }
  tbody tr:hover td { background:rgba(15,124,192,.05); }
  .empty { font-family:var(--mono); font-size:.78rem; color:var(--text3); font-style:italic; }
  a { color:var(--accent); }
  @media(max-width:700px){ .page{padding:1rem;} .topbar{padding:0 1rem;} .hero{grid-template-columns:repeat(2,1fr);} }
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-logo">M365 Usage Dashboard</div>
  <div class="topbar-server"></div>
  <div class="topbar-time"></div>
</div>
<div class="page">
  <div id="server-strip" class="server-strip"></div>
  <div id="hero" class="hero"></div>
  <div id="sections" class="sections"></div>
</div>
<script>
const DATA = $htmlData;
document.querySelector('.topbar-server').textContent = DATA.tenant;
document.querySelector('.topbar-time').textContent = 'Generated: ' + DATA.reportDate;
document.title = 'M365 Usage Report - ' + DATA.tenant;
function esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function kpi(value,label,sub,cls){return '<div class="kpi ' + cls + '"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>';}
function renderTable(rows){ if(!rows || rows.length===0){ return '<p class="empty">No data returned for this service.</p>'; } let html = '<div class="table-scroll"><table><thead><tr><th>Name</th><th>Principal</th><th>Used (GB)</th></tr></thead><tbody>'; for(const row of rows){ const name = row.Url ? '<a href="' + esc(row.Url) + '" target="_blank" rel="noreferrer">' + esc(row.DisplayName) + '</a>' : esc(row.DisplayName); html += '<tr><td>' + name + '</td><td>' + esc(row.Principal) + '</td><td>' + esc(row.UsedGB) + '</td></tr>'; } html += '</tbody></table></div>'; return html; }
document.getElementById('server-strip').innerHTML = [['Tenant', DATA.tenant],['Services', String(DATA.services.length)],['Items', String(DATA.totalItems)],['Active', String(DATA.activeItems)],['Total Used (GB)', String(DATA.totalGB)],['Report Generated', DATA.reportDate]].map(function(pair){ return '<div class="strip-item"><span class="strip-label">' + pair[0] + '</span><span class="strip-value">' + esc(pair[1]) + '</span></div>'; }).join('');
document.getElementById('hero').innerHTML = [kpi(DATA.services.length, 'Services', 'selected reports', 'neutral'),kpi(DATA.totalItems, 'Items', 'returned rows', 'neutral'),kpi(DATA.activeItems, 'Active', 'with usage > 0 GB', 'ok'),kpi(DATA.totalGB, 'Total GB', 'combined storage', 'warn')].join('');
const sections = document.getElementById('sections');
sections.innerHTML = DATA.services.map(function(service){ return '<div class="card"><div class="card-header"><span class="card-title">' + esc(service.Name) + ' Usage</span><span class="card-badge">' + esc(service.TotalItems) + ' item(s)</span></div><div class="card-body"><div class="summary-grid"><div class="summary-item"><div class="summary-label">Total Items</div><div class="summary-value">' + esc(service.TotalItems) + '</div></div><div class="summary-item"><div class="summary-label">Active Items</div><div class="summary-value">' + esc(service.ActiveItems) + '</div></div><div class="summary-item"><div class="summary-label">Total Used (GB)</div><div class="summary-value">' + esc(service.TotalGB) + '</div></div></div>' + renderTable(service.Rows) + '</div></div>'; }).join('');
</script>
</body>
</html>
"@

    $outDir = [System.IO.Path]::GetDirectoryName($Path)
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $html | Out-File -LiteralPath $Path -Encoding UTF8
    Write-Host "[+] HTML dashboard exported to: $Path" -ForegroundColor Green
}

function Show-FinalSummary {
    param(
        [Parameter(Mandatory)]
        [object[]]$ServiceSummaries,
        [string]$HtmlPath
    )

    Write-SectionHeader "EXECUTION COMPLETE"
    Write-Host "  [+] All selected reports have been generated." -ForegroundColor Green
    foreach ($summary in $ServiceSummaries) {
        Write-Host ("  [+] {0,-11}: {1,6} item(s) | {2,6} active | {3,8} GB" -f $summary.Name, $summary.TotalItems, $summary.ActiveItems, $summary.TotalGB) -ForegroundColor Green
    }
    if ($HtmlPath) {
        Write-Host "  [+] HTML dashboard : $HtmlPath" -ForegroundColor Green
    }
}

Assert-RequiredModules
Connect-ToGraph
Resolve-TenantLabel

$selectedReports = Normalize-ReportSelection -SelectedReports $Reports
$serviceSummaries = [System.Collections.Generic.List[object]]::new()

foreach ($reportName in $selectedReports) {
    switch ($reportName) {
        "OneDrive" {
            $serviceSummaries.Add((Get-ServiceSummary -Name "OneDrive" -Rows (Get-OneDriveReport))) | Out-Null
        }
        "SharePoint" {
            $serviceSummaries.Add((Get-ServiceSummary -Name "SharePoint" -Rows (Get-SharePointReport))) | Out-Null
        }
        "Mailbox" {
            $serviceSummaries.Add((Get-ServiceSummary -Name "Mailbox" -Rows (Get-MailboxReport))) | Out-Null
        }
    }
}

$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$htmlPath = Add-TimestampToPath -Path $(if ($ExportHtml) { $ExportHtml } else { $OutputPath }) -BaseName "M365_UsageReport" -Extension ".html"
Export-UsageHtmlReport -Path $htmlPath -TenantName $script:TenantLabel -ServiceSummaries $serviceSummaries -ReportDate $reportDate
Show-FinalSummary -ServiceSummaries $serviceSummaries -HtmlPath $htmlPath
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
