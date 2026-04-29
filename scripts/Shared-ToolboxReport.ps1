[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$script:ToolboxTenantLabel = ""

function Write-ToolboxStructuredEvent {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('progress', 'artifact', 'metric', 'state')]
        [string]$Type,

        [hashtable]$Body = @{}
    )

    $payload = [ordered]@{
        type      = $Type
        timestamp = (Get-Date).ToString('o')
    }

    foreach ($key in $Body.Keys) {
        $payload[$key] = $Body[$key]
    }

    $json = $payload | ConvertTo-Json -Depth 6 -Compress
    Write-Output ("::toolbox::{0}" -f $json)
}

function Write-ToolboxArtifactEvent {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [string]$Kind = 'file'
    )

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $item) {
        return
    }

    Write-ToolboxStructuredEvent -Type 'artifact' -Body @{
        path = $item.FullName
        kind = $Kind
        size = $item.Length
        name = $item.Name
    }
}

function Get-ToolboxExceptionMessage {
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception
    )

    $messages = [System.Collections.Generic.List[string]]::new()
    $current = $Exception

    while ($null -ne $current) {
        if ($current.Message -and -not $messages.Contains($current.Message)) {
            [void]$messages.Add($current.Message.Trim())
        }

        $current = $current.InnerException
    }

    return ($messages -join " | ")
}

function Write-SectionHeader {
    param([string]$Title)

    $line = "=" * 50
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host (" " * [Math]::Floor((50 - $Title.Length) / 2) + $Title) -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Write-ProgressStep {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-ToolboxStructuredEvent -Type 'progress' -Body @{ message = $Message }
    Write-Host "[+] $Message" -ForegroundColor Cyan
}

function Add-TimestampToPath {
    param(
        [string]$Path,
        [string]$BaseName = "Report",
        [string]$Extension = ".html",
        [string]$OutputPath
    )

    $stamp = Get-Date -Format "dd.MM.yy-HH.mm.ss"
    $tenant = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel.Split(".")[0] } else { "M365" }
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

function Assert-GraphModules {
    param(
        [string[]]$RequiredModules = @("Microsoft.Graph.Authentication")
    )

    Write-ProgressStep "Checking required PowerShell modules"
    Write-Host ""
    Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan

    $commonVersions = $null
    foreach ($moduleName in $RequiredModules) {
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
        throw "No common Microsoft Graph module version is installed across: $($RequiredModules -join ', ')."
    }

    $selectedVersion = $commonVersions |
        ForEach-Object { [version]$_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    foreach ($moduleName in $RequiredModules) {
        Import-Module $moduleName -RequiredVersion $selectedVersion -Force -WarningAction SilentlyContinue
        Write-Host "  [+] $moduleName v$selectedVersion" -ForegroundColor Green
    }
}

function Connect-ToolboxGraph {
    param(
        [string]$TenantId,
        [string[]]$Scopes
    )

    Write-SectionHeader "CONNECTING TO MICROSOFT GRAPH"
    Write-ProgressStep "Preparing Microsoft Graph device sign-in"

    if ($TenantId) {
        Write-Host "[*] Requested tenant: $TenantId"
    }
    else {
        Write-Host "[*] Requested tenant: auto-detect"
    }

    Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
    Write-Host "[*] When the code appears, open https://microsoft.com/devicelogin" -ForegroundColor Yellow

    Disconnect-MgGraph -ErrorAction SilentlyContinue

    $connectParams = @{
        Scopes       = $Scopes
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
        if ($context -and $context.TenantId) {
            $script:ToolboxTenantLabel = [string]$context.TenantId
        }
        Write-ProgressStep "Connected to Microsoft Graph"
        Write-Host "[+] Connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        $errorMessage = Get-ToolboxExceptionMessage -Exception $_.Exception
        throw "Failed to connect to Microsoft Graph: $errorMessage"
    }
}

function Resolve-ToolboxTenantLabel {
    Write-ProgressStep "Resolving tenant label"
    try {
        $domainResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains" -ErrorAction Stop
        $domains = @($domainResponse.value)
        $initial = $domains | Where-Object { $_.isInitial -eq $true } | Select-Object -First 1
        $default = $domains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1

        if ($default.id) {
            $script:ToolboxTenantLabel = [string]$default.id
        }
        elseif ($initial.id) {
            $script:ToolboxTenantLabel = [string]$initial.id
        }
    }
    catch {
        if (-not $script:ToolboxTenantLabel) {
            $context = Get-MgContext
            if ($context -and $context.TenantId) {
                $script:ToolboxTenantLabel = [string]$context.TenantId
            }
        }
    }
}

function Invoke-GraphCollection {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    Write-ProgressStep "Requesting Microsoft Graph collection data"
    $items = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri

    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop

        if ($response.value) {
            foreach ($item in @($response.value)) {
                [void]$items.Add($item)
            }
        }
        elseif ($response) {
            [void]$items.Add($response)
        }

        $odataNextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        $nextLink = if ($odataNextLinkProperty) { [string]$odataNextLinkProperty.Value } else { $null }
    }

    return @($items)
}

function Import-GraphCsvReport {
    param(
        [Parameter(Mandatory)]
        [string]$RequestUri
    )

    Write-ProgressStep "Downloading Microsoft Graph CSV report"
    $tempFile = [System.IO.Path]::GetTempFileName()

    try {
        if ($RequestUri -match "/reports/getSharePointSiteUsageDetail\(period='([^']+)'\)") {
            Get-MgReportSharePointSiteUsageDetail -Period $matches[1] -OutFile $tempFile -ErrorAction Stop | Out-Null
        }
        elseif ($RequestUri -match "/reports/getSharePointSiteUsageDetail\(date=([0-9-]+)\)") {
            Get-MgReportSharePointSiteUsageDetail -Date ([datetime]$matches[1]) -OutFile $tempFile -ErrorAction Stop | Out-Null
        }
        elseif ($RequestUri -match "/reports/getOneDriveUsageAccountDetail\(period='([^']+)'\)") {
            Get-MgReportOneDriveUsageAccountDetail -Period $matches[1] -OutFile $tempFile -ErrorAction Stop | Out-Null
        }
        elseif ($RequestUri -match "/reports/getOneDriveUsageAccountDetail\(date=([0-9-]+)\)") {
            Get-MgReportOneDriveUsageAccountDetail -Date ([datetime]$matches[1]) -OutFile $tempFile -ErrorAction Stop | Out-Null
        }
        elseif ($RequestUri -match "/reports/getMailboxUsageDetail\(period='([^']+)'\)") {
            Get-MgReportMailboxUsageDetail -Period $matches[1] -OutFile $tempFile -ErrorAction Stop | Out-Null
        }
        else {
            Invoke-MgGraphRequest -Method GET -Uri $RequestUri -OutputFilePath $tempFile -ErrorAction Stop | Out-Null
        }
        $rawContent = [System.IO.File]::ReadAllText($tempFile, [System.Text.Encoding]::UTF8)
        $rawContent = $rawContent -replace "^[^R]*Report Refresh Date", "Report Refresh Date"
        return @($rawContent | ConvertFrom-Csv)
    }
    finally {
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }
}

function Normalize-DelimitedValue {
    param([string[]]$Value)

    return @(
        $Value |
            Where-Object { $_ } |
            ForEach-Object { $_ -split "[,\r\n]+" } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
    )
}

function Get-DirectoryObjectValue {
    param(
        $DirectoryObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $DirectoryObject) {
        return $null
    }

    $property = $DirectoryObject.PSObject.Properties |
        Where-Object { $_.Name -ieq $Name } |
        Select-Object -First 1
    if ($property) {
        return $property.Value
    }

    if ($DirectoryObject -is [System.Collections.IDictionary]) {
        foreach ($key in $DirectoryObject.Keys) {
            if ([string]$key -ieq $Name) {
                return $DirectoryObject[$key]
            }
        }
    }

    $additionalProperties = $DirectoryObject.PSObject.Properties |
        Where-Object { $_.Name -eq 'AdditionalProperties' } |
        Select-Object -First 1
    if ($additionalProperties -and $additionalProperties.Value -is [System.Collections.IDictionary]) {
        foreach ($key in $additionalProperties.Value.Keys) {
            if ([string]$key -ieq $Name) {
                return $additionalProperties.Value[$key]
            }
        }
    }

    return $null
}

function Get-DirectoryObjectLabel {
    param($DirectoryObject)

    if (-not $DirectoryObject) {
        return ""
    }

    $displayName = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'displayName')
    $userPrincipalName = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'userPrincipalName')

    if ($userPrincipalName) {
        return "{0} ({1})" -f $displayName, $userPrincipalName
    }

    $mail = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'mail')
    if ($mail) {
        return "{0} ({1})" -f $displayName, $mail
    }

    if ($displayName) {
        return $displayName
    }

    $id = [string](Get-DirectoryObjectValue -DirectoryObject $DirectoryObject -Name 'id')
    if ($id) {
        return $id
    }

    return ""
}

function Get-GraphAuthMethodLabel {
    param($Method)

    if (-not $Method) {
        return "Unknown"
    }

    $typeHints = [System.Collections.Generic.List[string]]::new()
    foreach ($typeHint in @(
        [string](Get-DirectoryObjectValue -DirectoryObject $Method -Name '@odata.type'),
        [string](Get-DirectoryObjectValue -DirectoryObject $Method -Name 'OdataType')
    )) {
        if ($typeHint) {
            [void]$typeHints.Add($typeHint)
        }
    }

    foreach ($typeName in @($Method.PSObject.TypeNames)) {
        if ($typeName) {
            [void]$typeHints.Add([string]$typeName)
        }
    }

    try {
        $fullTypeName = [string]$Method.GetType().FullName
        if ($fullTypeName) {
            [void]$typeHints.Add($fullTypeName)
        }
    }
    catch {}

    $typeText = $typeHints -join ' '
    switch -Regex ($typeText) {
        'microsoftAuthenticatorAuthenticationMethod' { return 'Authenticator App' }
        'phoneAuthenticationMethod' { return 'Phone' }
        'fido2AuthenticationMethod' { return 'FIDO2 Key' }
        'softwareOathAuthenticationMethod' { return 'Software OATH' }
        'emailAuthenticationMethod' { return 'Email OTP' }
        'windowsHelloForBusinessAuthenticationMethod' { return 'Windows Hello' }
        'passwordAuthenticationMethod' { return 'Password' }
    }

    if ($typeHints.Count -gt 0) {
        return (($typeHints[0] -replace '^#microsoft\.graph\.', '') -replace '^Microsoft\.Graph\.PowerShell\.Models\.', '')
    }

    return "Unknown"
}

function Get-GeoLocationString {
    param($Location)

    if (-not $Location) {
        return ''
    }

    $parts = @()
    foreach ($name in @('City', 'State', 'CountryOrRegion')) {
        if ($Location.PSObject.Properties.Name -contains $name -and $Location.$name) {
            $parts += [string]$Location.$name
        }
    }

    return ($parts -join ', ')
}

function ConvertTo-ReportJson {
    param([object]$Value)

    return $Value | ConvertTo-Json -Depth 10 -Compress
}

function Export-ToolboxHtmlReport {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Title,

        [string]$Tenant = "",
        [string]$Subtitle = "",
        [array]$Kpis = @(),
        [array]$StripItems = @(),
        [array]$Sections = @()
    )

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if ($directory) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    Write-ProgressStep "Exporting HTML dashboard"

    $payload = @{
        title      = $Title
        tenant     = $Tenant
        subtitle   = $Subtitle
        reportDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        kpis       = @($Kpis)
        stripItems = @($StripItems)
        sections   = @($Sections)
    }

    $json = ConvertTo-ReportJson -Value $payload

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>$Title</title>
  <style>
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    :root{--bg:#f5f7fb;--bg2:#ffffff;--bg3:#f7f9fc;--bg4:#eef3f9;--border:#d6dfec;--border2:#b8c8dc;--text:#17324d;--text2:#4b6887;--text3:#6f88a3;--accent:#0f7cc0;--accent2:#1091d5;--ok:#15803d;--warn:#b45309;--crit:#dc2626;--mono:'JetBrains Mono',monospace;--sans:'DM Sans',sans-serif;--r:8px;--r2:12px}
    body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.6}
    body::before{content:'';position:fixed;inset:0;background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='0.025'/%3E%3C/svg%3E");pointer-events:none;z-index:0;opacity:.35}
    .topbar{position:sticky;top:0;z-index:20;background:rgba(245,247,251,.88);backdrop-filter:blur(16px);border-bottom:1px solid var(--border);padding:0 2rem;height:52px;display:flex;align-items:center;gap:1rem}
    .topbar-logo,.topbar-server,.topbar-time,.strip-value,.kpi-value,.pill,th,td{font-family:var(--mono)}
    .topbar-logo{font-size:.72rem;font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:var(--accent)}
    .topbar-server{color:var(--text2);font-size:.78rem}
    .topbar-time{margin-left:auto;color:var(--text3);font-size:.7rem}
    .page{max-width:1440px;margin:0 auto;padding:1.75rem 2rem 2rem;position:relative;z-index:1}
    .card,.kpi,.strip-item{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2)}
    .server-strip{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.85rem;margin-bottom:1.25rem}
    .strip-item{min-width:0;min-height:74px;padding:1rem 1.1rem;display:flex;flex-direction:column;gap:.2rem}
    .strip-label,.kpi-label{font-size:.63rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text3)}
    .strip-value{font-size:.82rem;color:var(--accent)}
    .hero{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.85rem;margin-bottom:1.25rem}
    .kpi{min-width:0;min-height:98px;padding:1rem 1.1rem;position:relative;overflow:hidden}
    .kpi::after{content:'';position:absolute;top:0;left:0;right:0;height:2px}
    .kpi.neutral::after{background:var(--accent)}
    .kpi.ok::after{background:var(--ok)}
    .kpi.warn::after{background:var(--warn)}
    .kpi.crit::after{background:var(--crit)}
    .kpi-value{font-size:1.6rem;font-weight:700;line-height:1;color:var(--accent);margin-top:.25rem}
    .kpi.ok .kpi-value{color:var(--ok)}
    .kpi.warn .kpi-value{color:var(--warn)}
    .kpi.crit .kpi-value{color:var(--crit)}
    .kpi-sub{font-size:.68rem;color:var(--text3);margin-top:.35rem}
    .sections{display:grid;gap:1rem}
    .card-header{display:flex;align-items:center;gap:.75rem;padding:.85rem 1.25rem;border-bottom:1px solid var(--border)}
    .card-title{font-size:.72rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text2);flex:1}
    .card-badge{font-size:.7rem;padding:.18rem .55rem;border-radius:4px;background:rgba(15,124,192,.1);color:var(--accent)}
    .card-body{padding:1.25rem}
    .card-body.table-card-body{padding:0}
    .card-text{font-family:var(--mono);font-size:.78rem;color:var(--text2)}
    .table-card-shell{padding:1.1rem 1.25rem 1.25rem}
    .table-filter{display:grid;gap:.35rem;margin-bottom:.8rem}
    .table-filter-label{font-size:.63rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--text3)}
    .table-filter-input{width:100%;max-width:320px;border:1px solid var(--border);border-radius:var(--r);background:var(--bg3);color:var(--text);font-family:var(--mono);font-size:.74rem;padding:.55rem .7rem;outline:none}
    .table-filter-input:focus{border-color:var(--accent2)}
    .pill{display:inline-block;padding:.12rem .45rem;border-radius:4px;font-size:.7rem}
    .badge-stack{display:flex;flex-wrap:wrap;gap:.35rem}
    .ok{background:rgba(21,128,61,.1);color:var(--ok)}
    .warn{background:rgba(180,83,9,.1);color:var(--warn)}
    .crit{background:rgba(220,38,38,.1);color:var(--crit)}
    .neutral{background:rgba(15,124,192,.1);color:var(--accent)}
    .table-scroll{max-height:620px;overflow:auto;border-top:1px solid var(--border);border-bottom-left-radius:var(--r2);border-bottom-right-radius:var(--r2)}
    table{width:100%;border-collapse:collapse;font-size:.77rem;table-layout:fixed}
    thead{background:var(--bg3);position:sticky;top:0;z-index:1}
    th{position:relative;padding:.55rem .9rem;text-align:left;font-size:.63rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);border-bottom:1px solid var(--border);white-space:nowrap}
    td{padding:.5rem .9rem;border-bottom:1px solid var(--border);color:var(--text);font-size:.76rem;vertical-align:top;word-break:break-word}
    tr:last-child td{border-bottom:none}
    tbody tr:hover td{background:rgba(15,124,192,.05)}
    tbody tr.table-row-action{cursor:pointer}
    tbody tr.table-row-action td:first-child{text-decoration:underline;text-underline-offset:.18em}
    .empty{font-family:var(--mono);font-size:.78rem;color:var(--text3);font-style:italic}
    .card-text{white-space:pre-wrap;word-break:break-word}
    a{color:var(--accent)}
    .col-resizer{position:absolute;top:0;right:0;width:10px;height:100%;cursor:col-resize;user-select:none;touch-action:none}
    .col-resizer::after{content:'';position:absolute;top:20%;bottom:20%;right:4px;width:2px;border-radius:999px;background:transparent;transition:background .15s ease}
    th:hover .col-resizer::after,.col-resizer.active::after{background:var(--accent2)}
    @media(max-width:980px){.server-strip,.hero{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @media(max-width:700px){.page{padding:1rem}.topbar{padding:0 1rem}.server-strip,.hero{grid-template-columns:1fr}}
  </style>
</head>
<body>
  <div class="topbar">
    <div class="topbar-logo">$Title</div>
    <div class="topbar-server"></div>
    <div class="topbar-time"></div>
  </div>
  <div class="page">
    <div id="server-strip" class="server-strip"></div>
    <div id="hero" class="hero"></div>
    <div id="sections" class="sections"></div>
  </div>
  <script>
  const DATA = $json;
  document.querySelector('.topbar-server').textContent = DATA.tenant || DATA.subtitle || '';
  document.querySelector('.topbar-time').textContent = 'Generated: ' + DATA.reportDate;
  document.title = DATA.title + (DATA.tenant ? ' - ' + DATA.tenant : '');
  function esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function attr(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/"/g,'&quot;');}
  function pillClass(v){
    const text=String(v==null?'':v).toLowerCase();
    if(text.includes('success')||text.includes('ok')||text.includes('enabled')||text.includes('registered')||text==='yes'){return 'ok';}
    if(text.includes('warning')||text.includes('review')||text.includes('report-only')||text.includes('external')||text.includes('preview')){return 'warn';}
    if(text.includes('fail')||text.includes('no mfa')||text.includes('not registered')||text.includes('high')||text.includes('disabled')){return 'crit';}
    return 'neutral';
  }
  function renderStrip(items){
    if(!items||items.length===0){return '<div class="strip-item"><span class="strip-label">Tenant</span><span class="strip-value">'+esc(DATA.tenant||'N/A')+'</span></div>';}
    return items.map(function(item){return '<div class="strip-item"><span class="strip-label">'+esc(item.label)+'</span><span class="strip-value">'+esc(item.value)+'</span></div>';}).join('');
  }
  function renderKpis(items){
    return (items||[]).map(function(item){const cls=item.cls||'neutral';return '<div class="kpi '+cls+'"><div class="kpi-label">'+esc(item.label)+'</div><div class="kpi-value">'+esc(item.value)+'</div><div class="kpi-sub">'+esc(item.sub||'')+'</div></div>';}).join('');
  }
  function renderRows(columns,rows,section){
    if(!rows||rows.length===0){return '<div class="empty">No records returned.</div>';}
    let html='';
    if(section && section.filterKey){
      html+='<div class="table-filter"><label class="table-filter-label" for="table-filter-'+attr(section.id||section.title||'section')+'">'+esc(section.filterLabel||'Filter')+'</label><input id="table-filter-'+attr(section.id||section.title||'section')+'" class="table-filter-input" type="search" placeholder="'+attr(section.filterPlaceholder||'Type to filter rows')+'" data-filter-key="'+attr(section.filterKey)+'" autocomplete="off" /></div>';
    }
    html+='<div class="table-scroll"><table><colgroup>';
    for(let index=0;index<columns.length;index++){html+='<col />';}
    html+='</colgroup><thead><tr>';
    for(const col of columns){html+='<th>'+esc(col.header||col.key)+'</th>';}
    html+='</tr></thead><tbody>';
    for(const row of rows){
      const filterValue=section&&section.filterKey?String(row[section.filterKey]==null?'':row[section.filterKey]).toLowerCase():'';
      const actionValue=section&&section.rowAction?String(row[section.rowAction.sourceKey]==null?'':row[section.rowAction.sourceKey]):'';
      const rowClass=section&&section.rowAction?' class="table-row-action"':'';
      const filterAttr=section&&section.filterKey?' data-filter-value="'+attr(filterValue)+'"':'';
      const actionAttrs=section&&section.rowAction
        ? ' data-action-target="'+attr(section.rowAction.targetSectionId||'')+'" data-action-filter="'+attr(actionValue)+'" tabindex="0" role="button" aria-label="'+attr(section.rowAction.ariaLabel||'Filter details for this row')+'"'
        : '';
      html+='<tr'+rowClass+filterAttr+actionAttrs+'>';
      for(const col of columns){
        const raw=row[col.key];
        if(col.type==='pill'){html+='<td><span class="pill '+pillClass(raw)+'">'+esc(raw||'—')+'</span></td>';continue;}
        if(col.type==='link' && raw){html+='<td><a href="'+esc(raw)+'" target="_blank" rel="noreferrer">'+esc(raw)+'</a></td>';continue;}
        if(col.type==='multiline'){
          const lines=String(raw==null?'':raw).split(/\r?\n/).filter(Boolean);
          html+='<td>'+(lines.length?lines.map(function(line){return esc(line);}).join('<br>'):'—')+'</td>';
          continue;
        }
        if(col.type==='badgelist'){
          const lines=String(raw==null?'':raw).split(/\r?\n/).filter(Boolean);
          html+='<td>'+(lines.length?'<div class="badge-stack">'+lines.map(function(line){return '<span class="pill neutral">'+esc(line)+'</span>';}).join('')+'</div>':'—')+'</td>';
          continue;
        }
        html+='<td>'+esc(raw||'—')+'</td>';
      }
      html+='</tr>';
    }
    html+='</tbody></table></div>';
    return html;
  }
  function enableResizableColumns(){
    document.querySelectorAll('.table-scroll table').forEach(function(table){
      const headers=table.querySelectorAll('thead th');
      const columns=table.querySelectorAll('colgroup col');
      if(!headers.length){return;}
      table.style.width='max-content';
      table.style.minWidth='100%';
      headers.forEach(function(header,index){
        const column=columns[index];
        if(column && !column.style.width){
          column.style.width=Math.max(90,header.getBoundingClientRect().width) + 'px';
        }
      });
      headers.forEach(function(header,index){
        if(header.querySelector('.col-resizer')){return;}
        const resizer=document.createElement('span');
        resizer.className='col-resizer';
        header.appendChild(resizer);
        let startX=0;
        let startWidth=0;
        const column=columns[index];
        const onMove=function(event){
          const nextWidth=Math.max(90,startWidth + (event.clientX - startX));
          if(column){
            column.style.width=nextWidth + 'px';
          }
          header.style.width=nextWidth + 'px';
        };
        const onUp=function(){
          resizer.classList.remove('active');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
        };
        resizer.addEventListener('mousedown', function(event){
          event.preventDefault();
          startX=event.clientX;
          startWidth=column ? column.getBoundingClientRect().width : header.getBoundingClientRect().width;
          resizer.classList.add('active');
          window.addEventListener('mousemove', onMove);
          window.addEventListener('mouseup', onUp);
        });
      });
    });
  }
  function enableTableFilters(){
    document.querySelectorAll('.table-filter-input').forEach(function(input){
      const tableScroll=input.closest('.card-body').querySelector('.table-scroll');
      const rows=tableScroll?tableScroll.querySelectorAll('tbody tr'):[];
      const applyFilter=function(){
        const query=input.value.trim().toLowerCase();
        rows.forEach(function(row){
          const haystack=row.getAttribute('data-filter-value')||'';
          row.style.display=!query||haystack.includes(query)?'':'none';
        });
      };
      input.addEventListener('input', applyFilter);
      applyFilter();
    });
  }
  function setSectionFilter(targetSectionId, value){
    const input=document.getElementById('table-filter-'+targetSectionId);
    if(!input){return;}
    input.value=value||'';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus({ preventScroll: true });
    const targetCard=input.closest('.card');
    if(targetCard){targetCard.scrollIntoView({ behavior:'smooth', block:'start' });}
  }
  function enableRowActions(){
    document.querySelectorAll('tr[data-action-target]').forEach(function(row){
      const activate=function(){
        setSectionFilter(row.getAttribute('data-action-target'), row.getAttribute('data-action-filter')||'');
      };
      row.addEventListener('click', activate);
      row.addEventListener('keydown', function(event){
        if(event.key==='Enter' || event.key===' '){
          event.preventDefault();
          activate();
        }
      });
    });
  }
  function renderSections(items){
    return (items||[]).map(function(section){
      let inner='';
      let bodyClass='card-body';
      if(section.text){inner='<div class="card-text">'+esc(section.text)+'</div>';}
      if(section.columns){inner='<div class="table-card-shell">'+renderRows(section.columns,section.rows||[],section)+'</div>';bodyClass+=' table-card-body';}
      return '<div class="card"'+(section.id?' id="'+attr(section.id)+'"':'')+'><div class="card-header"><span class="card-title">'+esc(section.title)+'</span>'+(section.badge?'<span class="card-badge">'+esc(section.badge)+'</span>':'')+'</div><div class="'+bodyClass+'">'+inner+'</div></div>';
    }).join('');
  }
  document.getElementById('server-strip').innerHTML=renderStrip(DATA.stripItems);
  document.getElementById('hero').innerHTML=renderKpis(DATA.kpis);
  document.getElementById('sections').innerHTML=renderSections(DATA.sections);
  enableTableFilters();
  enableRowActions();
  enableResizableColumns();
  </script>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($Path, $html, [System.Text.Encoding]::UTF8)
    Write-ToolboxArtifactEvent -Path $Path -Kind 'html'
}
