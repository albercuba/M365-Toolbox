[CmdletBinding()]
param(
    [string]$TenantId,

    [ValidateSet("OneDrive", "SharePoint", "Mailbox")]
    [string[]]$Reports = @("OneDrive", "SharePoint", "Mailbox"),

    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$script:TenantLabel = ""

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
        [string]$Extension = ".csv"
    )

    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $tenant = if ($script:TenantLabel) { $script:TenantLabel.Split(".")[0] } else { "M365" }
    $fileName = "{0}_{1}_{2}{3}" -f $BaseName, $tenant, $stamp, $Extension

    if (-not $Path) {
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
            $ctx = Get-MgContext
            if ($ctx -and $ctx.TenantId) {
                $script:TenantLabel = [string]$ctx.TenantId
            }
        }
    }
}

function ConvertTo-GB {
    param([double]$Bytes)
    [math]::Round($Bytes / 1GB, 2)
}

function Import-ReportCsv {
    param(
        [scriptblock]$DownloadAction
    )

    $tempFile = [System.IO.Path]::GetTempFileName()

    try {
        & $DownloadAction $tempFile
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

    $rows = Import-ReportCsv -DownloadAction {
        param($filePath)
        Get-MgReportOneDriveUsageAccountDetail -Period D7 -OutFile $filePath -ErrorAction Stop | Out-Null
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  [!] No OneDrive usage rows were returned."
        return [System.Collections.Generic.List[object]]::new()
    }

    $report = [System.Collections.Generic.List[object]]::new()
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
            'Owner Display Name'   = [string]$row.'Owner Display Name'
            'Owner Principal Name' = [string]$row.'Owner Principal Name'
            'Used (GB)'            = $usedGb
        }) | Out-Null
    }

    Write-Host "  [+] OneDrive report complete. $($report.Count) row(s)." -ForegroundColor Green
    return $report
}

function Get-SharePointReport {
    Write-SectionHeader "SHAREPOINT USAGE REPORT"
    Write-Host "[*] Fetching SharePoint usage data from the Reports API..." -ForegroundColor Cyan

    $rows = Import-ReportCsv -DownloadAction {
        param($filePath)
        Get-MgReportSharePointSiteUsageDetail -Period D7 -OutFile $filePath -ErrorAction Stop | Out-Null
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  [!] No SharePoint usage rows were returned."
        return [System.Collections.Generic.List[object]]::new()
    }

    $report = [System.Collections.Generic.List[object]]::new()
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
            'Site Id'              = [string]$row.'Site Id'
            'Site URL'             = $siteUrl
            'Owner Principal Name' = [string]$row.'Owner Principal Name'
            'Owner Display Name'   = [string]$row.'Owner Display Name'
            'Used (GB)'            = $usedGb
        }) | Out-Null
    }

    Write-Host "  [+] SharePoint report complete. $($report.Count) row(s)." -ForegroundColor Green
    return $report
}

function Get-MailboxReport {
    Write-SectionHeader "MAILBOX USAGE REPORT"
    Write-Host "[*] Fetching mailbox usage data from the Reports API..." -ForegroundColor Cyan

    $rows = Import-ReportCsv -DownloadAction {
        param($filePath)
        Get-MgReportMailboxUsageDetail -Period D7 -OutFile $filePath -ErrorAction Stop | Out-Null
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  [!] No mailbox usage rows were returned."
        return [System.Collections.Generic.List[object]]::new()
    }

    $report = [System.Collections.Generic.List[object]]::new()
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
            'User Principal Name' = [string]$row.'User Principal Name'
            'Display Name'        = [string]$row.'Display Name'
            'Used (GB)'           = $usedGb
        }) | Out-Null
    }

    Write-Host "  [+] Mailbox report complete. $($report.Count) row(s)." -ForegroundColor Green
    return $report
}

function Export-Report {
    param(
        [Parameter(Mandatory)]
        [object[]]$Data,
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string]$Label
    )

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Warning "  [!] No data to export for $Label."
        return
    }

    $totalGb = [math]::Round((($Data | Measure-Object -Property 'Used (GB)' -Sum).Sum), 2)
    $firstRow = $Data[0]
    $totalRow = [ordered]@{}

    foreach ($propertyName in $firstRow.PSObject.Properties.Name) {
        if ($propertyName -eq 'Used (GB)') {
            $totalRow[$propertyName] = $totalGb
        }
        elseif ($propertyName -eq ($firstRow.PSObject.Properties.Name | Select-Object -First 1)) {
            $totalRow[$propertyName] = 'TOTAL'
        }
        else {
            $totalRow[$propertyName] = ''
        }
    }

    $outDir = [System.IO.Path]::GetDirectoryName($FilePath)
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
    [PSCustomObject]$totalRow | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Append

    Write-Host "  [+] $Label report saved to: $FilePath" -ForegroundColor Green
    Write-Host "  [+] Total used space ($Label): $totalGb GB" -ForegroundColor Green
}

function Show-FinalSummary {
    param([hashtable]$Exports)

    Write-SectionHeader "EXECUTION COMPLETE"
    Write-Host "  [+] All selected reports have been generated." -ForegroundColor Green
    foreach ($key in $Exports.Keys) {
        Write-Host ("  [+] {0,-11}: {1}" -f $key, $Exports[$key]) -ForegroundColor Green
    }
}

Assert-RequiredModules
Connect-ToGraph
Resolve-TenantLabel

$selectedReports = @($Reports | Select-Object -Unique)
$exports = @{}

foreach ($reportName in $selectedReports) {
    switch ($reportName) {
        "OneDrive" {
            $oneDriveData = Get-OneDriveReport
            $oneDrivePath = Add-TimestampToPath -Path $OutputPath -BaseName "OneDrive_UsageReport"
            Export-Report -Data $oneDriveData -FilePath $oneDrivePath -Label "OneDrive"
            $exports["OneDrive"] = $oneDrivePath
        }
        "SharePoint" {
            $sharePointData = Get-SharePointReport
            $sharePointPath = Add-TimestampToPath -Path $OutputPath -BaseName "SharePoint_UsageReport"
            Export-Report -Data $sharePointData -FilePath $sharePointPath -Label "SharePoint"
            $exports["SharePoint"] = $sharePointPath
        }
        "Mailbox" {
            $mailboxData = Get-MailboxReport
            $mailboxPath = Add-TimestampToPath -Path $OutputPath -BaseName "Mailbox_UsageReport"
            Export-Report -Data $mailboxData -FilePath $mailboxPath -Label "Mailbox"
            $exports["Mailbox"] = $mailboxPath
        }
    }
}

Show-FinalSummary -Exports $exports
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
