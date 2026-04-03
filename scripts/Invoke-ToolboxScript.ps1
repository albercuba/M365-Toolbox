[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,

    [string[]]$ScriptArgumentList = @()
)

function Get-RequiredModuleNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $moduleNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $requireMatches = [regex]::Matches($content, '(?im)^\s*#requires\s+-modules\s+(.+)$')

    foreach ($match in $requireMatches) {
        $spec = $match.Groups[1].Value

        foreach ($moduleSpec in ($spec -split ',')) {
            $trimmed = $moduleSpec.Trim()
            if (-not $trimmed) {
                continue
            }

            if ($trimmed -match "ModuleName\s*=\s*['""]([^'""]+)['""]") {
                [void]$moduleNames.Add($Matches[1])
                continue
            }

            if ($trimmed -match "^['""]([^'""]+)['""]$") {
                [void]$moduleNames.Add($Matches[1])
                continue
            }

            if ($trimmed -match '^[A-Za-z0-9_.-]+$') {
                [void]$moduleNames.Add($trimmed)
            }
        }
    }

    return @($moduleNames)
}

function Install-MissingModules {
    param(
        [string[]]$ModuleNames
    )

    if (-not $ModuleNames -or $ModuleNames.Count -eq 0) {
        return
    }

    $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repository -and $repository.InstallationPolicy -ne "Trusted") {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    foreach ($moduleName in $ModuleNames) {
        if (Get-Module -ListAvailable -Name $moduleName) {
            continue
        }

        Write-Host "[*] Installing missing module '$moduleName'..." -ForegroundColor Yellow
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    }
}

try {
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script file not found: $ScriptPath"
    }

    $requiredModules = Get-RequiredModuleNames -Path $ScriptPath
    Install-MissingModules -ModuleNames $requiredModules

    $InformationPreference = "Continue"

    & $ScriptPath @ScriptArgumentList 3>&1 4>&1 5>&1 6>&1 | ForEach-Object {
        if ($null -ne $_ -and $_.PSObject.Properties.Match('Message').Count -gt 0) {
            Write-Output ([string]$_.Message)
        }
        elseif ($null -ne $_ -and $_.PSObject.Properties.Match('MessageData').Count -gt 0) {
            Write-Output ([string]$_.MessageData)
        }
        if ($_ -is [System.Management.Automation.InformationRecord]) {
            Write-Output $_.MessageData
        }
        elseif ($_ -is [System.Management.Automation.WarningRecord]) {
            Write-Output $_.Message
        }
        elseif ($_ -is [System.Management.Automation.VerboseRecord]) {
            Write-Output $_.Message
        }
        elseif ($_ -is [System.Management.Automation.DebugRecord]) {
            Write-Output $_.Message
        }
        else {
            Write-Output $_
        }
    }

    if ($LASTEXITCODE) {
        exit $LASTEXITCODE
    }

    exit 0
}
catch {
    Write-Error $_
    exit 1
}
