[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,

    [Parameter(ValueFromRemainingArguments = $true)]
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

function Convert-ToolboxOutputRecord {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $Record
    )

    process {
        if ($null -eq $Record) {
            return
        }

        if ($Record -is [System.Management.Automation.InformationRecord]) {
            $messageData = $Record.MessageData
            if ($null -ne $messageData -and $messageData.PSObject.Properties.Match("Message").Count -gt 0) {
                Write-Output ([string]$messageData.Message)
                return
            }

            Write-Output ([string]$messageData)
            return
        }

        if ($Record -is [System.Management.Automation.WarningRecord]) {
            Write-Output ([string]$Record.Message)
            return
        }

        if ($Record -is [System.Management.Automation.VerboseRecord]) {
            Write-Output ([string]$Record.Message)
            return
        }

        if ($Record -is [System.Management.Automation.DebugRecord]) {
            Write-Output ([string]$Record.Message)
            return
        }

        if ($Record.PSObject.Properties.Match("Message").Count -gt 0 -and
            $Record.PSObject.Properties.Match("ForegroundColor").Count -gt 0) {
            Write-Output ([string]$Record.Message)
            return
        }

        if ($Record.PSObject.Properties.Match("MessageData").Count -gt 0) {
            $messageData = $Record.MessageData
            if ($null -ne $messageData -and $messageData.PSObject.Properties.Match("Message").Count -gt 0) {
                Write-Output ([string]$messageData.Message)
                return
            }

            Write-Output ([string]$messageData)
            return
        }

        Write-Output $Record
    }
}

function Split-ScriptInvocationArguments {
    param(
        [string[]]$ArgumentList
    )

    $namedArguments = @{}
    $positionalArguments = [System.Collections.Generic.List[string]]::new()

    for ($index = 0; $index -lt $ArgumentList.Count; $index++) {
        $argument = $ArgumentList[$index]

        if ($argument -match '^-{1,2}([^=]+)$') {
            $name = $Matches[1]

            if (($index + 1) -lt $ArgumentList.Count -and $ArgumentList[$index + 1] -notmatch '^-{1,2}[^\\/].*') {
                $namedArguments[$name] = $ArgumentList[$index + 1]
                $index++
            }
            else {
                $namedArguments[$name] = $true
            }

            continue
        }

        $positionalArguments.Add($argument)
    }

    return [PSCustomObject]@{
        Named      = $namedArguments
        Positional = @($positionalArguments)
    }
}

try {
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script file not found: $ScriptPath"
    }

    $requiredModules = Get-RequiredModuleNames -Path $ScriptPath
    Install-MissingModules -ModuleNames $requiredModules

    $InformationPreference = "Continue"
    $invocationArguments = Split-ScriptInvocationArguments -ArgumentList $ScriptArgumentList
    $namedArguments = $invocationArguments.Named
    $positionalArguments = $invocationArguments.Positional

    & $ScriptPath @namedArguments @positionalArguments

    exit 0
}
catch {
    Write-Error $_
    exit 1
}
