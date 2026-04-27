[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

Write-Host ""
Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
Write-Host "  [+] ExchangeOnlineManagement" -ForegroundColor Green

Write-SectionHeader "CONNECTING TO EXCHANGE ONLINE"
if ($TenantId) { Write-Host "[*] Requested tenant: $TenantId" } else { Write-Host "[*] Requested tenant: auto-detect" }
Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
$exoParams = @{ ShowBanner = $false; ShowProgress = $false; Device = $true }
if ($TenantId) { $exoParams.Organization = $TenantId }
Connect-ExchangeOnline @exoParams | Out-Null

try {
    $script:ToolboxTenantLabel = if ($TenantId) { $TenantId } else { "ExchangeOnline" }
    Write-Host "[+] Connected to Exchange Online" -ForegroundColor Green
    Write-SectionHeader "COLLECTING DKIM AND DMARC DATA"

    $domains = @(Get-AcceptedDomain)
    $rows = foreach ($domain in $domains) {
        $dkim = $null
        try { $dkim = Get-DkimSigningConfig -Identity $domain.DomainName -ErrorAction Stop } catch {}
        $dkimEnabled = Get-OptionalPropertyValue -InputObject $dkim -PropertyName "Enabled"

        $dmarcRecord = "Not queried"
        try {
            $lookup = Resolve-DnsName -Name ("_dmarc.{0}" -f $domain.DomainName) -Type TXT -ErrorAction Stop
            $dmarcRecord = (($lookup | Select-Object -ExpandProperty Strings) -join " ")
        }
        catch {
        }

        [pscustomobject]@{
            Domain     = [string]$domain.DomainName
            DomainType = [string]$domain.DomainType
            Default    = if ($domain.Default) { "Yes" } else { "No" }
            Dkim       = if ($null -eq $dkimEnabled) { "Unknown" } elseif ([bool]$dkimEnabled) { "Enabled" } else { "Disabled" }
            Dmarc      = if ($dmarcRecord -ne "Not queried" -and $dmarcRecord) { "Present" } else { "Missing" }
            DmarcValue = $dmarcRecord
        }
    }

    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "DkimDmarc" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 DKIM / DMARC Report" -Tenant $script:ToolboxTenantLabel -Subtitle "Mail domain signing posture for accepted domains" -Kpis @(
        @{ label = "Domains"; value = $rows.Count; sub = "Accepted domains"; cls = "neutral" },
        @{ label = "DKIM Enabled"; value = @($rows | Where-Object { $_.Dkim -eq "Enabled" }).Count; sub = "Domains signing outbound mail"; cls = "ok" },
        @{ label = "DMARC Present"; value = @($rows | Where-Object { $_.Dmarc -eq "Present" }).Count; sub = "Domains with DMARC"; cls = "ok" },
        @{ label = "DMARC Missing"; value = @($rows | Where-Object { $_.Dmarc -eq "Missing" }).Count; sub = "Domains without DMARC"; cls = if (@($rows | Where-Object { $_.Dmarc -eq "Missing" }).Count -gt 0) { "warn" } else { "ok" } }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Mail Domain Security"
            badge = "$($rows.Count) domains"
            columns = @(
                @{ key = "Domain"; header = "Domain" },
                @{ key = "DomainType"; header = "Type" },
                @{ key = "Default"; header = "Default"; type = "pill" },
                @{ key = "Dkim"; header = "DKIM"; type = "pill" },
                @{ key = "Dmarc"; header = "DMARC"; type = "pill" },
                @{ key = "DmarcValue"; header = "DMARC Record" }
            )
            rows = @($rows | Sort-Object Domain)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
