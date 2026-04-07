[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

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
    Write-SectionHeader "COLLECTING DISTRIBUTION GROUP DATA"

    $groups = @(Get-DistributionGroup -ResultSize Unlimited)
    $rows = foreach ($group in $groups) {
        $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
        [pscustomobject]@{
            Name                  = [string]$group.DisplayName
            PrimarySmtpAddress    = [string]$group.PrimarySmtpAddress
            Members               = $members.Count
            ManagedBy             = if ($group.ManagedBy) { (@($group.ManagedBy) -join ", ") } else { "None" }
            RequireAuthToSend     = if ($group.RequireSenderAuthenticationEnabled) { "Internal Only" } else { "External Senders Allowed" }
            HiddenFromAddressBook = if ($group.HiddenFromAddressListsEnabled) { "Yes" } else { "No" }
        }
    }

    $openToExternal = @($rows | Where-Object { $_.RequireAuthToSend -eq "External Senders Allowed" })
    $ownerless = @($rows | Where-Object { $_.ManagedBy -eq "None" })
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "DistributionGroups" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Distribution Group Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Distribution group exposure and ownership hygiene" -Kpis @(
        @{ label = "Groups"; value = $rows.Count; sub = "Distribution groups"; cls = "neutral" },
        @{ label = "External Senders"; value = $openToExternal.Count; sub = "Groups open to external mail"; cls = if ($openToExternal.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Ownerless"; value = $ownerless.Count; sub = "Groups without managers"; cls = if ($ownerless.Count -gt 0) { "crit" } else { "ok" } },
        @{ label = "Members"; value = (($rows | Measure-Object Members -Sum).Sum); sub = "Total memberships"; cls = "neutral" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Distribution Groups"
            badge = "$($rows.Count) groups"
            columns = @(
                @{ key = "Name"; header = "Group" },
                @{ key = "PrimarySmtpAddress"; header = "Primary SMTP" },
                @{ key = "Members"; header = "Members" },
                @{ key = "ManagedBy"; header = "Managed By" },
                @{ key = "RequireAuthToSend"; header = "Sender Scope"; type = "pill" },
                @{ key = "HiddenFromAddressBook"; header = "Hidden"; type = "pill" }
            )
            rows = @($rows | Sort-Object Members -Descending)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
