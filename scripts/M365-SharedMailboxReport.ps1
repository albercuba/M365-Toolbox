[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml,
    [switch]$IncludePermissions
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Write-Host ""
Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
Write-Host "  [+] ExchangeOnlineManagement" -ForegroundColor Green

Write-SectionHeader "CONNECTING TO EXCHANGE ONLINE"
if ($TenantId) {
    Write-Host "[*] Requested tenant: $TenantId"
}
else {
    Write-Host "[*] Requested tenant: auto-detect"
}
Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
$exoParams = @{
    ShowBanner   = $false
    ShowProgress = $false
    Device       = $true
}
if ($TenantId) {
    $exoParams.Organization = $TenantId
}
Connect-ExchangeOnline @exoParams | Out-Null

try {
    Write-Host "[+] Connected to Exchange Online" -ForegroundColor Green
    $script:ToolboxTenantLabel = if ($TenantId) { $TenantId } else { "ExchangeOnline" }

    Write-SectionHeader "COLLECTING SHARED MAILBOX DATA"
    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties DisplayName,PrimarySmtpAddress,UserPrincipalName,WhenCreated,ForwardingSmtpAddress,HiddenFromAddressListsEnabled)
    $rows = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $mailboxes) {
        $permissionCount = ""
        if ($IncludePermissions) {
            try {
                $permissions = @(Get-ExoMailboxPermission -Identity $mailbox.UserPrincipalName -ErrorAction Stop | Where-Object { $_.User -notmatch "NT AUTHORITY|S-1-5|SELF" -and -not $_.IsInherited })
                $permissionCount = $permissions.Count
            }
            catch {
                $permissionCount = "N/A"
            }
        }

        [void]$rows.Add([pscustomobject]@{
            DisplayName     = [string]$mailbox.DisplayName
            PrimarySmtp     = [string]$mailbox.PrimarySmtpAddress
            Created         = if ($mailbox.WhenCreated) { (Get-Date $mailbox.WhenCreated).ToString("yyyy-MM-dd") } else { "" }
            HiddenInGAL     = if ($mailbox.HiddenFromAddressListsEnabled) { "Yes" } else { "No" }
            Forwarding      = if ($mailbox.ForwardingSmtpAddress) { "Yes" } else { "No" }
            PermissionCount = $permissionCount
        })
    }

    $noDelegateCount = @($rows | Where-Object { $_.PermissionCount -eq 0 }).Count
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "SharedMailbox" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Shared Mailbox Report" -Tenant $script:ToolboxTenantLabel -Subtitle "Shared mailbox inventory, forwarding, and delegate visibility" -Kpis @(
        @{ label = "Shared"; value = $rows.Count; sub = "Shared mailboxes"; cls = "neutral" },
        @{ label = "Forwarding"; value = @($rows | Where-Object { $_.Forwarding -eq "Yes" }).Count; sub = "With forwarding"; cls = "warn" },
        @{ label = "Hidden"; value = @($rows | Where-Object { $_.HiddenInGAL -eq "Yes" }).Count; sub = "Hidden from GAL"; cls = "neutral" },
        @{ label = "No Delegates"; value = $noDelegateCount; sub = if ($IncludePermissions) { "Zero explicit permissions" } else { "Permissions skipped" }; cls = if ($IncludePermissions -and $noDelegateCount -gt 0) { "warn" } else { "neutral" } }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Permissions"; value = if ($IncludePermissions) { "Included" } else { "Skipped" } },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Shared Mailboxes"
            badge = "$($rows.Count) mailboxes"
            columns = @(
                @{ key = "DisplayName"; header = "Name" },
                @{ key = "PrimarySmtp"; header = "Primary SMTP" },
                @{ key = "Created"; header = "Created" },
                @{ key = "HiddenInGAL"; header = "Hidden"; type = "pill" },
                @{ key = "Forwarding"; header = "Forwarding"; type = "pill" },
                @{ key = "PermissionCount"; header = "Delegates" }
            )
            rows = @($rows | Sort-Object DisplayName)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
