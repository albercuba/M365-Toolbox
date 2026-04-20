[CmdletBinding()]
param(
    [string[]]$MailboxIdentity,
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

    Write-SectionHeader "COLLECTING MAILBOX PERMISSION DATA"
    $requestedMailboxFilters = @(Normalize-DelimitedValue -Value $MailboxIdentity)
    $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress,Alias)

    if ($requestedMailboxFilters.Count -gt 0) {
        $filterLookup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $requestedMailboxFilters) {
            [void]$filterLookup.Add($entry)
        }

        $mailboxes = @(
            $mailboxes | Where-Object {
                $candidates = @(
                    [string]$_.DisplayName,
                    [string]$_.Alias,
                    [string]$_.PrimarySmtpAddress,
                    [string]$_.UserPrincipalName
                ) | Where-Object { $_ }

                @($candidates | Where-Object { $filterLookup.Contains($_) }).Count -gt 0
            }
        )
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    $detailRows = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $mailboxes) {
        $mailboxPerms = @(Get-ExoMailboxPermission -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object { $_.User -notmatch "NT AUTHORITY|S-1-5|SELF" -and -not $_.IsInherited })
        $recipientPerms = @(Get-RecipientPermission -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object { $_.Trustee -and $_.Trustee -notmatch "NT AUTHORITY|S-1-5|SELF" })
        $mailboxLabel = if ($mailbox.DisplayName -and $mailbox.UserPrincipalName) {
            "{0} ({1})" -f $mailbox.DisplayName, $mailbox.UserPrincipalName
        }
        elseif ($mailbox.DisplayName) {
            [string]$mailbox.DisplayName
        }
        else {
            [string]$mailbox.UserPrincipalName
        }

        foreach ($permission in $mailboxPerms) {
            [void]$detailRows.Add([pscustomobject]@{
                Mailbox        = $mailboxLabel
                Delegate       = [string]$permission.User
                PermissionType = "Full Access"
            })
        }

        foreach ($permission in $recipientPerms) {
            [void]$detailRows.Add([pscustomobject]@{
                Mailbox        = $mailboxLabel
                Delegate       = [string]$permission.Trustee
                PermissionType = "Send As"
            })
        }

        [void]$rows.Add([pscustomobject]@{
            DisplayName       = [string]$mailbox.DisplayName
            UserPrincipalName = [string]$mailbox.UserPrincipalName
            MailboxFilter     = $mailboxLabel
            FullAccess        = $mailboxPerms.Count
            SendAs            = $recipientPerms.Count
            Exposure          = if (($mailboxPerms.Count + $recipientPerms.Count) -gt 3) { "Elevated" } elseif (($mailboxPerms.Count + $recipientPerms.Count) -gt 0) { "Delegated" } else { "Direct Owner Only" }
        })
    }

    $delegated = @($rows | Where-Object { $_.Exposure -ne "Direct Owner Only" })
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailboxPermissions" -OutputPath $OutputPath

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mailbox Permission Audit" -Tenant $script:ToolboxTenantLabel -Subtitle "Mailbox delegate exposure across user and shared mailboxes" -Kpis @(
        @{ label = "Mailboxes"; value = $rows.Count; sub = "Scoped mailboxes"; cls = "neutral" },
        @{ label = "Delegated"; value = $delegated.Count; sub = "Mailboxes with delegates"; cls = if ($delegated.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "Full Access"; value = (($rows | Measure-Object FullAccess -Sum).Sum); sub = "Assignments"; cls = "neutral" },
        @{ label = "Send As"; value = (($rows | Measure-Object SendAs -Sum).Sum); sub = "Assignments"; cls = "neutral" }
    ) -StripItems @(
        @{ label = "Tenant"; value = $script:ToolboxTenantLabel },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") },
        @{ label = "Mailbox Filter"; value = if ($requestedMailboxFilters.Count -gt 0) { $requestedMailboxFilters -join ", " } else { "All user and shared mailboxes" } }
    ) -Sections @(
        @{
            id = "mailbox-permission-summary"
            title = "Mailbox Permission Summary"
            badge = "$($rows.Count) mailboxes"
            rowAction = @{
                targetSectionId = "mailbox-permission-details"
                sourceKey = "MailboxFilter"
                ariaLabel = "Open mailbox permission details for this mailbox"
            }
            columns = @(
                @{ key = "DisplayName"; header = "Mailbox" },
                @{ key = "UserPrincipalName"; header = "UPN" },
                @{ key = "FullAccess"; header = "Full Access" },
                @{ key = "SendAs"; header = "Send As" },
                @{ key = "Exposure"; header = "Exposure"; type = "pill" }
            )
            rows = @($rows | Sort-Object FullAccess -Descending)
        },
        @{
            id = "mailbox-permission-details"
            title = "Mailbox Permission Details"
            badge = "$($detailRows.Count) assignment(s)"
            filterKey = "Mailbox"
            filterLabel = "Mailbox Filter"
            filterPlaceholder = "Type a Mailbox name or UPN"
            columns = @(
                @{ key = "Mailbox"; header = "Mailbox" },
                @{ key = "Delegate"; header = "Delegate" },
                @{ key = "PermissionType"; header = "Permission Type"; type = "pill" }
            )
            rows = @($detailRows | Sort-Object Mailbox, PermissionType, Delegate)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
