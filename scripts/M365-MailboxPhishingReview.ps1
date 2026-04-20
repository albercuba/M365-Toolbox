[CmdletBinding()]
param(
    [string]$TenantId,
    [Parameter(Mandatory)]
    [string]$MailboxUpn,
    [ValidateSet("inbox", "junkemail", "deleteditems")]
    [string]$Folder = "inbox",
    [int]$LookbackDays = 14,
    [int]$MaxMessages = 100,
    [string]$SubjectKeywords = "password,reset,invoice,mfa,urgent,verify,login,shared document,voicemail",
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

function Connect-ToolboxExchange {
    param([string]$Tenant)

    Write-Host ""
    Write-Host "[*] Checking required PowerShell modules..." -ForegroundColor Cyan
    Import-Module ExchangeOnlineManagement -Force -WarningAction SilentlyContinue
    Write-Host "  [+] ExchangeOnlineManagement" -ForegroundColor Green

    Write-SectionHeader "CONNECTING TO EXCHANGE ONLINE"
    if ($Tenant) {
        Write-Host "[*] Requested tenant: $Tenant"
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
    if ($Tenant) {
        $exoParams.Organization = $Tenant
    }
    Connect-ExchangeOnline @exoParams | Out-Null
    Write-Host "[+] Connected to Exchange Online" -ForegroundColor Green
}

function Normalize-KeywordList {
    param([string]$Value)

    return @(
        ($Value -split "[,\r\n]+" | ForEach-Object { $_.Trim() }) |
            Where-Object { $_ } |
            Select-Object -Unique
    )
}

function Get-HeaderMap {
    param($Headers)

    $map = @{}
    foreach ($header in @($Headers)) {
        if (-not $header.name) { continue }
        $map[[string]$header.name] = [string]$header.value
    }
    return $map
}

function Get-TenantPrimaryDomain {
    if ($script:ToolboxTenantLabel -and $script:ToolboxTenantLabel -match '\.') {
        return $script:ToolboxTenantLabel.ToLowerInvariant()
    }

    return ""
}

function Test-DisplayNameMismatch {
    param(
        [string]$DisplayName,
        [string]$SenderAddress
    )

    if (-not $DisplayName -or -not $SenderAddress) {
        return $false
    }

    $commonBrands = @("microsoft", "office", "sharepoint", "onedrive", "outlook", "docusign", "adobe", "paypal")
    $senderDomain = ($SenderAddress -split "@")[-1].ToLowerInvariant()
    foreach ($brand in $commonBrands) {
        if ($DisplayName.ToLowerInvariant().Contains($brand) -and -not $senderDomain.Contains($brand)) {
            return $true
        }
    }

    return $false
}

function Get-PhishingAssessment {
    param(
        $Message,
        [string[]]$Keywords,
        [string]$TenantDomain
    )

    $fromAddress = [string]$Message.from.emailAddress.address
    $fromName = [string]$Message.from.emailAddress.name
    $subject = [string]$Message.subject
    $preview = [string]$Message.bodyPreview
    $headers = Get-HeaderMap -Headers $Message.internetMessageHeaders
    $reasons = [System.Collections.Generic.List[string]]::new()
    $score = 0

    $senderDomain = if ($fromAddress -match "@") { ($fromAddress -split "@")[-1].ToLowerInvariant() } else { "" }
    if ($senderDomain -and $TenantDomain -and $senderDomain -ne $TenantDomain) {
        $score += 1
        [void]$reasons.Add("External sender")
    }

    if (Test-DisplayNameMismatch -DisplayName $fromName -SenderAddress $fromAddress) {
        $score += 2
        [void]$reasons.Add("Display name and domain mismatch")
    }

    foreach ($keyword in $Keywords) {
        if ($subject -match [regex]::Escape($keyword)) {
            $score += 1
            [void]$reasons.Add("Subject keyword '$keyword'")
        }
    }

    if ($preview -match 'http[s]?://|bit\.ly|tinyurl|lnkd\.in|aka\.ms') {
        $score += 1
        [void]$reasons.Add("Link in body preview")
    }

    if ($Message.hasAttachments) {
        $score += 1
        [void]$reasons.Add("Has attachments")
    }

    $authBlob = (($headers["Authentication-Results"] + " " + $headers["Received-SPF"] + " " + $headers["X-MS-Exchange-Organization-AuthAs"]) -join " ").ToLowerInvariant()
    if ($authBlob -match 'spf=fail|dkim=fail|dmarc=fail|compauth=fail') {
        $score += 3
        [void]$reasons.Add("Authentication check failed")
    }

    if ($headers["Reply-To"] -and $headers["Reply-To"] -notmatch [regex]::Escape($fromAddress)) {
        $score += 2
        [void]$reasons.Add("Reply-To differs from sender")
    }

    $severity = if ($score -ge 6) {
        "High"
    }
    elseif ($score -ge 3) {
        "Medium"
    }
    else {
        "Low"
    }

    return [pscustomobject]@{
        Score    = $score
        Severity = $severity
        Reasons  = if ($reasons.Count -gt 0) { $reasons -join "; " } else { "No strong phishing indicators" }
        AuthHint = if ($authBlob) { $authBlob.Trim() } else { "No auth headers reviewed" }
    }
}

function Get-MessageTraceContext {
    param(
        [string]$Mailbox,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    $traceRows = @(Get-MessageTrace -RecipientAddress $Mailbox -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -ErrorAction Stop)
    return @($traceRows)
}

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.Mail")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("Mail.Read", "User.Read.All")
Resolve-ToolboxTenantLabel
Connect-ToolboxExchange -Tenant $TenantId

try {
    Write-SectionHeader "COLLECTING MAILBOX MESSAGES"

    $keywords = Normalize-KeywordList -Value $SubjectKeywords
    $tenantDomain = Get-TenantPrimaryDomain
    $filterStart = (Get-Date).AddDays(-1 * $LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $messages = @(Get-MgUserMailFolderMessage -UserId $MailboxUpn -MailFolderId $Folder -Top $MaxMessages -Sort "receivedDateTime desc" -Filter "receivedDateTime ge $filterStart" -Property Id,Subject,ReceivedDateTime,From,HasAttachments,BodyPreview,WebLink,InternetMessageId,InternetMessageHeaders -ErrorAction Stop)
    $traceStart = (Get-Date).AddDays(-1 * $LookbackDays)
    $traceEnd = Get-Date
    $messageTrace = @(Get-MessageTraceContext -Mailbox $MailboxUpn -StartDate $traceStart -EndDate $traceEnd)

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($message in $messages) {
        $assessment = Get-PhishingAssessment -Message $message -Keywords $keywords -TenantDomain $tenantDomain
        $fromAddress = [string]$message.from.emailAddress.address
        $subject = [string]$message.subject

        $traceMatch = @(
            $messageTrace | Where-Object {
                $_.SenderAddress -eq $fromAddress -and
                $_.RecipientAddress -eq $MailboxUpn -and
                $_.Subject -eq $subject
            } | Select-Object -First 1
        )

        [void]$rows.Add([pscustomobject]@{
            ReceivedDateTime = [string]$message.receivedDateTime
            Sender           = $fromAddress
            SenderName       = [string]$message.from.emailAddress.name
            Subject          = $subject
            HasAttachments   = if ($message.hasAttachments) { "Yes" } else { "No" }
            Score            = $assessment.Score
            Severity         = $assessment.Severity
            Reasons          = $assessment.Reasons
            AuthHint         = $assessment.AuthHint
            TraceStatus      = if ($traceMatch) { [string]$traceMatch[0].Status } else { "No message trace match" }
            TraceReceived    = if ($traceMatch) { [string]$traceMatch[0].Received } else { "" }
            Preview          = $message.bodyPreview
            WebLink          = [string]$message.webLink
        })
    }

    $suspiciousRows = @($rows | Where-Object { $_.Score -ge 3 } | Sort-Object Score -Descending, ReceivedDateTime -Descending)
    $htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "MailboxPhishingReview" -OutputPath $OutputPath
    $tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

    Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Mailbox Phishing Review" -Tenant $tenantName -Subtitle "Message review for potential phishing emails in a specific mailbox" -Kpis @(
        @{ label = "Mailbox"; value = $MailboxUpn; sub = "Target mailbox"; cls = "neutral" },
        @{ label = "Messages"; value = $rows.Count; sub = "Messages reviewed"; cls = "neutral" },
        @{ label = "Suspicious"; value = $suspiciousRows.Count; sub = "Messages scoring medium or high"; cls = if ($suspiciousRows.Count -gt 0) { "warn" } else { "ok" } },
        @{ label = "High Risk"; value = @($rows | Where-Object { $_.Severity -eq 'High' }).Count; sub = "Messages with strong phishing signals"; cls = if (@($rows | Where-Object { $_.Severity -eq 'High' }).Count -gt 0) { "crit" } else { "ok" } }
    ) -StripItems @(
        @{ label = "Mailbox"; value = $MailboxUpn },
        @{ label = "Folder"; value = $Folder },
        @{ label = "Lookback"; value = "$LookbackDays day(s)" },
        @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
    ) -Sections @(
        @{
            title = "Potential Phishing Messages"
            badge = "$($suspiciousRows.Count) suspicious message(s)"
            columns = @(
                @{ key = "ReceivedDateTime"; header = "Received" },
                @{ key = "Sender"; header = "Sender" },
                @{ key = "SenderName"; header = "Sender Name" },
                @{ key = "Subject"; header = "Subject" },
                @{ key = "HasAttachments"; header = "Attachments"; type = "pill" },
                @{ key = "Severity"; header = "Severity"; type = "pill" },
                @{ key = "Score"; header = "Score" },
                @{ key = "Reasons"; header = "Reasons" },
                @{ key = "TraceStatus"; header = "Trace Status"; type = "pill" },
                @{ key = "TraceReceived"; header = "Trace Received" },
                @{ key = "Preview"; header = "Preview"; type = "multiline" },
                @{ key = "WebLink"; header = "Open Message"; type = "link" }
            )
            rows = $suspiciousRows
        },
        @{
            title = "All Reviewed Messages"
            badge = "$($rows.Count) message(s)"
            columns = @(
                @{ key = "ReceivedDateTime"; header = "Received" },
                @{ key = "Sender"; header = "Sender" },
                @{ key = "Subject"; header = "Subject" },
                @{ key = "Severity"; header = "Severity"; type = "pill" },
                @{ key = "Score"; header = "Score" },
                @{ key = "TraceStatus"; header = "Trace Status"; type = "pill" },
                @{ key = "Reasons"; header = "Reasons" }
            )
            rows = @($rows | Sort-Object ReceivedDateTime -Descending)
        }
    )

    Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
