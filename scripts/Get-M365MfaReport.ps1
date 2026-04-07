param(
    [Parameter(Mandatory)]
    [string]$TenantId
)

function Connect-ToGraph {
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )

    $ErrorActionPreference = 'Stop'

    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "          CONNECTING TO MICROSOFT GRAPH" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "[*] Requested tenant: $TenantId"
    Write-Host "[*] Starting device code authentication..." -ForegroundColor Yellow
    Write-Host "[*] When the code appears, open https://login.microsoft.com/device" -ForegroundColor Yellow

    Disconnect-MgGraph -ErrorAction SilentlyContinue

    try {
        Connect-MgGraph `
            -TenantId $TenantId `
            -Scopes "User.Read.All","Directory.Read.All","AuditLog.Read.All" `
            -UseDeviceAuthentication `
            -NoWelcome `
            -ContextScope Process `
            -ErrorAction Stop

        Write-Host "[+] Connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        $msg = $_.Exception.Message

        if ($msg -match "AADSTS90002") {
            throw "Tenant '$TenantId' was not found. Check the tenant name."
        }

        throw "Failed to connect to Microsoft Graph: $msg"
    }
}

Connect-ToGraph -TenantId $TenantId
