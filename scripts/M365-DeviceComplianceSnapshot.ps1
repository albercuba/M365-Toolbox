[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath,
    [string]$ExportHtml
)

. (Join-Path $PSScriptRoot "Shared-ToolboxReport.ps1")

Assert-GraphModules -RequiredModules @("Microsoft.Graph.Authentication", "Microsoft.Graph.DeviceManagement")
Connect-ToolboxGraph -TenantId $TenantId -Scopes @("DeviceManagementManagedDevices.Read.All")
Resolve-ToolboxTenantLabel

Write-SectionHeader "COLLECTING DEVICE COMPLIANCE DATA"

$devices = @()
$collectionWarnings = @()

try {
    $devices = @(Get-MgDeviceManagementManagedDevice -All -Property "deviceName,userPrincipalName,operatingSystem,complianceState,managedDeviceOwnerType,lastSyncDateTime" -ErrorAction Stop)
}
catch {
    $errorMessage = Get-ToolboxExceptionMessage -Exception $_.Exception
    if ($errorMessage -match 'Request not applicable to target tenant') {
        $collectionWarnings += "Managed devices are not available for this tenant. This usually means Intune or Microsoft Endpoint Manager is not enabled, licensed, or configured for the connected tenant."
    }
    else {
        $collectionWarnings += "Managed devices query failed: $errorMessage"
    }
}

foreach ($warningMessage in $collectionWarnings) {
    Write-Warning "  [!] $warningMessage"
}

$rows = @(
    foreach ($device in $devices) {
        [pscustomobject]@{
            DeviceName      = [string]$device.deviceName
            UserPrincipal   = [string]$device.userPrincipalName
            OS              = [string]$device.operatingSystem
            ComplianceState = [string]$device.complianceState
            Ownership       = [string]$device.managedDeviceOwnerType
            LastSync        = if ($device.lastSyncDateTime) { (Get-Date $device.lastSyncDateTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
        }
    }
)

$htmlPath = Add-TimestampToPath -Path $ExportHtml -BaseName "DeviceCompliance" -OutputPath $OutputPath
$tenantName = if ($script:ToolboxTenantLabel) { $script:ToolboxTenantLabel } else { "Unknown tenant" }

Export-ToolboxHtmlReport -Path $htmlPath -Title "M365 Device Compliance Snapshot" -Tenant $tenantName -Subtitle "Managed device compliance and ownership summary" -Kpis @(
    @{ label = "Devices"; value = $rows.Count; sub = "Managed devices"; cls = "neutral" },
    @{ label = "Compliant"; value = @($rows | Where-Object { $_.ComplianceState -eq "compliant" }).Count; sub = "Passing compliance"; cls = "ok" },
    @{ label = "Noncompliant"; value = @($rows | Where-Object { $_.ComplianceState -eq "noncompliant" }).Count; sub = "Failing compliance"; cls = "crit" },
    @{ label = "OS Types"; value = @($rows | Select-Object -ExpandProperty OS -Unique).Count; sub = "Platform count"; cls = "neutral" }
) -StripItems @(
    @{ label = "Tenant"; value = $tenantName },
    @{ label = "Generated"; value = (Get-Date).ToString("yyyy-MM-dd HH:mm") }
) -Sections @(
    @{
        title = "Collection Status"
        badge = if ($collectionWarnings.Count -gt 0) { "Warning" } else { "Healthy" }
        text = if ($collectionWarnings.Count -gt 0) { ($collectionWarnings -join "`n") } else { "Managed device data was collected successfully." }
    },
    @{
        title = "Managed Devices"
        badge = "$($rows.Count) devices"
        columns = @(
            @{ key = "DeviceName"; header = "Device" },
            @{ key = "UserPrincipal"; header = "User" },
            @{ key = "OS"; header = "OS" },
            @{ key = "ComplianceState"; header = "Compliance"; type = "pill" },
            @{ key = "Ownership"; header = "Ownership"; type = "pill" },
            @{ key = "LastSync"; header = "Last Sync" }
        )
        rows = @($rows | Sort-Object ComplianceState, DeviceName)
    }
)

Write-Host "[+] HTML dashboard exported to: $htmlPath" -ForegroundColor Green
