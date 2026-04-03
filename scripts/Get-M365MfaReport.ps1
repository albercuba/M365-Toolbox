<#
.SYNOPSIS
    Generates a comprehensive MFA authentication methods report for all users
    in a Microsoft 365 tenant.
...
        $connectParams = @{
            Scopes                  = $scopes
            NoWelcome               = $true
            ContextScope            = "Process"
            UseDeviceAuthentication = $true
            ErrorAction             = "Stop"
            InformationAction       = "Continue"
        }
...
Write-Host "====================================================`n" -ForegroundColor Cyan

$xlAvailable = Assert-RequiredModules -NeedExcelExport ([bool]$ExportXlsx)
$null = Connect-ToGraph

$adminIds   = Get-AdminUserIds -RoleNames $AdminRoles
...