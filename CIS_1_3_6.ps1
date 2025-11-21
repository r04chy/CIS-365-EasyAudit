<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.6 (L2): Ensure the customer lockbox feature is enabled (Automated)

Intent:
  Customer Lockbox requires Microsoft support engineers to get explicit approval
  before accessing your tenant data, and logs that access. CIS requires this to be enabled.

Implementation:
  In Exchange Online, this is controlled by the organization config flag:
    CustomerLockboxEnabled = $true  --> COMPLIANT

This script:
  - Connects to Exchange Online
  - Reads Get-OrganizationConfig.CustomerLockboxEnabled
  - Evaluates CIS 1.3.6 (enabled vs disabled)
  - Outputs PASS/FAIL and sets $global:CISCheckResult
  - Optionally exports to CSV

REQUIRES:
  - Exchange Online PowerShell V3 module:
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
  - Permissions:
        Exchange Administrator (or Global Admin)
#>

[CmdletBinding()]
param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Exchange Online
# -----------------------------------------------------------
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
catch {
    Write-Error "ExchangeOnlineManagement module not found. Install it with:`n  Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    return
}

try {
    Connect-ExchangeOnline -ShowProgress:$false
}
catch {
    Write-Error "Failed to connect to Exchange Online. Check your permissions and network connectivity."
    return
}

# -----------------------------------------------------------
# 2. Retrieve Customer Lockbox setting
# -----------------------------------------------------------
Write-Host "Retrieving organization configuration (CustomerLockboxEnabled)..." -ForegroundColor Cyan

$orgConfig = $null
try {
    $orgConfig = Get-OrganizationConfig -ErrorAction Stop | Select-Object CustomerLockboxEnabled
}
catch {
    Write-Error "Failed to retrieve organization configuration. Error: $_"
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

if (-not $orgConfig) {
    Write-Error "Organization configuration was not returned."
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

$lockboxEnabled = $orgConfig.CustomerLockboxEnabled

# -----------------------------------------------------------
# 3. Evaluate CIS 1.3.6
#    Requirement: CustomerLockboxEnabled = $true
# -----------------------------------------------------------
$compliant = ($lockboxEnabled -eq $true)

$result = [pscustomobject]@{
    Control             = "1.3.6 (L2)"
    SettingName         = "Customer Lockbox"
    CustomerLockboxEnabled = $lockboxEnabled
    Compliant           = $compliant
}

# -----------------------------------------------------------
# 4. Output summary & PASS/FAIL
# -----------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.3.6 (L2) – Customer Lockbox Feature =====" -ForegroundColor Cyan
Write-Host ("CustomerLockboxEnabled : {0}" -f $lockboxEnabled)
Write-Host "Requirement: CustomerLockboxEnabled = True" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

$result | Format-Table -AutoSize

if ($compliant) {
    Write-Host "`nRESULT: PASS – Customer Lockbox is enabled." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}
else {
    Write-Host "`nRESULT: FAIL – Customer Lockbox is NOT enabled." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}

# -----------------------------------------------------------
# 5. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $result | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nCustomer Lockbox setting report exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report. Error: $_"
    }
}

# -----------------------------------------------------------
# 6. Cleanup
# -----------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nCIS Control 1.3.6 check complete.`n" -ForegroundColor Cyan
