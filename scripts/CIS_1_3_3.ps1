
<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.3 (L2): Ensure 'External sharing' of calendars is not available (Automated)

This script:
  - Connects to Exchange Online
  - Retrieves the "Default Sharing Policy"
  - Checks whether it is disabled (Enabled = $False)
  - Outputs:
      * The Default Sharing Policy state
      * PASS/FAIL for CIS 1.3.3
  - Optionally exports a CSV with the result

REQUIRES:
  - Exchange Online PowerShell V3 module:
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
  - Permissions:
        Exchange Online admin role (e.g. Exchange Administrator, or higher)
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
# 2. Retrieve the Default Sharing Policy
# -----------------------------------------------------------
Write-Host "Retrieving 'Default Sharing Policy'..." -ForegroundColor Cyan

$policy = $null
try {
    $policy = Get-SharingPolicy -Identity "Default Sharing Policy" -ErrorAction Stop
}
catch {
    Write-Error "Could not retrieve 'Default Sharing Policy'. Verify it exists in your tenant. Error: $_"
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

if (-not $policy) {
    Write-Error "'Default Sharing Policy' was not returned from Exchange Online."
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

# -----------------------------------------------------------
# 3. Evaluate CIS 1.3.3
#    Requirement: Default Sharing Policy Enabled = $False
# -----------------------------------------------------------
$enabled   = $policy.Enabled
$compliant = ($enabled -eq $false)

$result = [pscustomobject]@{
    Control   = "1.3.3 (L2)"
    PolicyName= $policy.Name
    Enabled   = $enabled
    Compliant = $compliant
}

# -----------------------------------------------------------
# 4. Output summary & PASS/FAIL
# -----------------------------------------------------------
Write-Host ""
Write-Host "===== CIS 1.3.3 (L2) – External Calendar Sharing Disabled =====" -ForegroundColor Cyan
Write-Host "Policy checked      : $($policy.Name)"
Write-Host "Enabled (should be False) : $enabled"
Write-Host "===============================================================" -ForegroundColor Cyan

$result | Format-Table -AutoSize

if ($compliant) {
    Write-Host "`nRESULT: PASS – 'Default Sharing Policy' is disabled; external calendar sharing is not available." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}
else {
    Write-Host "`nRESULT: FAIL – 'Default Sharing Policy' is enabled; external calendar sharing is available." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}

# -----------------------------------------------------------
# 5. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $result | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nExternal calendar sharing report exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report. Error: $_"
    }
}

# -----------------------------------------------------------
# 6. Cleanup
# -----------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nCIS Control 1.3.3 check complete.`n" -ForegroundColor Cyan
