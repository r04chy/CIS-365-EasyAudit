<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.3.4 (L1): Ensure 'User owned apps and services' is restricted (Automated)

Intent:
  Users should not be able to install or manage their own Outlook add-ins / mailbox apps.
  In Exchange Online, this is implemented by ensuring that the default Role Assignment Policy
  does NOT include the following roles:
    - My Custom Apps
    - My Marketplace Apps
    - My ReadWriteMailbox Apps

This script:
  - Connects to Exchange Online
  - Locates the default Role Assignment Policy
  - Checks whether those three roles are assigned to it
  - Outputs per-role compliance and an overall CIS PASS/FAIL
  - Optionally exports results to CSV
  - Sets $global:CISCheckResult = "PASS" or "FAIL"

REQUIRES:
  - Exchange Online PowerShell V3 module:
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
  - Permissions:
        Exchange Administrator (or higher)
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
# 2. Locate the default Role Assignment Policy
# -----------------------------------------------------------
Write-Host "Retrieving Role Assignment Policies..." -ForegroundColor Cyan

$rolePolicies = $null
try {
    $rolePolicies = Get-RoleAssignmentPolicy -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve role assignment policies. Error: $_"
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

if (-not $rolePolicies) {
    Write-Error "No Role Assignment Policies were returned from Exchange Online."
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

# Prefer IsDefault flag; fall back to "Default Role Assignment Policy"
$defaultPolicy = $rolePolicies | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1

if (-not $defaultPolicy) {
    $defaultPolicy = $rolePolicies | Where-Object { $_.Name -eq "Default Role Assignment Policy" } | Select-Object -First 1
}

if (-not $defaultPolicy) {
    Write-Error "Could not identify the default Role Assignment Policy."
    $global:CISCheckResult = "FAIL"
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

Write-Host ("Using default Role Assignment Policy: {0}" -f $defaultPolicy.Name) -ForegroundColor Cyan

# -----------------------------------------------------------
# 3. Check user-owned app roles on the default policy
# -----------------------------------------------------------
$rolesToCheck = @(
    "My Custom Apps",
    "My Marketplace Apps",
    "My ReadWriteMailbox Apps"
)

$results = @()

foreach ($roleName in $rolesToCheck) {
    $assigned = $false
    $assignmentObjects = $null

    try {
        $assignmentObjects = Get-ManagementRoleAssignment `
            -Role $roleName `
            -RoleAssignee $defaultPolicy.Name `
            -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Error while checking role '$roleName' on policy '$($defaultPolicy.Name)': $_"
    }

    if ($assignmentObjects -and $assignmentObjects.Count -gt 0) {
        $assigned = $true
    }

    $results += [pscustomobject]@{
        Control              = "1.3.4 (L1)"
        DefaultPolicyName    = $defaultPolicy.Name
        RoleName             = $roleName
        RoleAssignedToPolicy = $assigned
        CompliantForRole     = (-not $assigned)   # Compliant if NOT assigned
    }
}

# -----------------------------------------------------------
# 4. Overall CIS evaluation
#    Requirement: NONE of the roles are assigned to the default policy
# -----------------------------------------------------------
$nonCompliant = $results | Where-Object { $_.RoleAssignedToPolicy -eq $true }
$compliant    = $results | Where-Object { $_.RoleAssignedToPolicy -eq $false }

Write-Host ""
Write-Host "===== CIS 1.3.4 (L1) – User Owned Apps and Services =====" -ForegroundColor Cyan
Write-Host ("Default Role Assignment Policy : {0}" -f $defaultPolicy.Name)
Write-Host "The following roles should NOT be assigned to this policy:" -ForegroundColor Cyan
Write-Host "  - My Custom Apps" -ForegroundColor Cyan
Write-Host "  - My Marketplace Apps" -ForegroundColor Cyan
Write-Host "  - My ReadWriteMailbox Apps" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$results | Format-Table DefaultPolicyName, RoleName, RoleAssignedToPolicy, CompliantForRole -AutoSize

if ($nonCompliant.Count -gt 0) {
    Write-Host "`nNon-compliant roles found on the default Role Assignment Policy:" -ForegroundColor Red
    $nonCompliant | Format-Table RoleName, RoleAssignedToPolicy -AutoSize

    Write-Host "`nRESULT: FAIL – One or more 'user owned app' roles are assigned to the default Role Assignment Policy." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
}
else {
    Write-Host "`nRESULT: PASS – No 'user owned app' roles are assigned to the default Role Assignment Policy." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}

# -----------------------------------------------------------
# 5. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    try {
        $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nUser owned apps/services role assignment report exported to: $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to export CSV report. Error: $_"
    }
}

# -----------------------------------------------------------
# 6. Cleanup
# -----------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nCIS Control 1.3.4 check complete.`n" -ForegroundColor Cyan
