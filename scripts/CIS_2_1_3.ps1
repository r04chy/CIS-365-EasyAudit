<#
.SYNOPSIS
CIS Microsoft 365 Foundations Benchmark v6.0.0
Control 2.1.3 (L1) – Ensure notifications for internal users sending malware is Enabled

.DESCRIPTION
Audits Exchange Online anti-malware policy to ensure administrators are
notified when internal senders send messages containing malware.

PASS:
 - EnableInternalSenderAdminNotifications = True
 - InternalSenderAdminAddress is defined (non-empty)

FAIL:
 - Either flag is False, missing, or no admin address is defined.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.3'
$ControlTitle = 'Ensure notifications for internal users sending malware is Enabled'

$Status  = 'ERROR'
$Details = @()

function Add-Detail {
    param(
        [Parameter(Mandatory)][string]$Text
    )
    $script:Details += $Text
}

try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    $connectedHere = $false
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        Add-Detail "Using existing Exchange Online session."
    }
    catch {
        Add-Detail "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connectedHere = $true
    }

    # Determine highest-priority malware filter policy via rule
    $rules = Get-MalwareFilterRule -ErrorAction SilentlyContinue |
             Where-Object { $_.State -ne 'Disabled' } |
             Sort-Object Priority

    $policyName = $null

    if ($rules -and $rules.Count -gt 0) {
        $primaryRule = $rules | Select-Object -First 1
        $policyName  = $primaryRule.MalwareFilterPolicy
        Add-Detail ("Highest-priority enabled Malware Filter rule: Name = '{0}', Priority = {1}, MalwareFilterPolicy = '{2}'" -f `
            $primaryRule.Name, $primaryRule.Priority, $policyName)
    }
    else {
        Add-Detail "No enabled Malware Filter rules found; evaluating all Malware Filter policies."
        # When no rules exist, evaluate Default policy by name
        $policyName = 'Default'
    }

    $policy = Get-MalwareFilterPolicy -Identity $policyName -ErrorAction Stop

    $notifyEnabled = $false
    $adminAddress  = $null

    if ($policy.PSObject.Properties.Name -contains 'EnableInternalSenderAdminNotifications') {
        $notifyEnabled = [bool]$policy.EnableInternalSenderAdminNotifications
        Add-Detail ("Policy '{0}': EnableInternalSenderAdminNotifications = {1}" -f $policy.Identity, $notifyEnabled)
    }
    else {
        Add-Detail ("Policy '{0}' does not expose EnableInternalSenderAdminNotifications; treating as non-compliant." -f $policy.Identity)
    }

    if ($policy.PSObject.Properties.Name -contains 'InternalSenderAdminAddress') {
        $adminAddress = $policy.InternalSenderAdminAddress
        Add-Detail ("Policy '{0}': InternalSenderAdminAddress = {1}" -f $policy.Identity, ($adminAddress -join ', '))
    }
    else {
        Add-Detail ("Policy '{0}' does not expose InternalSenderAdminAddress; treating as missing." -f $policy.Identity)
    }

    # Determine whether an admin notification address is present
    $hasAddress = $false
    if ($null -ne $adminAddress) {
        if ($adminAddress -is [System.Array]) {
            $hasAddress = ($adminAddress.Count -gt 0)
        }
        else {
            $hasAddress = -not [string]::IsNullOrWhiteSpace([string]$adminAddress)
        }
    }

    if ($notifyEnabled -and $hasAddress) {
        $Status = 'PASS'
        Add-Detail "Notifications for internal users sending malware are enabled and an administrator address is configured."
    }
    else {
        $Status = 'FAIL'
        Add-Detail "Notifications for internal users sending malware are NOT fully configured."
        if (-not $notifyEnabled) {
            Add-Detail "  - EnableInternalSenderAdminNotifications is False or missing."
        }
        if (-not $hasAddress) {
            Add-Detail "  - InternalSenderAdminAddress is not defined or empty."
        }

        Add-Detail ""
        Add-Detail "Example remediation (PowerShell):"
        Add-Detail ("  Set-MalwareFilterPolicy -Identity '{0}' -EnableInternalSenderAdminNotifications $true -InternalSenderAdminAddress 'security@example.com'" -f $policy.Identity)
    }

}
catch {
    if ($Status -ne 'FAIL') {
        $Status = 'ERROR'
    }
    Add-Detail ("Error evaluating {0}: {1}" -f $ControlId, $_.Exception.Message)
}
finally {
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Add-Detail "Disconnected temporary Exchange Online session."
    }
}

$global:CISCheckResult = $Status

[PSCustomObject]@{
    Control = $ControlId
    Title   = $ControlTitle
    Status  = $Status
    Details = $Details -join "`n"
}
