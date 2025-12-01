<#
.SYNOPSIS
CIS Microsoft 365 Foundations Benchmark v6.0.0
Control 2.1.2 (L1) – Ensure the Common Attachment Types Filter is enabled

.DESCRIPTION
Audits Exchange Online anti-malware policy to ensure the Common Attachment
Types Filter is enabled on the highest-priority policy. If no enabled malware
filter rules are found, falls back to the 'Default' policy.

PASS:
 - EnableFileFilter = True

FAIL:
 - EnableFileFilter is False or missing.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.2'
$ControlTitle = 'Ensure the Common Attachment Types Filter is enabled'

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
        $policyName = 'Default'
        Add-Detail "No enabled Malware Filter rules found; falling back to 'Default' policy."
    }

    $policy = Get-MalwareFilterPolicy -Identity $policyName -ErrorAction Stop

    if ($policy.PSObject.Properties.Name -contains 'EnableFileFilter') {
        $enableFileFilter = [bool]$policy.EnableFileFilter
        Add-Detail ("Policy '{0}': EnableFileFilter = {1}" -f $policy.Identity, $enableFileFilter)
    }
    else {
        $enableFileFilter = $false
        Add-Detail ("Policy '{0}' does not expose EnableFileFilter; treating as non-compliant." -f $policy.Identity)
    }

    if ($enableFileFilter) {
        $Status = 'PASS'
        Add-Detail "Common Attachment Types Filter is enabled (EnableFileFilter = True)."
    }
    else {
        $Status = 'FAIL'
        Add-Detail "Common Attachment Types Filter is NOT enabled (EnableFileFilter is False or missing)."
        Add-Detail ""
        Add-Detail "Example remediation (PowerShell):"
        Add-Detail ("  Set-MalwareFilterPolicy -Identity '{0}' -EnableFileFilter $true" -f $policy.Identity)
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
