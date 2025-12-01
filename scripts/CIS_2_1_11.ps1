<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.11 (L2) – Ensure comprehensive attachment filtering is applied

DESCRIPTION
Audits the highest-priority Malware Filter policy to ensure:
 - EnableFileFilter = True
 - FileTypes list contains at least 120 entries.

PASS:
 - Both conditions satisfied.

FAIL:
 - EnableFileFilter is False, or FileTypes count < 120.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.11'
$ControlTitle = 'Ensure comprehensive attachment filtering is applied'

$Status  = 'ERROR'
$Details = @()

function Add-Detail {
    param([Parameter(Mandatory)][string]$Text)
    $script:Details += $Text
}

try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    $connectedHere = $false
    if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $null = Get-ConnectionInformation -ErrorAction Stop
            Add-Detail "Using existing Exchange Online session."
        } catch {
            Add-Detail "Connecting to Exchange Online..."
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $connectedHere = $true
        }
    } else {
        Add-Detail "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connectedHere = $true
    }

    $rules = Get-MalwareFilterRule -ErrorAction SilentlyContinue |
             Where-Object { $_.State -ne 'Disabled' } |
             Sort-Object Priority

    $policyName = $null
    if ($rules -and $rules.Count -gt 0) {
        $primaryRule = $rules | Select-Object -First 1
        $policyName  = $primaryRule.MalwareFilterPolicy
        Add-Detail ("Highest-priority enabled Malware Filter rule: Name='{0}', Priority={1}, MalwareFilterPolicy='{2}'" -f `
            $primaryRule.Name, $primaryRule.Priority, $policyName)
    } else {
        $policyName = 'Default'
        Add-Detail "No enabled Malware Filter rules found; falling back to 'Default' policy."
    }

    $policy = Get-MalwareFilterPolicy -Identity $policyName -ErrorAction Stop

    $enableFileFilter = $policy.EnableFileFilter
    $fileTypes        = $policy.FileTypes
    $count            = if ($fileTypes) { $fileTypes.Count } else { 0 }

    Add-Detail ("Policy '{0}': EnableFileFilter = {1}" -f $policy.Identity, $enableFileFilter)
    Add-Detail ("Policy '{0}': FileTypes count = {1}" -f $policy.Identity, $count)

    if ($enableFileFilter -and $count -ge 120) {
        $Status = 'PASS'
        Add-Detail "Malware filter policy has EnableFileFilter=True and at least 120 file extensions configured."
    } else {
        $Status = 'FAIL'
        if (-not $enableFileFilter) { Add-Detail "  - EnableFileFilter is not True." }
        if ($count -lt 120) { Add-Detail ("  - FileTypes contains only {0} entries; CIS recommends at least 120." -f $count) }

        Add-Detail ""
        Add-Detail "Example remediation (PowerShell snippet):"
        Add-Detail "  # Build a comprehensive extension list (120+ items) and apply:"
        Add-Detail "  `$exts = @('ace','ade', ... )  # see CIS appendix list"
        Add-Detail ("  Set-MalwareFilterPolicy -Identity '{0}' -EnableFileFilter $true -FileTypes $exts" -f $policy.Identity)
    }

} catch {
    if ($Status -ne 'FAIL') { $Status = 'ERROR' }
    Add-Detail ("Error evaluating {0}: {1}" -f $ControlId, $_.Exception.Message)
} finally {
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
