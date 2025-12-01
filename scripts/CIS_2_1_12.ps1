<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.12 (L1) – Ensure the connection filter IP allow list is not used

DESCRIPTION
Checks HostedConnectionFilterPolicy objects to ensure IPAllowList is empty.

PASS:
 - All policies have empty / undefined IPAllowList.

FAIL:
 - Any policy has one or more IPAllowList entries.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.12'
$ControlTitle = 'Ensure the connection filter IP allow list is not used'

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

    $policies = Get-HostedConnectionFilterPolicy -ErrorAction Stop
    if (-not $policies) {
        $Status = 'FAIL'
        Add-Detail "No HostedConnectionFilterPolicy objects found."
    } else {
        $violations = @()

        foreach ($p in $policies) {
            $list = $p.IPAllowList
            Add-Detail ("Policy '{0}': IPAllowList = {1}" -f $p.Identity, ($list -join ', '))
            if ($list -and $list.Count -gt 0) {
                $violations += $p.Identity
            }
        }

        if ($violations.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "No HostedConnectionFilterPolicy has IPAllowList entries defined."
        } else {
            $Status = 'FAIL'
            Add-Detail "The following policies have IPAllowList entries:"
            foreach ($v in $violations) { Add-Detail ("  - {0}" -f $v) }
        }
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
