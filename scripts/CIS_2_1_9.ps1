<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.9 (L1) – Ensure that DKIM is enabled for all Exchange Online Domains

DESCRIPTION
Checks all DKIM signing configurations to ensure:
 - Enabled = True
 - Status  = 'Valid'

PASS:
 - All DKIM signing configs meet the above criteria.

FAIL:
 - Any DKIM config is disabled or not Valid.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.9'
$ControlTitle = 'Ensure that DKIM is enabled for all Exchange Online Domains'

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

    $configs = Get-DkimSigningConfig -ErrorAction Stop
    if (-not $configs -or $configs.Count -eq 0) {
        $Status = 'FAIL'
        Add-Detail "No DKIM signing configurations found."
    } else {
        $nonCompliant = @()

        foreach ($cfg in $configs) {
            Add-Detail ("DKIM config for {0}: Enabled={1}, Status={2}" -f $cfg.Name, $cfg.Enabled, $cfg.Status)
            if (-not $cfg.Enabled -or $cfg.Status -ne 'Valid') {
                $nonCompliant += $cfg.Name
            }
        }

        if ($nonCompliant.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "All DKIM configurations are enabled and valid."
        } else {
            $Status = 'FAIL'
            Add-Detail "DKIM is not fully enabled/valid for the following domains:"
            foreach ($n in $nonCompliant) { Add-Detail ("  - {0}" -f $n) }
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
