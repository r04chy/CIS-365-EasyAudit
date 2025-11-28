<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.14 (L1) – Ensure inbound anti-spam policies do not contain allowed domains

DESCRIPTION
Checks HostedContentFilterPolicy objects to ensure AllowedSenderDomains is empty.

PASS:
 - No HostedContentFilterPolicy has AllowedSenderDomains entries.

FAIL:
 - Any HostedContentFilterPolicy has one or more AllowedSenderDomains.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.14'
$ControlTitle = 'Ensure inbound anti-spam policies do not contain allowed domains'

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

    $policies = Get-HostedContentFilterPolicy -ErrorAction Stop
    if (-not $policies) {
        $Status = 'FAIL'
        Add-Detail "No HostedContentFilterPolicy objects found."
    } else {
        $violations = @()

        foreach ($p in $policies) {
            $allowed = $p.AllowedSenderDomains
            Add-Detail ("Policy '{0}': AllowedSenderDomains = {1}" -f $p.Identity, ($allowed -join ', '))
            if ($allowed -and $allowed.Count -gt 0) {
                $violations += $p.Identity
            }
        }

        if ($violations.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "No inbound anti-spam policies define AllowedSenderDomains."
        } else {
            $Status = 'FAIL'
            Add-Detail "The following inbound anti-spam policies define AllowedSenderDomains:"
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
