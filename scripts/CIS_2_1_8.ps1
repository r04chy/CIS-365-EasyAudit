<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.8 (L1) – Ensure that SPF records are published for all Exchange Domains

DESCRIPTION
Checks all accepted domains in Exchange Online for SPF TXT records that contain:
 - v=spf1
 - include:spf.protection.outlook.com

PASS:
 - All accepted domains have an SPF TXT record including the above.

FAIL:
 - Any accepted domain is missing such a record or lookup fails.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.8'
$ControlTitle = 'Ensure that SPF records are published for all Exchange Domains'

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

    $domains = Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName
    if (-not $domains -or $domains.Count -eq 0) {
        $Status = 'FAIL'
        Add-Detail "No accepted domains found via Get-AcceptedDomain."
    } else {
        $nonCompliant = @()

        foreach ($d in $domains) {
            Add-Detail ("Checking SPF record for domain: {0}" -f $d)
            try {
                $txtRecords = Resolve-DnsName -Name $d -Type TXT -ErrorAction Stop
                $txtStrings = $txtRecords | Where-Object { $_.Type -eq 'TXT' } | Select-Object -ExpandProperty Strings

                $hasSpf = $false
                foreach ($s in $txtStrings) {
                    if ($s -match 'v=spf1' -and $s -match 'include:spf\.protection\.outlook\.com') {
                        $hasSpf = $true
                        break
                    }
                }

                if ($hasSpf) {
                    Add-Detail ("  SPF record OK for {0}" -f $d)
                } else {
                    Add-Detail ("  SPF record missing or incorrect for {0}" -f $d)
                    $nonCompliant += $d
                }
            } catch {
                Add-Detail ("  Error resolving TXT for {0}: {1}" -f $d, $_.Exception.Message)
                $nonCompliant += $d
            }
        }

        if ($nonCompliant.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "All accepted domains have SPF records including 'v=spf1' and 'include:spf.protection.outlook.com'."
        } else {
            $Status = 'FAIL'
            Add-Detail "The following domains are missing the required SPF configuration:"
            foreach ($nd in $nonCompliant) { Add-Detail ("  - {0}" -f $nd) }
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
