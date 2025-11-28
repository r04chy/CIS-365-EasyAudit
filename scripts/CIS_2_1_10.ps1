<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.10 (L1) – Ensure DMARC Records for all Exchange Online domains are published

DESCRIPTION
Checks DMARC TXT records for each accepted domain at _dmarc.domain.

PASS:
 - DMARC TXT record exists and contains:
    * v=DMARC1
    * p=quarantine or p=reject
    * pct=100
    * rua=mailto:...
    * ruf=mailto:...

FAIL:
 - Missing or incomplete DMARC record.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.10'
$ControlTitle = 'Ensure DMARC Records for all Exchange Online domains are published'

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
        Add-Detail "No accepted domains found."
    } else {
        $nonCompliant = @()

        foreach ($d in $domains) {
            $dmarcHost = "_dmarc.$d"
            Add-Detail ("Checking DMARC record for {0}" -f $dmarcHost)
            try {
                $txtRecords = Resolve-DnsName -Name $dmarcHost -Type TXT -ErrorAction Stop
                $txtStrings = $txtRecords | Where-Object { $_.Type -eq 'TXT' } | Select-Object -ExpandProperty Strings

                $ok = $false
                foreach ($s in $txtStrings) {
                    $lower = $s.ToLowerInvariant()
                    $hasV   = $lower -match 'v=dmarc1'
                    $hasP   = $lower -match 'p=quarantine' -or $lower -match 'p=reject'
                    $hasPct = $lower -match 'pct=100'
                    $hasRua = $lower -match 'rua=mailto:'
                    $hasRuf = $lower -match 'ruf=mailto:'

                    if ($hasV -and $hasP -and $hasPct -and $hasRua -and $hasRuf) {
                        $ok = $true
                        break
                    }
                }

                if ($ok) {
                    Add-Detail ("  DMARC record OK for {0}" -f $dmarcHost)
                } else {
                    Add-Detail ("  DMARC record missing or incomplete for {0}" -f $dmarcHost)
                    $nonCompliant += $d
                }

            } catch {
                Add-Detail ("  Error resolving DMARC for {0}: {1}" -f $dmarcHost, $_.Exception.Message)
                $nonCompliant += $d
            }
        }

        if ($nonCompliant.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "All accepted domains have DMARC records with required flags."
        } else {
            $Status = 'FAIL'
            Add-Detail "The following domains are missing compliant DMARC records:"
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
