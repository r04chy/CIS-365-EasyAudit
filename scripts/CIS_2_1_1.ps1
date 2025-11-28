<#
.SYNOPSIS
CIS Microsoft 365 Foundations Benchmark v6.0.0
Control 2.1.1 (L2) – Ensure Safe Links for Office Applications is Enabled

.DESCRIPTION
Audits Microsoft 365 Defender Safe Links configuration for Office applications
(Email, Teams, Office 365 apps) according to CIS 2.1.1.

Logic:

1. Get all enabled Safe Links rules, ordered by Priority.
2. Take the highest-priority enabled rule.
3. Get its associated Safe Links policy.
4. Check the following properties match CIS recommendations:

   EnableSafeLinksForEmail   = $true
   EnableSafeLinksForTeams   = $true
   EnableSafeLinksForOffice  = $true
   TrackClicks               = $true
   AllowClickThrough         = $false
   ScanUrls                  = $true
   EnableForInternalSenders  = $true
   DeliverMessageAfterScan   = $true
   DisableUrlRewrite         = $false

PASS:
 - Highest-priority active policy meets all recommended values.

FAIL:
 - No enabled Safe Links rules, or
 - Policy does not match all recommended values.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR

.REQUIREMENTS
 - ExchangeOnlineManagement module
 - Permission to read Safe Links policies and rules

#>

[CmdletBinding()]
param()

# ------------------ Metadata ------------------
$ControlId    = '2.1.1'
$ControlTitle = 'Ensure Safe Links for Office Applications is Enabled'

# ------------------ Output state ------------------
$Status  = 'ERROR'
$Details = @()

function Add-Detail {
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )
    $script:Details += $Text
}

try {
    # Ensure module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    # Connect to Exchange Online if not already connected
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

    # ------------------ Get Safe Links rules ------------------
    $rules = Get-SafeLinksRule -ErrorAction Stop |
             Where-Object { $_.State -ne 'Disabled' } |
             Sort-Object Priority

    if (-not $rules -or $rules.Count -eq 0) {
        $Status = 'FAIL'
        Add-Detail "No enabled Safe Links rules found. At least one enabled Safe Links rule/policy is required."
    }
    else {
        $primaryRule = $rules | Select-Object -First 1
        Add-Detail ("Highest-priority enabled Safe Links rule: Name = '{0}', Priority = {1}, SafeLinksPolicy = '{2}'" -f `
            $primaryRule.Name, $primaryRule.Priority, $primaryRule.SafeLinksPolicy)

        # ------------------ Get associated policy ------------------
        $policy = Get-SafeLinksPolicy -Identity $primaryRule.SafeLinksPolicy -ErrorAction Stop

        $actual = [ordered]@{}

        $propertiesToCheck = @(
            'EnableSafeLinksForEmail',
            'EnableSafeLinksForTeams',
            'EnableSafeLinksForOffice',
            'TrackClicks',
            'AllowClickThrough',
            'ScanUrls',
            'EnableForInternalSenders',
            'DeliverMessageAfterScan',
            'DisableUrlRewrite'
        )

        foreach ($prop in $propertiesToCheck) {
            if ($policy.PSObject.Properties.Name -contains $prop) {
                $actual[$prop] = [bool]$policy.$prop
            }
            else {
                $actual[$prop] = $null
                Add-Detail ("Warning: Policy '{0}' does not expose property '{1}'. Treating as non-compliant." -f $policy.Identity, $prop)
            }
        }

        Add-Detail ("Evaluating Safe Links policy values for policy: {0}" -f $policy.Identity)
        foreach ($k in $actual.Keys) {
            $val = $actual[$k]
            if ($null -eq $val) {
                Add-Detail ("  {0} = <null>" -f $k)
            }
            else {
                Add-Detail ("  {0} = {1}" -f $k, $val)
            }
        }

        # ------------------ Expected values per CIS ------------------
        $expected = @{
            EnableSafeLinksForEmail   = $true
            EnableSafeLinksForTeams   = $true
            EnableSafeLinksForOffice  = $true
            TrackClicks               = $true
            AllowClickThrough         = $false
            ScanUrls                  = $true
            EnableForInternalSenders  = $true
            DeliverMessageAfterScan   = $true
            DisableUrlRewrite         = $false
        }

        $mismatches = @()

        foreach ($key in $expected.Keys) {
            $expectedValue = $expected[$key]
            $actualValue   = $actual[$key]

            if ($null -eq $actualValue) {
                $mismatches += ("{0}: expected {1}, actual <missing property>" -f $key, $expectedValue)
            }
            elseif ($actualValue -ne $expectedValue) {
                $mismatches += ("{0}: expected {1}, actual {2}" -f $key, $expectedValue, $actualValue)
            }
        }

        if ($mismatches.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "Safe Links policy associated with the highest-priority enabled rule matches all CIS 2.1.1 recommended values."
        }
        else {
            $Status = 'FAIL'
            Add-Detail "Safe Links policy associated with the highest-priority enabled rule does NOT match all CIS 2.1.1 recommended values."
            Add-Detail "Mismatched settings:"
            foreach ($m in $mismatches) {
                Add-Detail ("  - {0}" -f $m)
            }

            Add-Detail ""
            Add-Detail "Example remediation (PowerShell) – adjust Identity as appropriate:"
            Add-Detail ("  Set-SafeLinksPolicy -Identity '{0}' " -f $policy.Identity)
            Add-Detail "      -EnableSafeLinksForEmail `$true"
            Add-Detail "      -EnableSafeLinksForTeams `$true"
            Add-Detail "      -EnableSafeLinksForOffice `$true"
            Add-Detail "      -TrackClicks `$true"
            Add-Detail "      -AllowClickThrough `$false"
            Add-Detail "      -ScanUrls `$true"
            Add-Detail "      -EnableForInternalSenders `$true"
            Add-Detail "      -DeliverMessageAfterScan `$true"
            Add-Detail "      -DisableUrlRewrite `$false"

            Add-Detail ""
            Add-Detail "Example: create dedicated CIS Safe Links policy + rule:"
            Add-Detail "  New-SafeLinksPolicy -Name 'CIS SafeLinks Policy' -EnableSafeLinksForEmail `$true -EnableSafeLinksForTeams `$true -EnableSafeLinksForOffice `$true -TrackClicks `$true -AllowClickThrough `$false -ScanUrls `$true -EnableForInternalSenders `$true -DeliverMessageAfterScan `$true -DisableUrlRewrite `$false"
            Add-Detail "  New-SafeLinksRule -Name 'CIS SafeLinks' -SafeLinksPolicy 'CIS SafeLinks Policy' -RecipientDomainIs (Get-AcceptedDomain).Name -Priority 0"
        }
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

# ------------------ Final output ------------------
$global:CISCheckResult = $Status

[PSCustomObject]@{
    Control = $ControlId
    Title   = $ControlTitle
    Status  = $Status
    Details = $Details -join "`n"
}
