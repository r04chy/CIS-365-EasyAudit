<#
.SYNOPSIS
CIS Microsoft 365 Foundations Benchmark v6.0.0
Control 2.1.5 (L2) – Ensure Safe Attachments for SharePoint, OneDrive,
and Microsoft Teams is Enabled

.DESCRIPTION
Audits the tenant-wide ATP policy to ensure Safe Attachments for
SharePoint, OneDrive, and Teams is enabled, and Safe Documents is
configured per CIS 2.1.5.

PASS:
 - EnableATPForSPOTeamsODB = True
 - EnableSafeDocs          = True
 - AllowSafeDocsOpen       = False

FAIL:
 - Any of the above does not match.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.5'
$ControlTitle = 'Ensure Safe Attachments for SharePoint, OneDrive, and Microsoft Teams is Enabled'

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

    $atpPolicies = Get-AtpPolicyForO365 -ErrorAction Stop

    if (-not $atpPolicies -or $atpPolicies.Count -eq 0) {
        $Status = 'FAIL'
        Add-Detail "No ATP policies for Office 365 (Get-AtpPolicyForO365) were found."
    }
    else {
        # Typically only one policy exists; if multiple, evaluate the first
        $policy = $atpPolicies | Select-Object -First 1
        Add-Detail ("Evaluating ATP policy for O365: {0}" -f $policy.Name)

        $enableSpoTeams   = $null
        $enableSafeDocs   = $null
        $allowSafeDocs    = $null

        if ($policy.PSObject.Properties.Name -contains 'EnableATPForSPOTeamsODB') {
            $enableSpoTeams = [bool]$policy.EnableATPForSPOTeamsODB
        }
        if ($policy.PSObject.Properties.Name -contains 'EnableSafeDocs') {
            $enableSafeDocs = [bool]$policy.EnableSafeDocs
        }
        if ($policy.PSObject.Properties.Name -contains 'AllowSafeDocsOpen') {
            $allowSafeDocs = [bool]$policy.AllowSafeDocsOpen
        }

        Add-Detail ("  EnableATPForSPOTeamsODB = {0}" -f $enableSpoTeams)
        Add-Detail ("  EnableSafeDocs          = {0}" -f $enableSafeDocs)
        Add-Detail ("  AllowSafeDocsOpen       = {0}" -f $allowSafeDocs)

        $expectedSpoTeams = $true
        $expectedSafeDocs = $true
        $expectedAllow    = $false

        $mismatches = @()

        if ($enableSpoTeams -ne $expectedSpoTeams) {
            $mismatches += ("EnableATPForSPOTeamsODB: expected {0}, actual {1}" -f $expectedSpoTeams, $enableSpoTeams)
        }
        if ($enableSafeDocs -ne $expectedSafeDocs) {
            $mismatches += ("EnableSafeDocs: expected {0}, actual {1}" -f $expectedSafeDocs, $enableSafeDocs)
        }
        if ($allowSafeDocs -ne $expectedAllow) {
            $mismatches += ("AllowSafeDocsOpen: expected {0}, actual {1}" -f $expectedAllow, $allowSafeDocs)
        }

        if ($mismatches.Count -eq 0) {
            $Status = 'PASS'
            Add-Detail "ATP policy for O365 matches CIS 2.1.5 recommended values."
        }
        else {
            $Status = 'FAIL'
            Add-Detail "ATP policy for O365 does NOT match all CIS 2.1.5 recommended values."
            Add-Detail "Mismatched settings:"
            foreach ($m in $mismatches) {
                Add-Detail ("  - {0}" -f $m)
            }

            Add-Detail ""
            Add-Detail "Example remediation (PowerShell):"
            Add-Detail ("  Set-AtpPolicyForO365 -Identity '{0}' -EnableATPForSPOTeamsODB $true -EnableSafeDocs $true -AllowSafeDocsOpen $false" -f $policy.Name)
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

$global:CISCheckResult = $Status

[PSCustomObject]@{
    Control = $ControlId
    Title   = $ControlTitle
    Status  = $Status
    Details = $Details -join "`n"
}
