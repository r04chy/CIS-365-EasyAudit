<#
.SYNOPSIS
CIS Microsoft 365 Foundations Benchmark v6.0.0
Control 1.3.9 (L1) – Ensure shared bookings pages are restricted to select users

.DESCRIPTION
Audits Microsoft Bookings / Shared Bookings configuration.

PASS if EITHER:
 - Shared Bookings mailbox creation is disabled on the default OWA policy
     (OwaMailboxPolicy-Default.BookingsMailboxCreationEnabled = False)
   OR
 - Bookings is disabled at the organization level
     (OrganizationConfig.BookingsEnabled = False)

If both are enabled, control FAILS.

The script:
 - Writes a structured object with Control, Title, Status, and Details
 - Sets $global:CISCheckResult = PASS | FAIL | ERROR

.REQUIREMENTS
 - ExchangeOnlineManagement module
 - Permission to read OrganizationConfig and OWA mailbox policy

#>

[CmdletBinding()]
param()

# ------------------ Metadata ------------------
$ControlId    = '1.3.9'
$ControlTitle = "Ensure shared bookings pages are restricted to select users"

# ------------------ Output object ------------------
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
        # Available in EXO v3+
        $null = Get-ConnectionInformation -ErrorAction Stop
        Add-Detail "Using existing Exchange Online session."
    } catch {
        Add-Detail "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connectedHere = $true
    }

    # --------- 1. OWA default policy: BookingsMailboxCreationEnabled ----------
    $owaPolicies = Get-OwaMailboxPolicy -ErrorAction Stop

    $defaultOwaPolicy = $owaPolicies |
        Where-Object { $_.IsDefault -eq $true } |
        Select-Object -First 1

    if (-not $defaultOwaPolicy) {
        # Fallback: classic default name
        $defaultOwaPolicy = $owaPolicies |
            Where-Object { $_.Name -eq 'OwaMailboxPolicy-Default' } |
            Select-Object -First 1
    }

    if (-not $defaultOwaPolicy) {
        throw "Unable to determine the default OWA mailbox policy."
    }

    if ($defaultOwaPolicy.PSObject.Properties.Name -contains 'BookingsMailboxCreationEnabled') {
        $bookingsCreationEnabled = [bool]$defaultOwaPolicy.BookingsMailboxCreationEnabled
        Add-Detail ("{0}.BookingsMailboxCreationEnabled = {1}" -f $defaultOwaPolicy.Name, $bookingsCreationEnabled)
    } else {
        $bookingsCreationEnabled = $true   # safest assumption
        Add-Detail ("{0}.BookingsMailboxCreationEnabled property not present; treating as Enabled (True) for safety." -f $defaultOwaPolicy.Name)
    }

    # --------- 2. Tenant-level BookingsEnabled -------------------------------
    $orgConfig = Get-OrganizationConfig -ErrorAction Stop

    if ($orgConfig.PSObject.Properties.Name -contains 'BookingsEnabled') {
        $bookingsEnabled = [bool]$orgConfig.BookingsEnabled
        Add-Detail ("OrganizationConfig.BookingsEnabled = {0}" -f $bookingsEnabled)
    } else {
        $bookingsEnabled = $true  # safest assumption
        Add-Detail "OrganizationConfig.BookingsEnabled not present; treating as Enabled (True) for safety."
    }

    # --------- Evaluation ----------------------------------------------------
    # PASS if:
    #   BookingsMailboxCreationEnabled = False on default OWA policy
    #      OR
    #   BookingsEnabled = False at org level
    if ($bookingsCreationEnabled -eq $false) {
        $Status = 'PASS'
        Add-Detail "Shared Bookings mailbox creation is disabled on the default OWA policy (BookingsMailboxCreationEnabled = False)."
    }
    elseif ($bookingsEnabled -eq $false) {
        $Status = 'PASS'
        Add-Detail "Bookings is disabled at the organization level (BookingsEnabled = False). This is a more restrictive, compliant state."
    }
    else {
        $Status = 'FAIL'
        Add-Detail "Bookings is enabled tenant-wide and Shared Bookings mailbox creation is enabled on the default OWA policy."
        Add-Detail "Remediation example:"
        Add-Detail '  Set-OwaMailboxPolicy "OwaMailboxPolicy-Default" -BookingsMailboxCreationEnabled:$false'
        Add-Detail "Optionally (more restrictive, not required for CIS compliance):"
        Add-Detail '  Set-OrganizationConfig -BookingsEnabled $false'
    }

} catch {
    $Status = 'ERROR'
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
    Control     = $ControlId
    Title       = $ControlTitle
    Status      = $Status
    Details     = $Details -join "`n"
}

