<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.6 (L1) – Ensure Exchange Online Spam Policies are set to notify administrators

DESCRIPTION
Checks the default outbound spam policy to ensure administrators are notified
when outbound spam is detected, and that BCC copies are sent to a monitored mailbox.

PASS:
 - BccSuspiciousOutboundMail = True
 - BccSuspiciousOutboundAdditionalRecipients has at least one address
 - NotifyOutboundSpam = True
 - NotifyOutboundSpamRecipients has at least one address

FAIL:
 - Any of the above is not true or empty.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.6'
$ControlTitle = 'Ensure Exchange Online Spam Policies are set to notify administrators'

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

    # Check cmdlet availability explicitly
    if (-not (Get-Command Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue)) {
        $Status = 'ERROR'
        Add-Detail "Get-HostedOutboundSpamFilterPolicy cmdlet is not available in this session."
        Add-Detail "Ensure you are connected to Exchange Online with a recent ExchangeOnlineManagement module."
        throw "Get-HostedOutboundSpamFilterPolicy not found."
    }

    $connectedHere = $false
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        Add-Detail "Using existing Exchange Online session."
    } catch {
        Add-Detail "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connectedHere = $true
    }

    # Get all outbound spam policies
    $policies = Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop

    if (-not $policies -or $policies.Count -eq 0) {
        $Status = 'FAIL'
        Add-Detail "No HostedOutboundSpamFilterPolicy objects found. At least the Default policy should exist."
    }
    else {
        # Prefer Default policy if present
        $policy = $policies | Where-Object { $_.Identity -eq 'Default' } | Select-Object -First 1
        if (-not $policy) {
            $policy = $policies | Select-Object -First 1
            Add-Detail ("Default policy not found; evaluating first available HostedOutboundSpamFilterPolicy '{0}'." -f $policy.Identity)
        }
        else {
            Add-Detail ("Evaluating HostedOutboundSpamFilterPolicy 'Default' (Identity='{0}')." -f $policy.Identity)
        }

        $bccFlag      = $policy.BccSuspiciousOutboundMail
        $bccRcpts     = $policy.BccSuspiciousOutboundAdditionalRecipients
        $notifyFlag   = $policy.NotifyOutboundSpam
        $notifyRcpts  = $policy.NotifyOutboundSpamRecipients

        Add-Detail ("  BccSuspiciousOutboundMail                = {0}" -f $bccFlag)
        Add-Detail ("  BccSuspiciousOutboundAdditionalRecipients = {0}" -f ($bccRcpts -join ', '))
        Add-Detail ("  NotifyOutboundSpam                       = {0}" -f $notifyFlag)
        Add-Detail ("  NotifyOutboundSpamRecipients             = {0}" -f ($notifyRcpts -join ', '))

        $hasBccRcpts    = $bccRcpts   -and $bccRcpts.Count   -gt 0
        $hasNotifyRcpts = $notifyRcpts -and $notifyRcpts.Count -gt 0

        if ($bccFlag -and $notifyFlag -and $hasBccRcpts -and $hasNotifyRcpts) {
            $Status = 'PASS'
            Add-Detail "Outbound spam notifications and BCC copies are configured with recipient addresses."
        }
        else {
            $Status = 'FAIL'
            Add-Detail "Outbound spam notification settings are NOT fully compliant with CIS 2.1.6."
            if (-not $bccFlag)        { Add-Detail "  - BccSuspiciousOutboundMail is not True." }
            if (-not $hasBccRcpts)    { Add-Detail "  - BccSuspiciousOutboundAdditionalRecipients is empty." }
            if (-not $notifyFlag)     { Add-Detail "  - NotifyOutboundSpam is not True." }
            if (-not $hasNotifyRcpts) { Add-Detail "  - NotifyOutboundSpamRecipients is empty." }

            Add-Detail ""
            Add-Detail "Example remediation (PowerShell):"
            Add-Detail "  \$BccEmailAddress    = @('secops@example.com')"
            Add-Detail "  \$NotifyEmailAddress = @('secops@example.com')"
            Add-Detail "  Set-HostedOutboundSpamFilterPolicy -Identity 'Default' `"
            Add-Detail "      -BccSuspiciousOutboundAdditionalRecipients \$BccEmailAddress `"
            Add-Detail "      -BccSuspiciousOutboundMail \$true `"
            Add-Detail "      -NotifyOutboundSpam \$true `"
            Add-Detail "      -NotifyOutboundSpamRecipients \$NotifyEmailAddress"
        }
    }

} catch {
    if ($Status -ne 'FAIL') { $Status = 'ERROR' }
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
