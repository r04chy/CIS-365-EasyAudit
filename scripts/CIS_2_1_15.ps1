<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.15 (L1) – Ensure outbound anti-spam message limits are in place

DESCRIPTION
Checks the Default HostedOutboundSpamFilterPolicy for outbound thresholds and actions.

PASS (CIS-aligned or stricter):
 - RecipientLimitExternalPerHour between 1 and 500
 - RecipientLimitInternalPerHour between 1 and 1000
 - RecipientLimitPerDay between 1 and 1000
 - ActionWhenThresholdReached = 'BlockUser'
 - NotifyOutboundSpamRecipients has at least one address

FAIL:
 - Any of the above conditions not met.
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.15'
$ControlTitle = 'Ensure outbound anti-spam message limits are in place'

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

    if (-not (Get-Command Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue)) {
        throw "Get-HostedOutboundSpamFilterPolicy cmdlet is not available in this session."
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

    $policy = Get-HostedOutboundSpamFilterPolicy -Identity 'Default' -ErrorAction Stop

    $extPerHour  = $policy.RecipientLimitExternalPerHour
    $intPerHour  = $policy.RecipientLimitInternalPerHour
    $perDay      = $policy.RecipientLimitPerDay
    $action      = $policy.ActionWhenThresholdReached
    $notifyRcpts = $policy.NotifyOutboundSpamRecipients

    Add-Detail "Default HostedOutboundSpamFilterPolicy values:"
    Add-Detail ("  RecipientLimitExternalPerHour = {0}" -f $extPerHour)
    Add-Detail ("  RecipientLimitInternalPerHour = {0}" -f $intPerHour)
    Add-Detail ("  RecipientLimitPerDay         = {0}" -f $perDay)
    Add-Detail ("  ActionWhenThresholdReached   = {0}" -f $action)
    Add-Detail ("  NotifyOutboundSpamRecipients = {0}" -f ($notifyRcpts -join ', '))

    $hasNotify = $notifyRcpts -and $notifyRcpts.Count -gt 0

    $okExt = ($extPerHour  -gt 0 -and $extPerHour  -le 500)
    $okInt = ($intPerHour  -gt 0 -and $intPerHour  -le 1000)
    $okDay = ($perDay      -gt 0 -and $perDay      -le 1000)
    $okAct = ($action -eq 'BlockUser')

    if ($okExt -and $okInt -and $okDay -and $okAct -and $hasNotify) {
        $Status = 'PASS'
        Add-Detail "Default outbound spam policy meets CIS 2.1.15 recommended or stricter limits."
    } else {
        $Status = 'FAIL'
        Add-Detail "Default outbound spam policy does NOT meet all CIS 2.1.15 requirements."
        if (-not $okExt) { Add-Detail ("  - RecipientLimitExternalPerHour should be between 1 and 500 (current: {0})" -f $extPerHour) }
        if (-not $okInt) { Add-Detail ("  - RecipientLimitInternalPerHour should be between 1 and 1000 (current: {0})" -f $intPerHour) }
        if (-not $okDay) { Add-Detail ("  - RecipientLimitPerDay should be between 1 and 1000 (current: {0})" -f $perDay) }
        if (-not $okAct) { Add-Detail ("  - ActionWhenThresholdReached should be 'BlockUser' (current: {0})" -f $action) }
        if (-not $hasNotify) { Add-Detail "  - NotifyOutboundSpamRecipients should contain at least one monitored mailbox." }

        Add-Detail ""
        Add-Detail "Example remediation (PowerShell):"
        Add-Detail "  `$params = @{"
        Add-Detail "      RecipientLimitExternalPerHour = 500"
        Add-Detail "      RecipientLimitInternalPerHour = 1000"
        Add-Detail "      RecipientLimitPerDay          = 1000"
        Add-Detail "      ActionWhenThresholdReached   = 'BlockUser'"
        Add-Detail "      NotifyOutboundSpamRecipients = @('admin@example.com')"
        Add-Detail "  }"
        Add-Detail "  Set-HostedOutboundSpamFilterPolicy -Identity 'Default' @params"
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
