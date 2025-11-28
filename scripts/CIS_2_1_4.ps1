<#
.SYNOPSIS
CIS Microsoft 365 Foundations Benchmark v6.0.0
Control 2.1.4 (L2) – Ensure Safe Attachments policy is enabled

.DESCRIPTION
Audits Safe Attachments configuration. Attempts to locate the effective Safe
Attachments policy by:

 1. Looking at enabled Safe Attachment rules ordered by Priority.
 2. Taking the highest-priority rule and resolving its SafeAttachmentPolicy.
 3. If no enabled rules exist, using the 'Built-In Protection Policy' if present,
    otherwise the first Safe Attachment policy.

PASS:
 - Enable   = True
 - Action   = 'Block'
 - QuarantineTag = 'AdminOnlyAccessPolicy'

FAIL:
 - Any of the above does not match.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.4'
$ControlTitle = 'Ensure Safe Attachments policy is enabled'

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

    $policy   = $null
    $policyId = $null

    # Try to derive policy from highest-priority Safe Attachment rule
    $saRules = Get-SafeAttachmentRule -ErrorAction SilentlyContinue |
               Where-Object { $_.State -ne 'Disabled' } |
               Sort-Object Priority

    if ($saRules -and $saRules.Count -gt 0) {
        $primaryRule = $saRules | Select-Object -First 1
        $policyId    = $primaryRule.SafeAttachmentPolicy
        Add-Detail ("Highest-priority enabled Safe Attachment rule: Name = '{0}', Priority = {1}, SafeAttachmentPolicy = '{2}'" -f `
            $primaryRule.Name, $primaryRule.Priority, $policyId)

        $policy = Get-SafeAttachmentPolicy -Identity $policyId -ErrorAction Stop
    }
    else {
        Add-Detail "No enabled Safe Attachment rules found; using 'Built-In Protection Policy' or first available Safe Attachment policy."

        try {
            $policy = Get-SafeAttachmentPolicy -Identity 'Built-In Protection Policy' -ErrorAction Stop
        }
        catch {
            $allPolicies = Get-SafeAttachmentPolicy -ErrorAction Stop
            if (-not $allPolicies -or $allPolicies.Count -eq 0) {
                throw "No Safe Attachment policies found in tenant."
            }
            $policy = $allPolicies | Select-Object -First 1
        }

        $policyId = $policy.Identity
    }

    Add-Detail ("Evaluating Safe Attachment policy: {0}" -f $policyId)

    $enable        = $null
    $action        = $null
    $quarantineTag = $null

    if ($policy.PSObject.Properties.Name -contains 'Enable') {
        $enable = [bool]$policy.Enable
    }
    if ($policy.PSObject.Properties.Name -contains 'Action') {
        $action = [string]$policy.Action
    }
    if ($policy.PSObject.Properties.Name -contains 'QuarantineTag') {
        $quarantineTag = [string]$policy.QuarantineTag
    }

    Add-Detail ("  Enable        = {0}" -f $enable)
    Add-Detail ("  Action        = {0}" -f $action)
    Add-Detail ("  QuarantineTag = {0}" -f $quarantineTag)

    $expectedEnable        = $true
    $expectedAction        = 'Block'
    $expectedQuarantineTag = 'AdminOnlyAccessPolicy'

    $mismatches = @()

    if ($enable -ne $expectedEnable) {
        $mismatches += ("Enable: expected {0}, actual {1}" -f $expectedEnable, $enable)
    }
    if ($action -ne $expectedAction) {
        $mismatches += ("Action: expected '{0}', actual '{1}'" -f $expectedAction, $action)
    }
    if ($quarantineTag -ne $expectedQuarantineTag) {
        $mismatches += ("QuarantineTag: expected '{0}', actual '{1}'" -f $expectedQuarantineTag, $quarantineTag)
    }

    if ($mismatches.Count -eq 0) {
        $Status = 'PASS'
        Add-Detail "Safe Attachments policy matches CIS 2.1.4 recommended values."
    }
    else {
        $Status = 'FAIL'
        Add-Detail "Safe Attachments policy does NOT match all CIS 2.1.4 recommended values."
        Add-Detail "Mismatched settings:"
        foreach ($m in $mismatches) {
            Add-Detail ("  - {0}" -f $m)
        }

        Add-Detail ""
        Add-Detail "Example remediation (PowerShell):"
        Add-Detail ("  Set-SafeAttachmentPolicy -Identity '{0}' -Enable $true -Action 'Block' -QuarantineTag 'AdminOnlyAccessPolicy'" -f $policyId)
        Add-Detail ""
        Add-Detail "Or create a dedicated CIS Safe Attachments policy + rule:"
        Add-Detail "  New-SafeAttachmentPolicy -Name 'CIS 2.1.4' -Enable $true -Action 'Block' -QuarantineTag 'AdminOnlyAccessPolicy'"
        Add-Detail "  New-SafeAttachmentRule -Name 'CIS 2.1.4 Rule' -SafeAttachmentPolicy 'CIS 2.1.4' -RecipientDomainIs 'exampledomain.com'"
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
