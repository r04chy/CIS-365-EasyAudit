<#
.SYNOPSIS
CIS Microsoft 365 Foundations v6.0.0
2.1.7 (L2) – Ensure that an anti-phishing policy has been created

.DESCRIPTION
Audits Microsoft 365 Defender Anti-Phishing configuration to ensure that an
anti-phishing policy exists and is configured per CIS 2.1.7 guidance.

PASS (simplified as per CIS intent):
 - At least one enabled AntiPhish policy with:
    * PhishThresholdLevel >= 3
    * EnableTargetedUserProtection = True
    * EnableOrganizationDomainsProtection = True
    * EnableMailboxIntelligence = True
    * EnableMailboxIntelligenceProtection = True
    * EnableSpoofIntelligence = True
    * TargetedUserProtectionAction = 'Quarantine'
    * TargetedDomainProtectionAction = 'Quarantine'
    * MailboxIntelligenceProtectionAction = 'Quarantine'
    * EnableFirstContactSafetyTips = True
    * EnableSimilarUsersSafetyTips = True
    * EnableSimilarDomainsSafetyTips = True
    * EnableUnusualCharactersSafetyTips = True
    * HonorDmarcPolicy = True
    * TargetedUsersToProtect not empty
 - And at least one enabled AntiPhish rule using that policy with a non-empty
   scope (RecipientDomainIs or SentToMemberOf).

FAIL:
 - No such policy/rule combination is found.

Sets:
 - $global:CISCheckResult = PASS | FAIL | ERROR
#>

[CmdletBinding()]
param()

$ControlId    = '2.1.7'
$ControlTitle = 'Ensure that an anti-phishing policy has been created'

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
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        Add-Detail "Using existing Exchange Online session."
    }
    catch {
        Add-Detail "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connectedHere = $true
    }

    $policies = Get-AntiPhishPolicy -ErrorAction Stop
    if (-not $policies) {
        $Status = 'FAIL'
        Add-Detail "No AntiPhish policies found."
    }
    else {
        $compliantPolicy = $null

        foreach ($p in $policies) {
            if (-not $p.Enabled) { continue }

            Add-Detail ("Evaluating AntiPhish policy: {0}" -f $p.Name)

            $targetUsersNonEmpty = $p.TargetedUsersToProtect -and $p.TargetedUsersToProtect.Count -gt 0

            $ok =
                $p.PhishThresholdLevel -ge 3 -and
                $p.EnableTargetedUserProtection -and
                $p.EnableOrganizationDomainsProtection -and
                $p.EnableMailboxIntelligence -and
                $p.EnableMailboxIntelligenceProtection -and
                $p.EnableSpoofIntelligence -and
                $p.TargetedUserProtectionAction -eq 'Quarantine' -and
                $p.TargetedDomainProtectionAction -eq 'Quarantine' -and
                $p.MailboxIntelligenceProtectionAction -eq 'Quarantine' -and
                $p.EnableFirstContactSafetyTips -and
                $p.EnableSimilarUsersSafetyTips -and
                $p.EnableSimilarDomainsSafetyTips -and
                $p.EnableUnusualCharactersSafetyTips -and
                $p.HonorDmarcPolicy -and
                $targetUsersNonEmpty

            if ($ok) {
                $compliantPolicy = $p
                Add-Detail "Found AntiPhish policy meeting CIS 2.1.7 settings: $($p.Name)"
                break
            }
        }

        if (-not $compliantPolicy) {
            $Status = 'FAIL'
            Add-Detail "No AntiPhish policy found that meets all CIS 2.1.7 criteria."
        }
        else {
            $rules = Get-AntiPhishRule -ErrorAction Stop |
                     Where-Object { $_.AntiPhishPolicy -eq $compliantPolicy.Name -and $_.State -ne 'Disabled' }

            if (-not $rules -or $rules.Count -eq 0) {
                $Status = 'FAIL'
                Add-Detail ("No enabled AntiPhishRule found referencing policy '{0}'." -f $compliantPolicy.Name)
            }
            else {
                $hasScope = $false
                foreach ($r in $rules) {
                    Add-Detail ("Found AntiPhishRule: Name='{0}', Priority={1}, RecipientDomainIs='{2}', SentToMemberOf='{3}'" -f `
                        $r.Name, $r.Priority, ($r.RecipientDomainIs -join ','), ($r.SentToMemberOf -join ','))

                    if ( ($r.RecipientDomainIs -and $r.RecipientDomainIs.Count -gt 0) -or
                         ($r.SentToMemberOf   -and $r.SentToMemberOf.Count   -gt 0) ) {
                        $hasScope = $true
                    }
                }

                if ($hasScope) {
                    $Status = 'PASS'
                    Add-Detail "At least one enabled AntiPhish rule using the compliant policy is scoped to users/groups/domains."
                }
                else {
                    $Status = 'FAIL'
                    Add-Detail "No AntiPhishRule with the compliant policy appears to scope any recipients."
                }
            }
        }
    }

}
catch {
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
