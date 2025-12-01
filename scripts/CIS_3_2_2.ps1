param(
    [switch]$UseExistingSession
)

$ControlId   = "3.2.2"
$ControlName = "Ensure DLP policies are enabled for Microsoft Teams"

function Write-CheckResult {
    param(
        [string]$Status,
        [string]$Message
    )
    Write-Output ("{0}`t{1}`t{2}`t{3}" -f $ControlId, $ControlName, $Status, $Message)
}

try {
    if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
        throw "ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if (-not $UseExistingSession) {
        Connect-IPPSSession -ErrorAction Stop | Out-Null
    }

    $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction Stop

    if (-not $dlpPolicies) {
        Write-CheckResult -Status "FAIL" -Message "No DLP policies found in tenant."
        return
    }

    # Only policies that include Teams as a workload
    $teamsPolicies = $dlpPolicies | Where-Object { $_.Workload -match "Teams" }

    if (-not $teamsPolicies -or $teamsPolicies.Count -eq 0) {
        Write-CheckResult -Status "FAIL" -Message "No DLP policies found with Workload including 'Teams'."
        return
    }

    # PASS criteria (approximation of CIS guidance):
    # - At least one policy with Mode = Enable
    # - TeamsLocation includes 'All'
    # - TeamsLocationException empty or null
    $compliant = $teamsPolicies | Where-Object {
        $_.Mode -eq "Enable" -and
        ($_.TeamsLocation -contains "All" -or $_.TeamsLocation -eq "All") -and
        (
            -not $_.TeamsLocationException -or
            ($_.TeamsLocationException | Where-Object { $_ -and $_.Trim() -ne "" }).Count -eq 0
        )
    }

    if ($compliant -and $compliant.Count -gt 0) {
        $names = ($compliant | Select-Object -ExpandProperty Name) -join ", "
        Write-CheckResult -Status "PASS" -Message ("{0} Teams DLP policy(ies) enabled with TeamsLocation='All' and no exceptions: {1}" -f $compliant.Count, $names)
    }
    else {
        $detail = $teamsPolicies | Select-Object Name, Mode, TeamsLocation, TeamsLocationException | Format-Table -AutoSize | Out-String
        Write-CheckResult -Status "FAIL" -Message ("Teams DLP policies found but none meet CIS criteria (Mode='Enable', TeamsLocation='All', no exceptions). Details:`n{0}" -f $detail.Trim())
    }
}
catch {
    Write-CheckResult -Status "ERROR" -Message ("Using existing IPPSSession… " + $_.Exception.Message)
}
