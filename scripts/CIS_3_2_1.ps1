param(
    [switch]$UseExistingSession
)

$ControlId   = "3.2.1"
$ControlName = "Ensure DLP policies are enabled"

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
        # Connect to the Purview / Security & Compliance endpoint
        Connect-IPPSSession -ErrorAction Stop | Out-Null
    }

    $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction Stop

    if (-not $dlpPolicies) {
        Write-CheckResult -Status "FAIL" -Message "No DLP policies found in tenant."
        return
    }

    # Consider “enabled” policies those with Mode = Enable
    $enabledPolicies = $dlpPolicies | Where-Object { $_.Mode -eq "Enable" }

    if ($enabledPolicies -and $enabledPolicies.Count -gt 0) {
        $names = ($enabledPolicies | Select-Object -ExpandProperty Name) -join ", "
        Write-CheckResult -Status "PASS" -Message ("{0} DLP policy(ies) enabled: {1}" -f $enabledPolicies.Count, $names)
    }
    else {
        $namesAll = ($dlpPolicies | Select-Object -ExpandProperty Name) -join ", "
        Write-CheckResult -Status "FAIL" -Message ("DLP policies exist but none are in Mode='Enable'. Policies: {0}" -f $namesAll)
    }
}
catch {
    Write-CheckResult -Status "ERROR" -Message ("Using existing IPPSSession… " + $_.Exception.Message)
}
