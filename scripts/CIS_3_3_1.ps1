param(
    [switch]$UseExistingSession
)

$ControlId   = "3.3.1"
$ControlName = "Ensure Information Protection sensitivity label policies are published"

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

    $policies = Get-LabelPolicy -WarningAction Ignore -ErrorAction Stop |
        Where-Object { $_.Type -eq "PublishedSensitivityLabel" }

    if ($policies -and $policies.Count -gt 0) {
        $names = ($policies | Select-Object -ExpandProperty Name) -join ", "
        Write-CheckResult -Status "PASS" -Message ("{0} published Sensitivity Label policy(ies) found: {1}" -f $policies.Count, $names)
    }
    else {
        Write-CheckResult -Status "FAIL" -Message "No published Sensitivity Label policies found (Type='PublishedSensitivityLabel')."
    }
}
catch {
    Write-CheckResult -Status "ERROR" -Message ("Using existing IPPSSession… " + $_.Exception.Message)
}
