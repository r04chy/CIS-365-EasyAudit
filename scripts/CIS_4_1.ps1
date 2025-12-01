param(
    [switch]$UseExistingSession
)

$ControlId   = "4.1"
$ControlName = "Ensure devices without a compliance policy are marked 'not compliant'"

function Write-CheckResult {
    param(
        [string]$Status,
        [string]$Message
    )
    # Format: ControlId<TAB>Title<TAB>Status<TAB>Message
    Write-Output ("{0}`t{1}`t{2}`t{3}" -f $ControlId, $ControlName, $Status, $Message)
}

try {
    # Ensure Graph module is available
    if (-not (Get-Module Microsoft.Graph.Authentication -ListAvailable)) {
        throw "Microsoft.Graph modules are not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Connect to Graph if requested / needed
    if (-not $UseExistingSession) {
        Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" -ErrorAction Stop | Out-Null
    }

    $uri      = "https://graph.microsoft.com/v1.0/deviceManagement/settings"
    $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop

    if ($null -eq $response) {
        throw "No response from Graph API at $uri."
    }

    $secureByDefault = $response.secureByDefault

    if ($null -eq $secureByDefault) {
        Write-CheckResult -Status "ERROR" -Message "secureByDefault property not found in deviceManagement/settings response."
        return
    }

    if ($secureByDefault -eq $true) {
        Write-CheckResult -Status "PASS" -Message "secureByDefault is True – devices with no compliance policy are marked 'Not compliant'."
    }
    else {
        Write-CheckResult -Status "FAIL" -Message ("secureByDefault is {0} – expected True so devices without a compliance policy are 'Not compliant'." -f $secureByDefault)
    }
}
catch {
    Write-CheckResult -Status "ERROR" -Message ("Using existing Microsoft Graph session… " + $_.Exception.Message)
}
