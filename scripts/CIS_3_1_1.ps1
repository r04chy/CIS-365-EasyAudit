param(
    [switch]$UseExistingSession
)

$ControlId   = "3.1.1"
$ControlName = "Ensure Microsoft 365 audit log search is Enabled"

function Write-CheckResult {
    param(
        [string]$Status,
        [string]$Message
    )
    # Format: ControlId<TAB>Title<TAB>Status<TAB>Message
    Write-Output ("{0}`t{1}`t{2}`t{3}" -f $ControlId, $ControlName, $Status, $Message)
}

try {
    if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
        throw "ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if (-not $UseExistingSession) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
    }

    $config = Get-AdminAuditLogConfig -ErrorAction Stop

    if ($config.UnifiedAuditLogIngestionEnabled -eq $true) {
        Write-CheckResult -Status "PASS" -Message "UnifiedAuditLogIngestionEnabled is True (audit log search is enabled)."
    }
    else {
        Write-CheckResult -Status "FAIL" -Message "UnifiedAuditLogIngestionEnabled is $($config.UnifiedAuditLogIngestionEnabled). Expected True."
    }
}
catch {
    Write-CheckResult -Status "ERROR" -Message ("Using existing Exchange Online session… " + $_.Exception.Message)
}
