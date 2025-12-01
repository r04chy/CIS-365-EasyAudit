param(
    [switch]$UseExistingSession
)

$ControlId   = "4.2"
$ControlName = "Ensure device enrollment for personally owned devices is blocked by default"

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

    $uri      = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations"
    $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop

    if (-not $response -or -not $response.value) {
        throw "No deviceEnrollmentConfigurations returned from Graph API at $uri."
    }

    $config = $response.value | Where-Object { $_.id -match 'DefaultPlatformRestrictions' -and $_.priority -eq 0 }

    if (-not $config) {
        throw "DefaultPlatformRestrictions configuration with priority 0 was not found."
    }

    # Build a map of platforms we care about
    $platforms = @(
        @{ Name = "Windows";         Restriction = $config.windowsRestriction       },
        @{ Name = "iOS";             Restriction = $config.iosRestriction           },
        @{ Name = "AndroidForWork";  Restriction = $config.androidForWorkRestriction },
        @{ Name = "macOS";           Restriction = $config.macOSRestriction         },
        @{ Name = "Android";         Restriction = $config.androidRestriction       }
    )

    $nonCompliant = @()
    $details      = @()

    foreach ($p in $platforms) {
        $name        = $p.Name
        $restriction = $p.Restriction

        if (-not $restriction) {
            $nonCompliant += "$name (no restriction config found)"
            $details      += ("{0}: no restriction object present" -f $name)
            continue
        }

        $personalBlocked = $restriction.personalDeviceEnrollmentBlocked
        # platformBlocked is not in the sample script but is mentioned in the benchmark notes
        $platformBlocked = $restriction.platformBlocked

        $isCompliant = ($personalBlocked -eq $true) -or ($platformBlocked -eq $true)

        $details += ("{0}: personalDeviceEnrollmentBlocked={1}, platformBlocked={2}" -f $name, $personalBlocked, $platformBlocked)

        if (-not $isCompliant) {
            $nonCompliant += ("{0} (personalDeviceEnrollmentBlocked={1}, platformBlocked={2})" -f $name, $personalBlocked, $platformBlocked)
        }
    }

    if ($nonCompliant.Count -eq 0) {
        Write-CheckResult -Status "PASS" -Message ("All platforms block personally owned device enrollment (or platform is fully blocked). Details: {0}" -f ($details -join "; "))
    }
    else {
        Write-CheckResult -Status "FAIL" -Message ("One or more platforms do NOT block personally owned device enrollment by default. Non-compliant: {0}. Details: {1}" -f ($nonCompliant -join ", "), ($details -join "; "))
    }
}
catch {
    Write-CheckResult -Status "ERROR" -Message ("Using existing Microsoft Graph session… " + $_.Exception.Message)
}
