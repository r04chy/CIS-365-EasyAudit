<#
CIS Microsoft 365 Foundations v5.0.0
Control 1.1.4 (L1): Ensure administrative accounts use licenses with a reduced application footprint

This script:
  - Identifies admin accounts (privileged directory roles)
  - Retrieves their license assignments and enabled service plans
  - Flags whether they appear to have a "reduced application footprint"
  - Outputs:
      * Per-admin license + service plan details
      * A HasReducedFootprint flag
      * PASS/FAIL for CIS 1.1.4
  - Designed for both automation and human review

REQUIRES:
  - Microsoft Graph PowerShell SDK
      Install-Module Microsoft.Graph -Scope CurrentUser
  - Permissions (consented in your tenant):
      Directory.Read.All
      RoleManagement.Read.Directory
#>

param(
    [string]$ReportPath
)

# -----------------------------------------------------------
# 1. Connect to Microsoft Graph
# -----------------------------------------------------------
$scopes = @(
    "Directory.Read.All",
    "RoleManagement.Read.Directory"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes


# -----------------------------------------------------------
# 2. Define privileged / admin roles
#    (Same model as 1.1.1 / 1.1.2)
# -----------------------------------------------------------
$PrivilegedRoleNames = @(
    "Global Administrator",
    "Privileged Role Administrator",
    "Security Administrator",
    "Exchange Administrator",
    "SharePoint Administrator",
    "Teams Administrator",
    "User Administrator",
    "Authentication Administrator",
    "Application Administrator",
    "Cloud Application Administrator",
    "Compliance Administrator",
    "Helpdesk Administrator",
    "Intune Administrator",
    "Hybrid Identity Administrator",
    "Conditional Access Administrator"
)

Write-Host "Retrieving privileged roles..." -ForegroundColor Cyan

$allRoles   = Get-MgDirectoryRole -All
$adminRoles = $allRoles | Where-Object {
    $PrivilegedRoleNames -contains $_.DisplayName
}

if (-not $adminRoles) {
    Write-Warning "No privileged roles found. Exiting."
    $global:CISCheckResult = "FAIL"
    return
}

Write-Host "`nPrivileged roles detected:" -ForegroundColor Cyan
$adminRoles | Select-Object DisplayName, Id | Format-Table -AutoSize


# -----------------------------------------------------------
# 3. Collect admin users (user + roles)
# -----------------------------------------------------------
Write-Host "`nEnumerating admin accounts..." -ForegroundColor Cyan

$rawUserRoles = @()

foreach ($role in $adminRoles) {
    Write-Host "`nProcessing role: $($role.DisplayName)" -ForegroundColor Yellow

    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
    if (-not $members) {
        Write-Host "  No members assigned (PIM-eligible-only or empty)" -ForegroundColor DarkGray
        continue
    }

    foreach ($member in $members) {
        # Try to resolve as a user; skip non-user objects
        $user = $null
        try {
            $user = Get-MgUser -UserId $member.Id `
                -Property Id,DisplayName,UserPrincipalName,AccountEnabled `
                -ErrorAction Stop
        }
        catch {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
            continue
        }

        $rawUserRoles += [pscustomobject]@{
            UserId      = $user.Id
            DisplayName = $user.DisplayName
            UPN         = $user.UserPrincipalName
            Role        = $role.DisplayName
            AccountEnabled = $user.AccountEnabled
        }
    }
}

if (-not $rawUserRoles) {
    Write-Host "`nNo admin users found in the selected roles." -ForegroundColor Yellow
    $global:CISCheckResult = "FAIL"
    return
}


# -----------------------------------------------------------
# 4. Aggregate roles per user
# -----------------------------------------------------------
$grouped = $rawUserRoles | Group-Object UPN

$adminList = @()

foreach ($g in $grouped) {
    $first = $g.Group | Select-Object -First 1
    $roles = $g.Group.Role | Sort-Object -Unique

    $adminList += [pscustomobject]@{
        Control            = "1.1.4 (L1)"
        DisplayName        = $first.DisplayName
        UPN                = $first.UPN
        UserId             = $first.UserId
        Roles              = $roles
        AccountEnabled     = $first.AccountEnabled

        AssignedSkus       = @()
        EnabledServicePlans= @()
        HasReducedFootprint= $false   # to be computed
    }
}


# -----------------------------------------------------------
# 5. Define what "reduced application footprint" means
#    TUNE THIS for your org.
#
#    Below is a conservative example: if an admin has any of these
#    plans enabled, they are considered to have a full/expanded
#    application footprint (Exchange, Teams, SharePoint, OneDrive etc.).
# -----------------------------------------------------------
$HighRiskServicePlans = @(
    # Exchange Online
    "EXCHANGE_S_STANDARD",
    "EXCHANGE_S_ENTERPRISE",
    "EXCHANGE_S_FOUNDATION",
    "EXCHANGE_S_STANDARD_MIDMARKET",
    "EXCHANGE_S_ESSENTIALS",
    "EXCHANGE_S_BASIC",

    # SharePoint / OneDrive / Office web apps
    "SHAREPOINTSTANDARD",
    "SHAREPOINTENTERPRISE",
    "SHAREPOINTWAC",
    "ONEDRIVE_BASIC",
    "ONEDRIVE_ENTERPRISE",

    # Teams / Skype
    "MCOSTANDARD",
    "MCOEV",
    "MCOPSTN1",
    "MCOPSTN2",
    "TEAMS1",
    "TEAMS_FREE",
    "MCO_TEAMS_IW",

    # Yammer, Sway, Stream, etc. (expand as needed)
    "YAMMER_ENTERPRISE",
    "SWAY",
    "STREAM_O365_E3",
    "STREAM_O365_E5"
)

# Admin-safe examples (typically acceptable on admin-only accounts):
#   - AAD_P1, AAD_P2, RMS_S_ENTERPRISE, etc.
# You can explicitly ALLOW-only if you prefer a positive list approach.


# -----------------------------------------------------------
# 6. Retrieve license & service plan details per admin user
# -----------------------------------------------------------
Write-Host "`nRetrieving license and service plan information for each admin account..." -ForegroundColor Cyan

foreach ($entry in $adminList) {
    try {
        $licDetails = Get-MgUserLicenseDetail -UserId $entry.UserId -ErrorAction Stop

        if ($licDetails) {
            $entry.AssignedSkus = ($licDetails | Select-Object -ExpandProperty SkuPartNumber -ErrorAction SilentlyContinue) `
                                  | Sort-Object -Unique

            $enabledPlans = @()
            foreach ($ld in $licDetails) {
                foreach ($sp in $ld.ServicePlans) {
                    if ($sp.ProvisioningStatus -eq "Success") {
                        $enabledPlans += $sp.ServicePlanName
                    }
                }
            }

            $entry.EnabledServicePlans = $enabledPlans | Sort-Object -Unique
        } else {
            $entry.AssignedSkus        = @()
            $entry.EnabledServicePlans = @()
        }
    }
    catch {
        # If we fail to read license details, leave them empty
        $entry.AssignedSkus        = @()
        $entry.EnabledServicePlans = @()
    }

    # Compute HasReducedFootprint:
    # TRUE if no "high-risk" service plans are enabled.
    $hasHighRisk = $false
    foreach ($plan in $entry.EnabledServicePlans) {
        if ($HighRiskServicePlans -contains $plan) {
            $hasHighRisk = $true
            break
        }
    }

    $entry.HasReducedFootprint = -not $hasHighRisk
}


# -----------------------------------------------------------
# 7. Evaluate CIS 1.1.4 & output
# -----------------------------------------------------------
$final = $adminList | Select-Object `
    Control,
    DisplayName,
    UPN,
    @{Name="Roles";Expression={($_.Roles -join ", ")}},
    AccountEnabled,
    @{Name="AssignedSkus";Expression={($_.AssignedSkus -join ", ")}},
    @{Name="EnabledServicePlans";Expression={($_.EnabledServicePlans -join ", ")}},
    HasReducedFootprint

Write-Host ""
Write-Host "===== CIS 1.1.4 (L1) – Admin Account License Footprint =====" -ForegroundColor Cyan
Write-Host "Total administrative users found : $($final.Count)"
Write-Host "Control intent: Admin accounts should NOT have a full productivity/application footprint." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$final | Sort-Object DisplayName | Format-Table -AutoSize

$nonCompliant = $final | Where-Object { $_.HasReducedFootprint -eq $false }
$compliant    = $final | Where-Object { $_.HasReducedFootprint -eq $true }

Write-Host ""
Write-Host "Admins with reduced footprint (HasReducedFootprint = True)    : $($compliant.Count)" -ForegroundColor Green
Write-Host "Admins with full/wide footprint (HasReducedFootprint = False): $($nonCompliant.Count)" -ForegroundColor Red

if ($nonCompliant.Count -gt 0) {
    Write-Host "`nRESULT: FAIL – One or more admin accounts appear to have full/wide application licenses." -ForegroundColor Red
    $global:CISCheckResult = "FAIL"
} else {
    Write-Host "`nRESULT: PASS – All admin accounts appear to have reduced application footprints." -ForegroundColor Green
    $global:CISCheckResult = "PASS"
}

# -----------------------------------------------------------
# 8. Optional CSV export
# -----------------------------------------------------------
if ($ReportPath) {
    $final | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nAdmin license footprint report exported to: $ReportPath" -ForegroundColor Cyan
}

Write-Host "`nCIS Control 1.1.4 check complete.`n" -ForegroundColor Cyan
