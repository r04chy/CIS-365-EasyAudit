# CIS Microsoft 365 Foundations – PowerShell Automation (Work in Progress)

This repository provides **PowerShell automation scripts** for validating controls in the  
**CIS Microsoft 365 Foundations Benchmark v5.0.0**.

⚠️ **IMPORTANT: Work in Progress**  
These scripts are under active development. Some controls are fully automated, while others  
require manual review due to Microsoft API limitations. Many scripts may require adjustments  
depending on your tenant configuration.

Use these tools as a **starting point**, not a drop-in replacement for a full audit.

---

## ☕ Support the Project

If you want to support this project, please pick a benchmark and submit a PR,  
or if you find it useful, you can buy me a coffee on Ko‑Fi:

👉 **https://ko-fi.com/r04chy**

Your support helps keep this project moving!

---

# 📘 Purpose

The goal of this repository is to simplify auditing and validating CIS Microsoft 365 controls by:

- Automating as many checks as possible  
- Producing consistent & machine-readable PASS/FAIL outputs  
- Providing detailed visibility into risky configuration areas  
- Reducing the repetitive manual work of M365 security auditing  

This project is *not* an official CIS product.

---

# 🚧 Script Coverage Progress

Below is the current coverage of CIS v5.0.0 controls implemented in this repository.

---

## ✔ **Section 1.1 – Administrative Accounts**

| CIS Control | Script | Status | Notes |
|------------|--------|--------|-------|
| **1.1.1** – Ensure administrative accounts are cloud-only | `CIS_1_1_1.ps1` | Complete | Automated |
| **1.1.2** – Ensure two emergency access accounts exist | `CIS_1_1_2.ps1` | Complete | Includes full MFA profile |
| **1.1.3** – Ensure between 2–4 global admins are designated | `CIS_1_1_3.ps1` | Complete | Automated |
| **1.1.4** – Ensure admin accounts use reduced-footprint licenses | `CIS_1_1_4.ps1` | Complete | Automated, customizable footprint rules |

---

## ✔ **Section 1.2 – Identity Governance**

| CIS Control | Script | Status | Notes |
|------------|--------|--------|-------|
| **1.2.1** – Ensure only approved public groups exist | `CIS_1_2_1.ps1` | Complete | Automated |
| **1.2.2** – Ensure sign-in to shared mailboxes is blocked | `CIS_1_2_2.ps1` | Complete | Automated |

---

## ✔ **Section 1.3 – Account Policies**

| CIS Control | Script | Status | Notes |
|------------|--------|--------|-------|
| **1.3.1** – Password expiration policy set to “never expire” | `CIS_1_3_1.ps1` | Complete | Automated |
| **1.3.2** – Idle session timeout ≤ 3 hours (unmanaged devices) | `CIS_1_3_2.ps1` | Complete | Auto-detects SPO admin URL (SPO + PnP variants) |
| **1.3.3** – External calendar sharing disabled | `CIS_1_3_3.ps1` | Complete | Automated |
| **1.3.4** – Restrict user-owned apps/services | `CIS_1_3_4.ps1` | Complete | Automated |
| **1.3.5** – Forms internal phishing protection enabled | `CIS_1_3_5.ps1` | Partial | Depends on Forms API availability |
| **1.3.6** – Customer Lockbox enabled | `CIS_1_3_6.ps1` | Complete | Automated |
| **1.3.7** – Third-party storage services restricted | `CIS_1_3_7.ps1` | Complete | Automated (Service Principal check) |
| **1.3.8** – Sway external sharing restricted | `CIS_1_3_8.ps1` | Manual | No API exists — script provides manual instructions |

---

# 🛠 Requirements

Install required modules:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
```

### Common Microsoft Graph permission scopes used:

- Directory.Read.All  
- RoleManagement.Read.Directory  
- UserAuthenticationMethod.Read.All  
- Policy.Read.ConditionalAccess  
- Group.Read.All  
- Domain.Read.All  
- OrgSettings.Read.All  
- Application.Read.All  

### Exchange Online requirements:

- Exchange Administrator or Global Administrator  

### SharePoint Online requirements:

- SharePoint Administrator  
- Ability to authenticate to SPO or PnP with modern auth  

---

# ⚠️ Notes on Special Cases

### 🔸 1.3.5 – Microsoft Forms phishing protection  
The required API is not rolled out to all tenants.  
If unavailable, the script will warn and mark the control as unsupported.

### 🔸 1.3.8 – Sway external sharing  
This setting **cannot be automated**.  
Microsoft exposes no API for reading or setting it.  
The script provides manual validation guidance and sets:  

```
$global:CISCheckResult = "MANUAL"
```

---

# 🤝 Contributing

Contributions are welcome!

If submitting a PR:

- Reference the corresponding **CIS control number**  
- Follow naming conventions (`CIS_<section>_<control>.ps1`)  
- Ensure the script produces a PASS/FAIL/MANUAL state cleanly  
- Include clear commentary explaining what it checks and why  

---

# 📄 License

This project is MIT licensed — see `LICENSE.md` for details.
