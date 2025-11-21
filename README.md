# CIS Microsoft 365 Foundations – PowerShell Automation (Work in Progress)

This repository contains **PowerShell scripts** designed to help automate checks from the  
**CIS Microsoft 365 Foundations Benchmark (v5.0.0)**.

⚠️ **IMPORTANT: This project is a work in progress.**  
Many scripts are **incomplete**, **untested**, or **subject to major changes**.  
Some scripts **may not work as expected**, may require additional permissions, or may need  
significant refinement before being considered production-ready.

Use these scripts as a **starting point**, not as a guaranteed-complete compliance toolkit.

---

## ☕ Support the Project

If you want to support this project, please pick a benchmark and submit a PR, or if you find it useful, you can buy me a coffee on Ko‑Fi:

👉 **https://ko-fi.com/r04chy**


---

## 🚧 Project Status

Scripts currently included:

### ✔ Section 1.1 – Administrative Accounts
| CIS Control | Script | Status |
|------------|--------|--------|
| **1.1.1** – Ensure administrative accounts are cloud-only | `CIS_1_1_1.ps1` | Complete |
| **1.1.2** – Ensure two emergency access accounts exist | `CIS_1_1_2.ps1` | Complete |
| **1.1.3** – Ensure between two and four global admins | `CIS_1_1_3.ps1` | Complete |
| **1.1.4** – Ensure admin accounts use reduced-footprint licenses | `CIS_1_1_4.ps1` | Complete |

### ✔ Section 1.2 – Identity Governance
| CIS Control | Script | Status |
|------------|--------|--------|
| **1.2.1** – Ensure only approved public groups exist | `CIS_1_2_1.ps1` | Complete |
| **1.2.2** – Ensure sign-in to shared mailboxes is blocked | `CIS_1_2_2.ps1` | Complete |

### ✔ Section 1.3 – Account Policies
| CIS Control | Script | Status |
|------------|--------|--------|
| **1.3.1** – Ensure password expiration policy is set to “never expire” | `CIS_1_3_1.ps1` | Complete |
| **1.3.2** – Ensure idle session timeout ≤ 3 hours (unmanaged devices) | `CIS_1_3_2.ps1` | Complete (PnP + SPO variants) |

More controls will be added progressively.

---

## 📘 Purpose

The goal of this repository is to:

- Provide automation tooling for **CIS Microsoft 365 Foundations** checks  
- Assist security engineers with **audits, baselines, and compliance validation**  
- Reduce manual workload by using Microsoft Graph, Exchange Online, SharePoint Online, and PnP PowerShell

This is **not** an official CIS tool and should not be treated as a substitute for  
a full audit or professional review.

---

## 🛠 Requirements

Install required modules:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
```

Many scripts require high-privilege Microsoft Graph scopes, including:

- Directory.Read.All
- RoleManagement.Read.Directory
- UserAuthenticationMethod.Read.All
- Policy.Read.ConditionalAccess
- Group.Read.All
- Domain.Read.All

SharePoint-related controls require:

- SharePoint Administrator  
- PnP interactive login capability  
- A valid SharePoint Admin URL (auto-detected where possible)

---

## ⚠️ Disclaimer

This project is provided **as-is** with **no warranty**, express or implied.

Scripts may:

- Contain bugs  
- Change without notice  
- Require tenant-specific fixes  
- Break due to future Graph API changes  

**Always test in a non-production environment first.**

---

## 🤝 Contributing

Contributions are welcome!  
When contributing, please reference the corresponding **CIS control number** and follow  
the existing structure and naming conventions.

---

## 📄 License

This project is MIT licensed — see `LICENSE.md` for full details.
