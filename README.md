# CIS Microsoft 365 Foundations – PowerShell Automation (Work in Progress)

This repository contains **PowerShell scripts** designed to help automate checks from the  
**CIS Microsoft 365 Foundations Benchmark (v5.0.0)**.

⚠️ **IMPORTANT: This project is a work in progress.**  
Many scripts are **incomplete**, **untested**, or **subject to major changes**.  
Some scripts **may not work as expected**, may require additional permissions, or may need  
significant refinement before being considered production-ready.

Use these scripts as a **starting point**, not as a guaranteed-complete compliance toolkit.

## 🚧 Project Status

- Scripts are being added incrementally starting from **Section 1.1 – Administrative Controls**.
- Some controls (e.g., MFA detection, license footprint analysis) involve complex logic and  
  may require fine-tuning for your environment.
- Additional features (CSV outputs, combined reports, pipeline integration, etc.)  
  may be added later.

## 📘 Purpose

The goal of this repository is to:

- Provide automation tooling for **CIS Microsoft 365 Foundations** checks  
- Assist security engineers with **audits, baselines, and compliance validation**  
- Reduce manual workload by using the **Microsoft Graph PowerShell SDK**

This is **not** an official CIS tool and should not be treated as a substitute for  
a full audit or professional review.

## 🛠 Requirements

Install the Microsoft Graph PowerShell SDK:

```
Install-Module Microsoft.Graph -Scope CurrentUser
```

Most scripts require one or more privileged Graph scopes:

- Directory.Read.All
- RoleManagement.Read.Directory
- UserAuthenticationMethod.Read.All
- Policy.Read.ConditionalAccess

## 📁 Script Coverage (So Far)

| CIS Control | Status | Notes |
|-------------|--------|-------|
| **1.1.1** – Cloud-only Admin Accounts | Complete | Fully automated |
| **1.1.2** – Emergency Access Accounts | Complete | MFA profile included |
| **1.1.3** – Global Admin Count | Complete | Pass/Fail logic included |
| **1.1.4** – Reduced License Footprint | Complete | Requires customization |
| Other controls | 🚧 In Progress | Will be added gradually |

## ⚠️ Disclaimer

This project is provided **as-is** with **no warranty**, express or implied.

Always test in a **non-production** environment first.

## 🤝 Contributing

Contributions are welcome!  
When contributing, please reference the corresponding **CIS control number**.

## 📄 License

This project is MIT licensed — see `LICENSE.md` for details.
