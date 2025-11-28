CIS Microsoft 365 Foundations Benchmark – Automated Checks (v6.0.0)

This project provides PowerShell-based auditing scripts for the CIS Microsoft 365 Foundations Benchmark version 6.0.0 (31 October 2025). It is designed to help administrators, auditors, and security teams validate Microsoft 365 configuration against CIS Level 1 and Level 2 controls.

⸻

✅ Supported Benchmark Version

This repository now fully aligns with:

📘 CIS Microsoft 365 Foundations Benchmark v6.0.0
Release date: 2025-10-31

All scripts, control numbers, section references, and descriptions have been updated to match this version. Newly added controls in v6.0.0 are now included.

⸻

📂 Repository Structure

scripts/
  ├─ 1.x  Microsoft 365 Admin Center
  ├─ 2.x  Exchange Online / Anti-Spam
  ├─ 3.x  Audit & Compliance
  ├─ 4.x  Devices & Intune Prerequisites
  ├─ 5.x  Entra ID / Identity Governance
  ├─ 6.x  Exchange Configuration
  ├─ 7.x  SharePoint / OneDrive
  ├─ 8.x  Teams Admin Center
  └─ 9.x  Microsoft Fabric

Each script corresponds to a CIS control ID (e.g., 5.1.4.5.ps1 for LAPS).

⸻

🆕 New Controls Added in CIS Benchmark v6.0.0

The following controls are new in v6.0.0 and are now implemented in this repository:

Section 1 – Microsoft 365 Admin Center
	•	1.3.9 – Restrict shared bookings pages to select users

Section 2 – Anti-Spam & Mail
	•	2.1.15 – Ensure outbound anti‑spam message limits are in place

Section 5 – Entra ID Governance & Devices
	•	5.1.3.2 – Ensure users cannot create security groups
	•	5.1.4.1 – Restrict device join to Entra
	•	5.1.4.2 – Limit maximum devices per user
	•	5.1.4.3 – Ensure GA role is not added as local admin during Entra join
	•	5.1.4.4 – Limit local admin assignment during Entra join
	•	5.1.4.5 – Enable Local Administrator Password Solution (LAPS)
	•	5.1.4.6 – Restrict users from recovering BitLocker keys

Section 5.2 – Authentication Methods
	•	5.2.3.7 – Ensure email OTP authentication method is disabled

Section 6 – Exchange
	•	Several new/updated remediation requirements (see CHANGELOG)

⸻

📘 Script Coverage Progress (CIS v6.0.0)

This section lists every currently implemented script, grouped by CIS section.

⸻

Section 1 – Microsoft 365 Admin Center

1.3 – Account Policies

Control	Title	Script
1.3.1	Password expiration policy set to “never expire”	CIS_1_3_1.ps1
1.3.2	Idle session timeout for unmanaged devices	CIS_1_3_2.ps1
1.3.3	External calendar sharing disabled	CIS_1_3_3.ps1
1.3.4	Restrict user-owned apps/services	CIS_1_3_4.ps1
1.3.5	Microsoft Forms phishing protection enabled	CIS_1_3_5.ps1
1.3.6	Customer Lockbox enabled	CIS_1_3_6.ps1
1.3.7	Third-party storage services restricted	CIS_1_3_7.ps1
1.3.8	Sway external sharing restricted	CIS_1_3_8.ps1
1.3.9	Restrict shared bookings pages to select users	CIS_1_3_9.ps1


⸻

Section 2 – Anti-Spam & Exchange Online Protection (EOP)

2.1 – Exchange Online Protection Controls

Control	Title	Script
2.1.1	Safe Links for Office Apps enabled	CIS_2_1_1.ps1
2.1.2	Common Attachment Types Filter enabled	CIS_2_1_2.ps1
2.1.3	Notify admins when internal users send malware	CIS_2_1_3.ps1
2.1.4	Safe Attachments policy enabled	CIS_2_1_4.ps1
2.1.5	Safe Attachments for SPO/OneDrive/Teams	CIS_2_1_5.ps1
2.1.6	EOP spam policies notify administrators	CIS_2_1_6.ps1
2.1.7	Anti-phishing policy configured	CIS_2_1_7.ps1
2.1.8	SPF records published for all domains	CIS_2_1_8.ps1
2.1.9	DKIM enabled for all domains	CIS_2_1_9.ps1
2.1.10	DMARC records published for all domains	CIS_2_1_10.ps1
2.1.11	Comprehensive attachment filtering applied	CIS_2_1_11.ps1
2.1.12	Connection filter IP allow list not used	CIS_2_1_12.ps1
2.1.13	Connection filter safe list disabled	CIS_2_1_13.ps1
2.1.14	Inbound anti-spam policies do not allow domains	CIS_2_1_14.ps1
2.1.15	Outbound anti-spam message limits enforced	CIS_2_1_15.ps1


⸻

Section 5 – Entra ID Governance & Devices

Control	Title	Script
5.1.3.2	Ensure users cannot create security groups	5.1.3.2.ps1
5.1.4.1	Restrict device join to Entra	5.1.4.1.ps1
5.1.4.2	Limit maximum devices per user	5.1.4.2.ps1
5.1.4.3	Prevent GA role becoming local admin on join	5.1.4.3.ps1
5.1.4.4	Limit local admin assignment	5.1.4.4.ps1
5.1.4.5	Enable Local Administrator Password Solution (LAPS)	5.1.4.5.ps1
5.1.4.6	Restrict BitLocker key recovery rights	5.1.4.6.ps1


⸻

Section 5.2 – Authentication Methods

Control	Title	Script
5.2.3.7	Ensure email OTP authentication is disabled	5.2.3.7.ps1


⸻

🔧 Updated Controls (Audit Logic Changed) (Audit Logic Changed)

Many controls in v6.0.0 include updated audit procedures, passing values, and scripts.

The following categories were updated:
	•	Teams settings & policies (sections 8.1–8.6)
	•	Exchange Online remediation logic (notably 6.5.x)
	•	Authentication method assessments
	•	Access Review logic (5.3.2–5.3.3)
	•	Password Protection and Weak Auth Method detection

All scripts have been revised to reflect these.

⸻

▶️ Usage

Run any audit script individually:

.\scripts\5.1.4.5-Entra-LAPS.ps1

Or run the full benchmark audit:

.\run-all.ps1

Output is provided in both terminal format and JSON for ingestion into SIEM, evidence collection, or compliance reporting.

⸻

📄 CHANGELOG

A detailed changelog documenting updates from earlier benchmark versions to v6.0.0 is included in CHANGELOG.md.

⸻

☕ Support the Project

If you want to support this project, please pick a benchmark and submit a PR, or if you find it useful, you can buy me a coffee on Ko‑Fi:

👉 https://ko-fi.com/r04chy

⸻

📬 Contact / Contributions

Pull requests, corrections, and new checks are always welcome.

If you identify new changes in CIS benchmarks or Microsoft’s configuration surface, feel free to open an issue.
