<img width="282" height="131" alt="TUM_logo" src="https://github.com/user-attachments/assets/1896c0be-2ae6-413b-a2ed-bfa2ef3cd5ae" />

# ThinkPad Utility Management
Windows PowerShell diagnostic tool for ThinkPad hardware

A read-only diagnostic tool that reads battery health, cycle data, EC telemetry, warranty information, and memory configuration directly from Windows WMI and Lenovo EC interfaces.

> This project is not affiliated with or endorsed by Lenovo. ThinkPad is a trademark of Lenovo.

---

## Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / 11 |
| Shell | Windows PowerShell 5.1 or PowerShell 7+ |
| Privileges | **Administrator** (required — see below) |
| Lenovo SIF | Required for Options 2, 3, 5, 10, 11 |
| Commercial Vantage | Required for Option 11 (CDRT Odometer) |

---

## Why Administrator is Required

This script reads data directly from the system BIOS and Embedded Controller (EC) via two WMI namespaces that Windows restricts to Administrator-level processes:

- **`root\wmi`** — ACPI battery interfaces served by the inbox `wmiacpi.sys` driver. Crossing the kernel boundary to read ACPI control method output requires elevation.
- **`root\Lenovo`** — Registered by Lenovo System Interface Foundation (SIF). This namespace exposes both read and write methods (charge thresholds, BIOS settings), so the entire namespace is gated at Administrator level — Windows cannot split a namespace into read-only and read-write sections at the ACL level.

The script handles this automatically: if not already elevated, it will trigger a UAC prompt and relaunch itself in an elevated window.

---

## Running the Script

```powershell
# Right-click PowerShell → Run as Administrator, then:
.\ThinkPad_Utility_Management.ps1
```

Or let the script self-elevate — double-click it in Explorer or run it from a non-elevated shell and accept the UAC prompt.

---

## Code Signing and Execution Policy

PowerShell has an **execution policy** that controls which scripts are allowed to run. Depending on how your system is configured — especially in enterprise environments — you may encounter an error like:

```
File cannot be loaded because running scripts is disabled on this system.
```

### Check your current policy

```powershell
Get-ExecutionPolicy -List
```

The effective policy is the most restrictive one across all scopes.

### Policy levels that affect this script

| Policy | Behaviour |
|---|---|
| `Restricted` | No scripts run at all. Default on consumer Windows. |
| `RemoteSigned` | Scripts downloaded from the internet must be signed. Locally created scripts run freely. |
| `AllSigned` | **All** scripts must be signed by a trusted certificate, regardless of origin. Common in enterprise. |
| `Unrestricted` | All scripts run; downloaded scripts prompt for confirmation. |
| `Bypass` | Nothing is blocked. Typically used in CI pipelines. |

---

### Option A — Unblock the file (personal use, RemoteSigned policy)

If you downloaded this script from the internet (GitHub, X, a browser), Windows marks it with an NTFS Zone Identifier (`Zone.Identifier = 3`). Under `RemoteSigned`, this alone is enough to block execution even without a signing requirement.

Unblock it with:

```powershell
Unblock-File -Path .\ThinkPad_Utility_Management.ps1
```

Or right-click the `.ps1` file → Properties → check **Unblock** → OK.

This removes the Zone Identifier and tells PowerShell to treat the file as locally authored. No certificate is needed.

---

### Option B — Self-signed certificate (personal use, AllSigned policy)

If your policy is `AllSigned` and you manage your own machine, you can sign the script with a self-signed certificate. This satisfies the policy without purchasing a commercial certificate.

**Step 1 — Create a self-signed code signing certificate**

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=ThinkPad Utility Management" `
    -Type CodeSigning `
    -CertStoreLocation Cert:\CurrentUser\My `
    -HashAlgorithm SHA256
```

**Step 2 — Trust it on your own machine**

The certificate must be in both the Trusted Root and Trusted Publishers stores, otherwise PowerShell will not accept it:

```powershell
$store = [System.Security.Cryptography.X509Certificates.X509Store]::new("Root","CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

$store = [System.Security.Cryptography.X509Certificates.X509Store]::new("TrustedPublisher","CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()
```

**Step 3 — Sign the script**

```powershell
Set-AuthenticodeSignature `
    -FilePath .\ThinkPad_Utility_Management.ps1 `
    -Certificate $cert `
    -HashAlgorithm SHA256
```

**Step 4 — Verify**

```powershell
Get-AuthenticodeSignature .\ThinkPad_Utility_Management.ps1
```

`Status` should read `Valid`.

> **Important:** A self-signed certificate is only trusted on the machine where you created it. It will not satisfy `AllSigned` on any other machine unless you manually import and trust the certificate there too.

---

### Option C — Trusted CA-signed certificate (enterprise / distribution)

If you are distributing this script within an organisation, or publishing it for others to run under `AllSigned` policy without manual certificate import steps, you need a certificate from a trusted Certificate Authority (CA).

**Internal enterprise CA (Active Directory)**

Most enterprise environments have an Active Directory Certificate Services (ADCS) deployment. Request a code signing certificate from your internal CA:

```powershell
# Request via certlm.msc or:
Get-Certificate `
    -Template "CodeSigning" `
    -CertStoreLocation Cert:\CurrentUser\My
```

Certificates issued by your domain CA are automatically trusted on all domain-joined machines. This is the recommended path for IT-managed ThinkPad fleets.

**Commercial CA (public distribution)**

For public distribution where you want the script to run on any machine under `AllSigned` without importing anything, you need a code signing certificate from a publicly trusted CA. Common options:

| CA | Notes |
|---|---|
| DigiCert | Industry standard, widely used for PowerShell signing |
| Sectigo | Cost-effective option with good Windows trust chain coverage |
| GlobalSign | Common in enterprise procurement |

All of these are paid certificates. The certificate must be an **Extended Validation (EV)** or standard **Code Signing** certificate — not a TLS/SSL certificate.

**Sign with a CA-issued certificate**

```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -match "YourName" -and $_.HasPrivateKey } |
    Select-Object -First 1

Set-AuthenticodeSignature `
    -FilePath .\ThinkPad_Utility_Management.ps1 `
    -Certificate $cert `
    -TimestampServer "http://timestamp.digicert.com" `
    -HashAlgorithm SHA256
```

> Always use a **timestamp server** (`-TimestampServer`) when signing for distribution. Without it, the signature becomes invalid the moment the certificate expires — even for scripts signed while the certificate was valid.

---

### Option D — Bypass for a single session (testing only)

If you are testing on a machine you control and do not want to deal with certificates:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\ThinkPad_Utility_Management.ps1
```

This overrides the policy for that one invocation only and does not change the system-wide setting. **Do not use this in production or enterprise deployments.**

---

## Menu Options

| Option | Name | Requires |
|---|---|---|
| 1 | Battery Static Data | ACPI (`root\wmi`) |
| 2 | Lenovo Battery Cycles | SIF (`root\Lenovo`) |
| 3 | Warranty Info | SIF (`root\Lenovo`) |
| 4 | Full Charge Capacity | ACPI (`root\wmi`) |
| 5 | Comprehensive Battery Analysis | ACPI + SIF (SIF optional but degrades) |
| 6 | Memory Info | `root\cimv2` |
| 7 | Export Report | All available sources |
| 8 | Diagnostic Info | All sources (availability check) |
| 9 | Battery Age Estimation | ACPI + SIF (degrades gracefully) |
| 10 | Battery Charge Threshold | SIF (`root\Lenovo`) |
| 11 | CDRT Odometer | SIF + Commercial Vantage |
| 12 | About | — |
| 13 | Error Log Viewer | — |
| 14 | Exit | — |

---

## Battery Health Classification

Battery health is evaluated using capacity percentage and cycle count and reported across five tiers. The severity level drives the colour shown in the header and menu alert on every screen.

| Classification | Health % | Severity | Colour |
|---|---|---|---|
| Good | ≥ 80% | 0 | White |
| Warning | 60–79% | 1 | Cyan |
| Fair | 40–59% | 2 | Yellow |
| Poor | < 40% | 3 | Magenta |
| Critical | < 20% or status failure | 4 | Red |

Cycle count is factored in alongside capacity percentage:

- More than 500 cycles and below 75% capacity → **Fair**
- More than 1,000 cycles and below 60% capacity → **Poor**
- More than 1,500 cycles → **Critical**

A battery whose WMI status reports a failure keyword (`Failed`, `Critical`, `Error`, `Degraded`) is always classified as **Critical** regardless of capacity.

The **Warning** tier aligns with Lenovo Vantage's updated alert threshold, which flags batteries below 80% as needing attention.

---

## Battery Age Estimation (Option 9)

Age is estimated using two independent methods that are cross-validated against each other.

**Method 1 — Cycle-based.** Assumes a conservative baseline of 250 charge cycles per year, derived from ASUS, Lenovo, and independent battery longevity studies.

**Method 2 — Capacity-based.** Uses a simplified linear degradation model: 100% health corresponds to a new battery (0 years), and full degradation corresponds to approximately 5 years — equating to roughly 4% capacity loss per year.

When both methods are available, a weighted average is used (45% cycle-based, 55% capacity-based) with a confidence rating:

| Confidence | Condition |
|---|---|
| High | Both estimates agree within 1 year |
| Medium | Estimates diverge by 1–2 years, or only one data source is available |
| Low | Estimates diverge by more than 2 years |

Results are statistical approximations. Actual age may vary based on usage habits, storage temperature, and whether the battery has been replaced.

---

## Battery Charge Threshold Manager (Option 10)

Reads and writes battery charge thresholds directly to the ThinkPad BIOS/EC firmware via `Lenovo_BiosSetting`, `Lenovo_SetBiosSetting`, and `Lenovo_SaveBiosSettings` in `root\wmi`.

**Important notes before changing thresholds:**

- Changes are written directly to firmware and take effect on the **next reboot**.
- Start threshold must always be lower than Stop threshold.
- Recommended range: Start 40–75%, Stop 60–80%.
- Setting Stop to 100% disables the upper limit.
- A BIOS Supervisor Password may be required depending on system policy.
- Not all ThinkPad models expose threshold settings via WMI.

**WMI return codes:**

| Code | Meaning |
|---|---|
| `Success` | Change staged — reboot required to apply |
| `Not Supported` | Setting not available on this model |
| `Invalid Parameter` | Value or setting name is incorrect |
| `Access Denied` | BIOS Supervisor Password required |
| `BIOS Error` | Firmware-level failure |

---

## Genuine Battery Detection

The script checks the battery manufacturer string reported by firmware against a list of known Lenovo-authorised suppliers. A warning is displayed for any battery that does not match a known entry.

**Built-in supplier list:** LENOVO, PANASONIC, SANYO, SONY, MURATA, LGC, LG, SDI, SAMSUNG, CELXPERT, ATL, COSMX, BYD, NVT, SUNWODA, SMP.

To extend detection without editing the script, place a `manufacturers.json` file in the same directory as the `.ps1` file. If present and valid, it replaces the built-in list entirely. If missing, unreadable, or empty, the script falls back to the built-in list silently.

**`manufacturers.json` format:**

```json
{
  "genuineManufacturers": [
    { "name": "LENOVO",   "fullName": "Lenovo",                    "region": "CN", "notes": "" },
    { "name": "CELXPERT", "fullName": "Celxpert Electronics Corp.", "region": "TW", "notes": "Confirmed Lenovo-authorised supplier" },
    { "name": "NEWMFR",   "fullName": "New Manufacturer Inc.",      "region": "XX", "notes": "" }
  ]
}
```

Only the `name` field is required. It is matched case-insensitively as a substring against the manufacturer string reported by the battery firmware.

---

## CDRT Odometer (Option 11)

The Commercial Deployment Readiness Tool (CDRT) Odometer is deployed by Lenovo Commercial Vantage. It extends `Lenovo_Odometer` in `root\Lenovo` with three cumulative lifetime counters:

| Field | Description |
|---|---|
| CPU Uptime | Total cumulative minutes the CPU has been active, displayed as minutes, hours, and days |
| Shock Events | Accelerometer events exceeding the CDRT vibration threshold — includes minor bumps and bag movement, not only drops. Counts in the hundreds are normal for a well-travelled machine. |
| Thermal Events | Number of times the CPU throttled due to reaching a critical temperature |

This data is valuable for enterprise asset management and pre-owned ThinkPad evaluation.

---

## Driver Dependencies

### Lenovo System Interface Foundation (SIF)

Required for Options 2, 3, 5, 10, and 11. SIF registers the `root\Lenovo` WMI namespace and the EC-facing classes used by this script: `Lenovo_Odometer`, `Lenovo_Battery`, `Lenovo_BiosSetting`, `Lenovo_SetBiosSetting`, `Lenovo_SaveBiosSettings`, and `Lenovo_WarrantyInformation`.

Install from: **support.lenovo.com** → Drivers & Software → search "System Interface Foundation"

> The `root\Lenovo` namespace is a ThinkPad-exclusive firmware feature. It is not present on IdeaPad, Yoga, Legion, ThinkBook, or other Lenovo product lines by design — this cannot be resolved by installing SIF on a non-ThinkPad device.

### Lenovo Commercial Vantage

Required for Option 11 (CDRT Odometer). Commercial Vantage deploys the CDRT MOF files that extend `Lenovo_Odometer` with CPU uptime, shock event, and thermal event tracking.

- **Microsoft Store:** search "Lenovo Commercial Vantage"
- **Enterprise:** deploy via SCCM or MDM using the Lenovo CDRT package from support.lenovo.com

> Commercial Vantage is intended for business ThinkPads. Consumer models use Lenovo Vantage, which does not include the CDRT Odometer.

---

## Data Sources

The script queries multiple WMI sources per feature and falls back gracefully when a preferred source is unavailable.

| WMI Class | Namespace | Used For |
|---|---|---|
| `Lenovo_Battery` | `root\Lenovo` | Full battery identity, health, charge state, electrical readings |
| `Lenovo_Odometer` | `root\Lenovo` | Battery cycle count, shock events, thermal events, CDRT CPU uptime |
| `Lenovo_BiosSetting` | `root\wmi` | Reading charge thresholds |
| `Lenovo_SetBiosSetting` | `root\wmi` | Writing charge thresholds |
| `Lenovo_SaveBiosSettings` | `root\wmi` | Committing BIOS changes |
| `Lenovo_WarrantyInformation` | `root\Lenovo` | Warranty serial, start/end dates |
| `BatteryFullChargedCapacity` | `root\wmi` | Full charge capacity (ACPI fallback) |
| `BatteryStaticData` | `root\wmi` | Design capacity, manufacturer, serial number |
| `Win32_Battery` | `root\cimv2` | Estimated charge percentage (fallback) |
| `Win32_PhysicalMemory` | `root\cimv2` | RAM modules, capacity, speed, manufacturer |
| `Win32_ComputerSystem` | `root\cimv2` | System model string |
| `Win32_BIOS` | `root\cimv2` | BIOS version and release date |

Design capacity is resolved from `BatteryStaticData` first, with `powercfg /batteryreport /XML` as a fallback if the WMI class returns zero or is unavailable.

---

## Export Report (Option 7)

Generates a plain-text diagnostic report saved to `%USERPROFILE%\ThinkPad_Report.txt`. If any errors were logged during the session, a separate error log is saved to `%USERPROFILE%\ThinkPad_ErrorLog.txt`.

The report includes: Windows battery data, Lenovo EC cycle/shock/thermal data, warranty information, full charge capacity, health analysis, and comprehensive battery classification.

---

## Privacy

This script runs entirely on the local machine. No data is transmitted, collected, or stored outside of the optional local export report, which is written to `%USERPROFILE%` only when Option 7 is used.

---

## Power User Guide

Want to build your own ThinkPad diagnostic tool using the same architecture?

[Open the Power User Guide](PowerUsers.md)

---

## Legal

This project is not affiliated with or endorsed by Lenovo. ThinkPad is a trademark of Lenovo.

Provided for educational and diagnostic purposes only. This script makes no modifications to system configuration, firmware, or settings outside of the explicit charge threshold write in Option 10, which requires deliberate user confirmation. Battery age estimation results are statistical approximations. Actual age may vary based on usage habits, temperature, and service history.

Built with Windows PowerShell, ChatGPT GPT 5.4, and Claude Sonnet 4.6.

## EU Battery Regulation Compliance
 
This tool is designed in compliance with **Regulation (EU) 2023/1542**, effective
**18 February 2027**, which establishes consumer rights around portable battery
replacement and access to battery health information.
 
### Background: the Lenovo battery lockdown
 
From approximately 2016, certain Lenovo and ThinkPad devices enforced a **battery
manufacturer whitelist via BIOS firmware**. Batteries from manufacturers not present
on the allowlist were flagged as non-genuine and, in some configurations, blocked
from charging. Diagnostic software that surfaced this whitelist check — including
earlier versions of this tool — would emit warnings such as "non-genuine battery
detected" for third-party replacement cells.
 
Under EU Battery Regulation, this behaviour is no longer appropriate. Consumers
have an **explicit legal right** to replace portable batteries with third-party
cells. Treating a non-OEM manufacturer name as a warning or failure condition
conflicts with that right and was removed in **v2.0** of this tool.
 
### What changed in v2.0
 
- **Manufacturer allowlist check removed as a pass/fail gate.** The function
  `Get-GenuineManufacturerList` is retained for informational display only
  (e.g. identifying a known OEM supplier) but produces no warnings, failures,
  or output implying a third-party battery is unsafe or non-compliant.
- **Non-genuine warning suppressed across all battery functions.** Four
  separate call sites previously capable of emitting a non-genuine warning
  have been updated with `# DEPRECATED (v2.0): Non-genuine warning removed —
  EU Battery Regulation compliance.`
- **Serial/barcode mismatch detection** replaces manufacturer gating in the
  deployment readiness workflow. A mismatched serial is a legitimate data
  integrity concern; a non-OEM manufacturer name is not.
 
### Battery health data access
 
The EU regulation also requires that consumers have access to battery health
information sufficient to make informed replacement decisions. This tool
surfaces the following data points directly from Windows WMI and ACPI
interfaces, without relying on Lenovo Vantage or any OEM software:
 
- Full charge capacity vs. design capacity (health percentage)
- Cycle count (where available via `root\Lenovo` or `Lenovo_Odometer`)
- Battery temperature with severity classification
- Degradation rate and projected time to 80% replacement threshold
- Persistent trend log with configurable logging interval
 
The 80% threshold used for replacement advisories aligns with the level at
which the EU regulation considers a battery to have reached end of useful life
for consumer purposes.
 
### Regulation reference
 
> Regulation (EU) 2023/1542 of the European Parliament and of the Council
> of 12 July 2023 concerning batteries and waste batteries.
> OJ L 191, 28.7.2023, p. 1–117.
> Effective date for portable battery replaceability requirements: 18 February 2027.
 
This tool makes no modifications to system firmware, BIOS settings, or Windows
configuration. It is a read-only diagnostic utility. Nothing in this tool
constitutes legal advice.