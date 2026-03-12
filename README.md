# ThinkPad-Utility-Management
A read-only diagnostic tool that reads battery health, cycle data, EC telemetry, warranty information, and memory configuration directly from Windows WMI and Lenovo EC interfaces.

> This project is not affiliated with or endorsed by Lenovo. ThinkPad is a trademark of Lenovo.
---

## Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / 11 |
| Shell | Windows PowerShell 5.1 or PowerShell 7+ |
| Privileges | **Administrator** (required — see below) |
| Lenovo SIF | Required for Options 2, 5, 11 |
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
# Retrieve the certificate from your store (filter by thumbprint or subject)
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

This overrides the policy for that one invocation only and does not change the system-wide setting. **Do not use this in production or enterprise deployments** — it defeats the purpose of having an execution policy.

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

## Driver Dependencies

### Lenovo System Interface Foundation (SIF)
Required for Options 2, 3, 5, 10, and 11. SIF registers the `root\Lenovo` WMI namespace and the EC-facing classes (`Lenovo_Odometer`, `Lenovo_BiosSetting`, `Lenovo_WarrantyInformation`).

Install from: **support.lenovo.com** → Drivers & Software → search "System Interface Foundation"

### Lenovo Commercial Vantage
Required for Option 11 (CDRT Odometer). Commercial Vantage deploys the CDRT MOF files that extend `Lenovo_Odometer` with CPU uptime, shock event, and thermal event tracking.

- **Microsoft Store:** search "Lenovo Commercial Vantage"
- **Enterprise:** deploy via SCCM or MDM using the Lenovo CDRT package from support.lenovo.com

> Commercial Vantage is intended for business ThinkPads. Consumer models use Lenovo Vantage, which does not include the CDRT Odometer.

---

## Privacy

This script runs entirely on the local machine. No data is transmitted, collected, or stored outside of the optional local export report (`ThinkPad_Report.txt`), which is written to `%USERPROFILE%` only when Option 7 is used.

---

## Legal

This project is not affiliated with or endorsed by Lenovo. ThinkPad is a trademark of Lenovo.

Provided for educational and diagnostic purposes only. Battery age estimation results are statistical approximations. Actual age may vary based on usage habits, temperature, and service history.

Built with Windows PowerShell, ChatGPT GPT 5.4, and Claude Sonnet 4.6.
