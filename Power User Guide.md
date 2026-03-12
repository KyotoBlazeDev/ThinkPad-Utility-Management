# TUM Power User Guide
## How to Build Your Own ThinkPad Utility Script

This guide teaches you how to create your own diagnostic tool using the same
architecture and patterns found in ThinkPad Utility Management (TUM). No prior
WMI or PowerShell experience required — every concept is explained from scratch
with working code examples.

---

## Table of Contents

1. [What is WMI?](#1-what-is-wmi)
2. [Script Architecture](#2-script-architecture)
3. [The Script Skeleton](#3-the-script-skeleton)
4. [WMI Query Patterns and Fallback Chains](#4-wmi-query-patterns-and-fallback-chains)
5. [Safe Property Access](#5-safe-property-access)
6. [Error Logging](#6-error-logging)
7. [Adding a New Menu Option](#7-adding-a-new-menu-option)
8. [The Header and State System](#8-the-header-and-state-system)
9. [Full Working Example](#9-full-working-example)
10. [Quick Reference Cheatsheet](#10-quick-reference-cheatsheet)

---

## 1. What is WMI?

**WMI (Windows Management Instrumentation)** is Windows' built-in system for
reading hardware and software information. Think of it as a database of
everything about your computer — battery, memory, BIOS, installed software,
running processes — all queryable from PowerShell.

### Namespaces and Classes

WMI is organised into **namespaces** (folders) and **classes** (tables inside
those folders). Each class has **properties** (columns) and **instances** (rows).

```
root\cimv2           ← Standard Windows namespace
  └── Win32_Battery          ← Class: one row per battery
        ├── DeviceID          ← Property
        ├── Manufacturer
        └── EstimatedChargeRemaining

root\wmi             ← ACPI/hardware namespace
  └── BatteryStaticData
        ├── SerialNumber
        └── DesignedCapacity

root\Lenovo          ← ThinkPad-exclusive (requires SIF driver)
  └── Lenovo_Odometer
        ├── Battery_cycles
        └── Shock_events
```

### Your First WMI Query

```powershell
# Get all batteries from the standard Windows namespace
$batteries = Get-CimInstance -ClassName Win32_Battery

# Print the manufacturer of the first battery
Write-Host $batteries[0].Manufacturer
```

> **CimInstance vs WmiObject**
> TUM uses `Get-CimInstance` for most queries. One exception:
> `BatteryStaticData` **must** use `Get-WmiObject` — `Get-CimInstance`
> throws a generic failure on that class. When in doubt, try `Get-CimInstance`
> first. If it fails with no clear error, switch to `Get-WmiObject`.

---

## 2. Script Architecture

TUM follows a clear, repeatable structure. Understanding it lets you add
features confidently without breaking anything.

```
┌─────────────────────────────────────────────┐
│  1. UAC Elevation Check                     │  Runs at startup
│  2. Script-scope Variables                  │  Shared state
│  3. Helper Functions                        │  Reusable utilities
│  4. Feature Functions                       │  One per menu option
│  5. Startup Sequence                        │  Disclaimer → CIM → Cache
│  6. Main Menu Loop                          │  do { } while ($choice -ne "X")
└─────────────────────────────────────────────┘
```

### Script-scope Variables

Script-scope variables (`$script:`) are shared across all functions. They are
set once and reused, so expensive operations (like WMI queries) don't run
multiple times.

```powershell
$script:ErrorLog     = @()    # Array of error objects
$script:SystemInfo   = $null  # Cached model/BIOS info
$script:CimSession   = $null  # Shared CIM connection
$script:BatteryAlert = $null  # Cached battery health summary
```

### The CIM Session

Instead of opening a new WMI connection on every query, TUM creates **one
shared session** at startup and passes it to every `Get-CimInstance` call.

```powershell
# Created once at startup
$script:CimSession = New-CimSession

# Used in every query
Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery

# Torn down cleanly on exit
Remove-CimSession -CimSession $script:CimSession
```

If `New-CimSession` fails (e.g. WMI service issue), `$script:CimSession`
stays `$null`. Passing `$null` to `-CimSession` is harmless — PowerShell
falls back to an implicit local session automatically.

---

## 3. The Script Skeleton

Copy this as your starting point for any new TUM-style script.

```powershell
Clear-Host

# ── 1. UAC Elevation ─────────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal]
          [Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        $scriptPath = $MyInvocation.MyCommand.Path
        if (-not $scriptPath) { throw "Script path unavailable." }
        Start-Process powershell.exe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" `
            -Verb RunAs -ErrorAction Stop
        Exit
    }
    catch {
        Write-Host "Could not elevate: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press ENTER to exit"
        Exit
    }
}

# ── 2. Script-scope State ────────────────────────────────────────────────
$script:ErrorLog   = @()
$script:SystemInfo = $null
$script:CimSession = $null

# ── 3. Helper Functions ──────────────────────────────────────────────────

function Get-SafeWmiProperty {
    param([object]$Object, [string]$PropertyName, [object]$DefaultValue = "Unavailable")
    try {
        if ($Object -and $null -ne $Object.$PropertyName) { return $Object.$PropertyName }
    } catch {}
    return $DefaultValue
}

function Log-Error {
    param($Namespace, $Class, $Exception)
    $script:ErrorLog += [PSCustomObject]@{
        TimeStamp  = Get-Date
        Namespace  = $Namespace
        Class      = $Class
        Message    = $Exception.Exception.Message
        StackTrace = $Exception.Exception.StackTrace
    }
}

function Test-WmiClass {
    param([string]$Namespace = "root\cimv2", [string]$ClassName)
    try {
        return $null -ne (Get-CimClass -Namespace $Namespace -ClassName $ClassName `
                          -ErrorAction SilentlyContinue)
    } catch { return $false }
}

function New-ScriptCimSession {
    try { $script:CimSession = New-CimSession -ErrorAction Stop }
    catch { <# non-fatal — falls back to implicit sessions #> }
}

function Get-SystemInfo {
    if ($script:SystemInfo) { return $script:SystemInfo }
    $model = "Unknown"; $biosVersion = "Unknown"; $biosDate = "Unknown"
    try {
        $cs = Get-CimInstance -CimSession $script:CimSession `
                  -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) { $model = "$($cs.Manufacturer) $($cs.Model)".Trim() }
    } catch {}
    try {
        $bios = Get-CimInstance -CimSession $script:CimSession `
                    -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) {
            $biosVersion = $bios.SMBIOSBIOSVersion
            if ($bios.ReleaseDate) {
                $biosDate = ([System.Management.ManagementDateTimeConverter]::
                             ToDateTime($bios.ReleaseDate)).ToString("yyyy-MM-dd")
            }
        }
    } catch {}
    $script:SystemInfo = [PSCustomObject]@{
        Model = $model; BIOSVersion = $biosVersion; BIOSDate = $biosDate
    }
    return $script:SystemInfo
}

function Show-Header {
    Clear-Host
    $si = Get-SystemInfo
    Write-Host "======================================="
    Write-Host "  MY CUSTOM DIAGNOSTIC TOOL  v1.0"
    Write-Host "======================================="
    Write-Host "Device : $env:COMPUTERNAME"
    Write-Host "Model  : $($si.Model)"
    Write-Host "BIOS   : $($si.BIOSVersion)  ($($si.BIOSDate))"
    Write-Host ""
}

# ── 4. Feature Functions ─────────────────────────────────────────────────
# (Add your functions here — see Section 7)

# ── 5. Startup Sequence ──────────────────────────────────────────────────
New-ScriptCimSession

# ── 6. Main Menu Loop ────────────────────────────────────────────────────
do {
    Show-Header
    Write-Host "1. My First Feature"
    Write-Host "2. Exit"
    Write-Host ""
    $choice = Read-Host "Select option"

    switch ($choice) {
        "1" { MyFirstFeature }
    }
} while ($choice -ne "2")

# Teardown
if ($script:CimSession) {
    Remove-CimSession -CimSession $script:CimSession -ErrorAction SilentlyContinue
    $script:CimSession = $null
}
Clear-Host
Write-Host "Exiting..."
```

---

## 4. WMI Query Patterns and Fallback Chains

Real hardware data is unreliable — a class may exist on one ThinkPad model
but not another. TUM handles this with **fallback chains**: try the richest
source first, fall back to simpler sources if it fails.

### Pattern 1 — Single Class Query

Use this when there is one reliable source and failure just means "not available".

```powershell
try {
    $battery = Get-CimInstance -CimSession $script:CimSession `
                   -ClassName Win32_Battery -ErrorAction SilentlyContinue

    if ($battery) {
        Write-Host "Manufacturer : $($battery.Manufacturer)"
        Write-Host "Charge       : $($battery.EstimatedChargeRemaining)%"
    }
    else {
        Write-Host "No battery found." -ForegroundColor Yellow
    }
}
catch {
    Log-Error "root\cimv2" "Win32_Battery" $_
    Write-Host "Query failed." -ForegroundColor Red
}
```

### Pattern 2 — Availability Check Before Query

Always check if a class exists before querying it, especially for
`root\wmi` and `root\Lenovo` classes that may not be present.

```powershell
if (-not (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryStaticData")) {
    Write-Host "BatteryStaticData not available on this system." -ForegroundColor Yellow
    Read-Host "Press ENTER"
    return
}

# Safe to query now
$data = Get-WmiObject -Namespace root\wmi -Class BatteryStaticData
```

### Pattern 3 — Two-Source Fallback Chain

This is the core TUM pattern. Try the best source first; if it fails or
returns nothing, fall back to the next source.

```powershell
$designCapacity = $null

# Source 1: BatteryStaticData (most accurate, requires ACPI support)
# NOTE: Must use Get-WmiObject here — Get-CimInstance fails on this class
try {
    $static = @(Get-WmiObject -Namespace root\wmi `
                    -Class BatteryStaticData -ErrorAction SilentlyContinue)
    if ($static -and $static.Count -gt 0 -and $static[0].DesignedCapacity -gt 0) {
        $designCapacity = [string]$static[0].DesignedCapacity
    }
} catch {}

# Source 2: powercfg battery report (universal Windows fallback)
if (-not $designCapacity) {
    try {
        $xmlPath = Join-Path $env:TEMP "battery_$PID.xml"
        $null = & powercfg /batteryreport /XML /OUTPUT $xmlPath 2>$null
        if (Test-Path $xmlPath) {
            [xml]$report = Get-Content $xmlPath -ErrorAction SilentlyContinue
            Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
            $val = @($report.BatteryReport.Batteries.Battery)[0].DesignCapacity
            if ($val -and [int64]$val -gt 0) { $designCapacity = [string]$val }
        }
    } catch {}
}

Write-Host "Design Capacity: $(if ($designCapacity) { "$designCapacity mWh" } else { "Unavailable" })"
```

### Pattern 4 — Three-Source Fallback Chain (Lenovo → ACPI → Win32)

Used in TUM's Full Charge Capacity and Comprehensive Battery Analysis functions.

```powershell
$usedPrimarySource = $false

# Source 1: Lenovo_Battery (richest — requires SIF)
if (Test-LenovoNamespace) {
    try {
        $lb = Get-CimInstance -CimSession $script:CimSession `
                  -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop
        if ($lb) {
            $usedPrimarySource = $true
            Write-Host "Source: Lenovo_Battery" -ForegroundColor DarkGray
            # ... display data from $lb
        }
    } catch {
        Log-Error "root\Lenovo" "Lenovo_Battery" $_
    }
}

# Source 2: ACPI BatteryFullChargedCapacity
if (-not $usedPrimarySource) {
    try {
        $acpi = @(Get-WmiObject -Namespace root\wmi `
                      -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
        if ($acpi -and $acpi.Count -gt 0) {
            $usedPrimarySource = $true
            Write-Host "Source: ACPI fallback" -ForegroundColor DarkGray
            # ... display data from $acpi
        }
    } catch {}
}

# Source 3: Win32_Battery (last resort)
if (-not $usedPrimarySource) {
    try {
        $win32 = Get-CimInstance -CimSession $script:CimSession `
                     -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($win32) {
            Write-Host "Source: Win32_Battery (limited data)" -ForegroundColor DarkGray
            # ... display limited data
        }
    } catch {}
}
```

### Checking the Lenovo Namespace

Before querying any `root\Lenovo` class, always verify the namespace exists:

```powershell
function Test-LenovoNamespace {
    try {
        $ns = Get-CimInstance -CimSession $script:CimSession `
                  -Namespace root -ClassName __Namespace -ErrorAction SilentlyContinue |
              Where-Object Name -eq "Lenovo"
        return $null -ne $ns
    }
    catch {
        try {
            $null = Get-CimClass -Namespace root\Lenovo `
                        -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue
            return $true
        }
        catch { return $false }
    }
}
```

### Iterating Multiple Batteries

Many of TUM's battery functions need to handle dual-battery ThinkPads.
Always wrap the results in `@()` to force an array — otherwise a single
result comes back as a plain object, not an array, and your loop breaks.

```powershell
# @() ensures $batteries is always an array, even with 1 result
$batteries = @(Get-WmiObject -Namespace root\wmi `
                   -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

$index = 0
foreach ($bat in $batteries) {
    Write-Host "Battery #$index"
    Write-Host "Full Charge: $($bat.FullChargedCapacity) mWh"
    $index++
}
```

---

## 5. Safe Property Access

WMI objects sometimes return `$null` for a property, or the property may not
exist at all on certain firmware versions. Accessing a missing property
directly will throw an error and crash your script.

TUM solves this with the `Get-SafeWmiProperty` helper:

```powershell
function Get-SafeWmiProperty {
    param(
        [object]$Object,
        [string]$PropertyName,
        [object]$DefaultValue = "Unavailable"   # returned if property is null/missing
    )
    try {
        if ($Object -and $null -ne $Object.$PropertyName) {
            return $Object.$PropertyName
        }
    }
    catch { <# silently return default #> }
    return $DefaultValue
}
```

**Usage comparison:**

```powershell
# ❌ UNSAFE — crashes if Manufacturer is null or property doesn't exist
Write-Host $battery.Manufacturer

# ✅ SAFE — returns "Unavailable" instead of crashing
Write-Host (Get-SafeWmiProperty -Object $battery -PropertyName "Manufacturer")

# ✅ SAFE — custom default value
Write-Host (Get-SafeWmiProperty -Object $battery -PropertyName "Manufacturer" `
                                -DefaultValue "Unknown")
```

### Parsing Numeric Values Safely

WMI properties that look like numbers are sometimes returned as strings, and
sometimes not returned at all. Always use `TryParse` instead of casting:

```powershell
# ❌ UNSAFE — throws if value is null or non-numeric
$gb = [int64]$module.Capacity / 1GB

# ✅ SAFE — TryParse returns false instead of throwing
[int64]$capacityBytes = 0
$capacityRaw = Get-SafeWmiProperty -Object $module -PropertyName "Capacity"

if ($capacityRaw -and $capacityRaw -ne "Unavailable" -and
    [int64]::TryParse($capacityRaw.ToString(), [ref]$capacityBytes) -and
    $capacityBytes -gt 0) {
    $capacityGB = [math]::Round($capacityBytes / 1GB, 2)
    Write-Host "Capacity: $capacityGB GB"
}
else {
    Write-Host "Capacity: Unavailable" -ForegroundColor Yellow
}
```

### Parsing Lenovo Decimal Strings

`Lenovo_Battery` reports capacity as localised decimal strings like `"39,96Wh"`
(comma as decimal separator in some locales). Strip non-numeric characters and
force the invariant culture when parsing:

```powershell
$raw = Get-SafeWmiProperty -Object $lb -PropertyName "DesignCapacity"  # e.g. "39,96Wh"

$cleaned = ($raw -replace '[^0-9,\.]', '') -replace ',', '.'  # "39.96"

[double]$value = 0
$parsed = [double]::TryParse(
    $cleaned,
    [System.Globalization.NumberStyles]::Any,
    [System.Globalization.CultureInfo]::InvariantCulture,
    [ref]$value
)

if ($parsed -and $value -gt 0) {
    Write-Host "Capacity: $value Wh"
}
```

---

## 6. Error Logging

TUM never lets an error silently disappear. Every caught exception is stored
in `$script:ErrorLog` with full context for later review.

### The Error Log Structure

```powershell
$script:ErrorLog += [PSCustomObject]@{
    TimeStamp        = Get-Date
    Namespace        = "root\wmi"         # which WMI namespace was being queried
    Class            = "BatteryStaticData" # which class failed
    Message          = $_.Exception.Message
    StackTrace       = $_.Exception.StackTrace
    ScriptName       = $_.InvocationInfo.ScriptName
    ScriptLineNumber = $_.InvocationInfo.ScriptLineNumber
    Context          = $null              # optional: extra context string
    FullException    = $_                 # full exception object, if needed
}
```

### Using the Log-Error Helper

```powershell
function Log-Error {
    param($Namespace, $Class, $Exception)
    $script:ErrorLog += [PSCustomObject]@{
        TimeStamp        = Get-Date
        Namespace        = $Namespace
        Class            = $Class
        Message          = $Exception.Exception.Message
        StackTrace       = $Exception.Exception.StackTrace
        ScriptName       = $Exception.InvocationInfo.ScriptName
        ScriptLineNumber = $Exception.InvocationInfo.ScriptLineNumber
        Context          = $null
        FullException    = $Exception
    }
}

# In your feature function:
try {
    $data = Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop
}
catch {
    Log-Error "root\cimv2" "Win32_Battery" $_
    Write-Host "Battery query failed." -ForegroundColor Red
}
```

### Showing the Error Count in the Header

When errors exist, the header notifies the user automatically:

```powershell
function Show-Header {
    # ... other header content ...
    if ($script:ErrorLog.Count -gt 0) {
        Write-Host "Errors : " -NoNewline
        Write-Host "$($script:ErrorLog.Count) error(s) logged  [Option X to view]" `
            -ForegroundColor Red
    }
    Write-Host ""
}
```

### Building an Error Log Viewer Function

```powershell
function Show-ErrorLog {
    Show-Header
    Write-Host "[ Error Log ]"
    Write-Host ""

    if ($script:ErrorLog.Count -eq 0) {
        Write-Host "No errors this session." -ForegroundColor Green
        Read-Host "Press ENTER"
        return
    }

    $i = 0
    foreach ($err in $script:ErrorLog) {
        $i++
        Write-Host ("=" * 50) -ForegroundColor DarkGray
        Write-Host "Error #$i" -ForegroundColor Red
        Write-Host "Time      : $($err.TimeStamp)"
        Write-Host "Namespace : $($err.Namespace)"
        Write-Host "Class     : $($err.Class)"
        Write-Host "Message   : " -NoNewline
        Write-Host $err.Message -ForegroundColor Yellow
        if ($err.ScriptLineNumber) {
            Write-Host "Line      : $($err.ScriptLineNumber)" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    Read-Host "Press ENTER"
}
```

---

## 7. Adding a New Menu Option

Every menu option in TUM follows the same four-step pattern.

### Step 1 — Write the Feature Function

Your function always starts with `Show-Header` and ends with `Read-Host "Press ENTER"`.

```powershell
function Show-ProcessorInfo {
    Show-Header
    Write-Host "[ Processor Information ]"
    Write-Host ""

    try {
        if (-not (Test-WmiClass -ClassName "Win32_Processor")) {
            Write-Host "Win32_Processor not available." -ForegroundColor Yellow
            Read-Host "Press ENTER"
            return
        }

        $cpus = @(Get-CimInstance -CimSession $script:CimSession `
                      -ClassName Win32_Processor -ErrorAction SilentlyContinue)

        if (-not $cpus -or $cpus.Count -eq 0) {
            Write-Host "No processor data found." -ForegroundColor Yellow
        }
        else {
            $index = 0
            foreach ($cpu in $cpus) {
                Write-Host "CPU #$index"
                Write-Host ("-" * 50)
                Write-Host "Name         : $(Get-SafeWmiProperty -Object $cpu -PropertyName 'Name')"
                Write-Host "Cores        : $(Get-SafeWmiProperty -Object $cpu -PropertyName 'NumberOfCores')"
                Write-Host "Threads      : $(Get-SafeWmiProperty -Object $cpu -PropertyName 'ThreadCount')"
                Write-Host "Max Speed    : $(Get-SafeWmiProperty -Object $cpu -PropertyName 'MaxClockSpeed') MHz"
                Write-Host ""
                $index++
            }
        }
    }
    catch {
        Log-Error "root\cimv2" "Win32_Processor" $_
        Write-Host "Failed to query processor information." -ForegroundColor Red
    }

    Read-Host "Press ENTER"
}
```

### Step 2 — Add it to the Menu

```powershell
do {
    Show-Header
    Write-Host "1. Battery Static Data"
    Write-Host "2. Processor Info"       # ← add your new option here
    Write-Host "3. Exit"
    Write-Host ""
    $choice = Read-Host "Select option"

    switch ($choice) {
        "1" { BatteryStaticData }
        "2" { Show-ProcessorInfo }        # ← call your function here
    }
} while ($choice -ne "3")
```

### Step 3 — Add Colour for Severity (Optional)

If your feature has a severity level, use this colour scale — it matches
TUM's battery severity system and keeps the visual language consistent.

```powershell
$severityColor = switch ($severity) {
    0 { "Green"   }   # Good / OK
    1 { "Cyan"    }   # Warning / Attention
    2 { "Yellow"  }   # Fair / Degraded
    3 { "Magenta" }   # Poor / Serious
    4 { "Red"     }   # Critical / Replace
    default { "White" }
}

Write-Host "Status : " -NoNewline
Write-Host $statusText -ForegroundColor $severityColor
```

### Step 4 — Return Data as an Object (Optional but Recommended)

If you need the data elsewhere (e.g. in the header, or in an export report),
separate the **data retrieval** from the **display** — just like TUM does with
`Get-BatteryAlertState` and `Get-MemoryState`.

```powershell
# Data function — returns a PSCustomObject, no Write-Host
function Get-ProcessorState {
    $result = [PSCustomObject]@{
        Name      = "Unknown"
        Cores     = 0
        Threads   = 0
        MaxSpeedMHz = 0
        Available = $false
    }

    try {
        $cpu = Get-CimInstance -CimSession $script:CimSession `
                   -ClassName Win32_Processor -ErrorAction SilentlyContinue |
               Select-Object -First 1

        if ($cpu) {
            $result.Name        = Get-SafeWmiProperty -Object $cpu -PropertyName "Name"
            $result.Cores       = Get-SafeWmiProperty -Object $cpu -PropertyName "NumberOfCores" -DefaultValue 0
            $result.Threads     = Get-SafeWmiProperty -Object $cpu -PropertyName "ThreadCount" -DefaultValue 0
            $result.MaxSpeedMHz = Get-SafeWmiProperty -Object $cpu -PropertyName "MaxClockSpeed" -DefaultValue 0
            $result.Available   = $true
        }
    }
    catch {
        Log-Error "root\cimv2" "Win32_Processor" $_
    }

    return $result
}

# Display function — calls the data function and shows results
function Show-ProcessorInfo {
    Show-Header
    Write-Host "[ Processor Information ]"
    Write-Host ""

    $proc = Get-ProcessorState

    if (-not $proc.Available) {
        Write-Host "Processor data unavailable." -ForegroundColor Yellow
    }
    else {
        Write-Host "Name    : $($proc.Name)"
        Write-Host "Cores   : $($proc.Cores)"
        Write-Host "Threads : $($proc.Threads)"
        Write-Host "Speed   : $($proc.MaxSpeedMHz) MHz"
    }

    Read-Host "Press ENTER"
}
```

---

## 8. The Header and State System

### Show-Header

`Show-Header` is called at the start of every feature function. It clears the
screen and prints consistent device info on every page so the user always knows
where they are.

```powershell
function Show-Header {
    Clear-Host
    $si = Get-SystemInfo   # uses cached value after first call
    Write-Host "======================================="
    Write-Host "  MY TOOL  v1.0"
    Write-Host "======================================="
    Write-Host "Device : $env:COMPUTERNAME"
    Write-Host "Model  : $($si.Model)"
    Write-Host "BIOS   : $($si.BIOSVersion)  ($($si.BIOSDate))"

    # Optional: show any cached alert state here
    # if ($script:MyAlert -and $script:MyAlert.Available) { ... }

    if ($script:ErrorLog.Count -gt 0) {
        Write-Host "Errors : " -NoNewline
        Write-Host "$($script:ErrorLog.Count) error(s)" -ForegroundColor Red
    }
    Write-Host ""
}
```

### Caching Expensive Data at Startup

If a data query is slow or used on every screen, run it once at startup and
cache the result in a `$script:` variable. Follow the TUM pattern:

```powershell
function Get-MyAlertState {
    # Return cached value if already populated
    if ($script:MyAlert) { return $script:MyAlert }

    $result = [PSCustomObject]@{
        Available = $false
        Summary   = ""
        Severity  = 0
    }

    try {
        # ... your query here ...
        $result.Available = $true
        $result.Summary   = "Everything OK"
        $result.Severity  = 0
    }
    catch {}

    $script:MyAlert = $result
    return $script:MyAlert
}

# At startup, before the menu loop:
$null = Get-MyAlertState
```

---

## 9. Full Working Example

A complete, self-contained script with two features — processor info and
disk info — built entirely on the patterns in this guide.

```powershell
Clear-Host

# ── UAC Elevation ────────────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal]
          [Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        $p = $MyInvocation.MyCommand.Path
        if (-not $p) { throw "Path unavailable." }
        Start-Process powershell.exe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$p`"" `
            -Verb RunAs -ErrorAction Stop
        Exit
    } catch {
        Write-Host "Elevation failed: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press ENTER to exit"; Exit
    }
}

# ── State ────────────────────────────────────────────────────────────────
$script:ErrorLog   = @()
$script:SystemInfo = $null
$script:CimSession = $null

# ── Helpers ──────────────────────────────────────────────────────────────
function Get-SafeWmiProperty {
    param([object]$Object, [string]$PropertyName, [object]$DefaultValue = "Unavailable")
    try { if ($Object -and $null -ne $Object.$PropertyName) { return $Object.$PropertyName } }
    catch {}
    return $DefaultValue
}

function Log-Error {
    param($Namespace, $Class, $Exception)
    $script:ErrorLog += [PSCustomObject]@{
        TimeStamp = Get-Date; Namespace = $Namespace; Class = $Class
        Message   = $Exception.Exception.Message
    }
}

function Test-WmiClass {
    param([string]$Namespace = "root\cimv2", [string]$ClassName)
    try { return $null -ne (Get-CimClass -Namespace $Namespace -ClassName $ClassName `
                            -ErrorAction SilentlyContinue) }
    catch { return $false }
}

function New-ScriptCimSession {
    try { $script:CimSession = New-CimSession -ErrorAction Stop } catch {}
}

function Get-SystemInfo {
    if ($script:SystemInfo) { return $script:SystemInfo }
    $model = "Unknown"; $bv = "Unknown"; $bd = "Unknown"
    try {
        $cs = Get-CimInstance -CimSession $script:CimSession `
                  -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) { $model = "$($cs.Manufacturer) $($cs.Model)".Trim() }
    } catch {}
    try {
        $b = Get-CimInstance -CimSession $script:CimSession `
                 -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($b) {
            $bv = $b.SMBIOSBIOSVersion
            if ($b.ReleaseDate) {
                $bd = ([System.Management.ManagementDateTimeConverter]::
                        ToDateTime($b.ReleaseDate)).ToString("yyyy-MM-dd")
            }
        }
    } catch {}
    $script:SystemInfo = [PSCustomObject]@{ Model=$model; BIOSVersion=$bv; BIOSDate=$bd }
    return $script:SystemInfo
}

function Show-Header {
    Clear-Host
    $si = Get-SystemInfo
    Write-Host "======================================="
    Write-Host "  MY DIAGNOSTIC TOOL  v1.0"
    Write-Host "======================================="
    Write-Host "Device : $env:COMPUTERNAME"
    Write-Host "Model  : $($si.Model)"
    Write-Host "BIOS   : $($si.BIOSVersion)  ($($si.BIOSDate))"
    if ($script:ErrorLog.Count -gt 0) {
        Write-Host "Errors : " -NoNewline
        Write-Host "$($script:ErrorLog.Count) error(s)  [Option 3 to view]" -ForegroundColor Red
    }
    Write-Host ""
}

# ── Feature: Processor Info ──────────────────────────────────────────────
function Show-ProcessorInfo {
    Show-Header
    Write-Host "[ Processor Information ]"
    Write-Host ""

    try {
        $cpus = @(Get-CimInstance -CimSession $script:CimSession `
                      -ClassName Win32_Processor -ErrorAction SilentlyContinue)
        if (-not $cpus -or $cpus.Count -eq 0) {
            Write-Host "No processor data found." -ForegroundColor Yellow
        } else {
            $i = 0
            foreach ($cpu in $cpus) {
                Write-Host "CPU #$i"
                Write-Host ("-" * 40)
                Write-Host "Name      : $(Get-SafeWmiProperty $cpu 'Name')"
                Write-Host "Cores     : $(Get-SafeWmiProperty $cpu 'NumberOfCores')"
                Write-Host "Threads   : $(Get-SafeWmiProperty $cpu 'ThreadCount')"
                Write-Host "Max Speed : $(Get-SafeWmiProperty $cpu 'MaxClockSpeed') MHz"
                Write-Host ""
                $i++
            }
        }
    } catch {
        Log-Error "root\cimv2" "Win32_Processor" $_
        Write-Host "Failed to query processor." -ForegroundColor Red
    }

    Read-Host "Press ENTER"
}

# ── Feature: Disk Info ───────────────────────────────────────────────────
function Show-DiskInfo {
    Show-Header
    Write-Host "[ Disk Information ]"
    Write-Host ""

    try {
        $disks = @(Get-CimInstance -CimSession $script:CimSession `
                       -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue)
        if (-not $disks -or $disks.Count -eq 0) {
            Write-Host "No disks found." -ForegroundColor Yellow
        } else {
            $i = 0
            foreach ($disk in $disks) {
                [int64]$sizeBytes = 0
                $sizeRaw = Get-SafeWmiProperty $disk "Size"
                if ($sizeRaw -ne "Unavailable") {
                    [int64]::TryParse($sizeRaw.ToString(), [ref]$sizeBytes) | Out-Null
                }
                $sizeGB = if ($sizeBytes -gt 0) {
                    "$([math]::Round($sizeBytes / 1GB, 1)) GB"
                } else { "Unavailable" }

                Write-Host "Disk #$i"
                Write-Host ("-" * 40)
                Write-Host "Model  : $(Get-SafeWmiProperty $disk 'Model')"
                Write-Host "Size   : $sizeGB"
                Write-Host "Serial : $(Get-SafeWmiProperty $disk 'SerialNumber')"
                Write-Host ""
                $i++
            }
        }
    } catch {
        Log-Error "root\cimv2" "Win32_DiskDrive" $_
        Write-Host "Failed to query disks." -ForegroundColor Red
    }

    Read-Host "Press ENTER"
}

# ── Feature: Error Log Viewer ────────────────────────────────────────────
function Show-ErrorLog {
    Show-Header
    Write-Host "[ Error Log ]"
    Write-Host ""
    if ($script:ErrorLog.Count -eq 0) {
        Write-Host "No errors this session." -ForegroundColor Green
        Read-Host "Press ENTER"; return
    }
    $i = 0
    foreach ($err in $script:ErrorLog) {
        $i++
        Write-Host "Error #$i" -ForegroundColor Red
        Write-Host "Time      : $($err.TimeStamp)"
        Write-Host "Class     : $($err.Namespace) \ $($err.Class)"
        Write-Host "Message   : $($err.Message)" -ForegroundColor Yellow
        Write-Host ""
    }
    Read-Host "Press ENTER"
}

# ── Startup ──────────────────────────────────────────────────────────────
New-ScriptCimSession

# ── Menu Loop ────────────────────────────────────────────────────────────
do {
    Show-Header
    Write-Host "1. Processor Info"
    Write-Host "2. Disk Info"
    Write-Host "3. Error Log"
    Write-Host "4. Exit"
    Write-Host ""
    $choice = Read-Host "Select option"

    switch ($choice) {
        "1" { Show-ProcessorInfo }
        "2" { Show-DiskInfo }
        "3" { Show-ErrorLog }
    }
} while ($choice -ne "4")

if ($script:CimSession) {
    Remove-CimSession -CimSession $script:CimSession -ErrorAction SilentlyContinue
}
Clear-Host
Write-Host "Exiting..."
```

---

## 10. Quick Reference Cheatsheet

### Common WMI Classes

| Class | Namespace | Key Properties | Note |
|---|---|---|---|
| `Win32_Battery` | `root\cimv2` | Manufacturer, EstimatedChargeRemaining, Status | Basic battery info |
| `Win32_PhysicalMemory` | `root\cimv2` | Capacity, Speed, Manufacturer, DeviceLocator | RAM modules |
| `Win32_Processor` | `root\cimv2` | Name, NumberOfCores, ThreadCount, MaxClockSpeed | CPU |
| `Win32_DiskDrive` | `root\cimv2` | Model, Size, SerialNumber | Physical disks |
| `Win32_ComputerSystem` | `root\cimv2` | Manufacturer, Model | Device model string |
| `Win32_BIOS` | `root\cimv2` | SMBIOSBIOSVersion, ReleaseDate | BIOS version |
| `BatteryStaticData` | `root\wmi` | DesignedCapacity, SerialNumber, ManufactureName | **Use Get-WmiObject** |
| `BatteryFullChargedCapacity` | `root\wmi` | FullChargedCapacity | Current max charge |
| `Lenovo_Odometer` | `root\Lenovo` | Battery_cycles, Shock_events, Thermal_events | Requires SIF |
| `Lenovo_Battery` | `root\Lenovo` | All battery fields | Requires SIF |
| `Lenovo_BiosSetting` | `root\wmi` | CurrentSetting | BIOS settings |

### Function Template

```powershell
function Show-MyFeature {
    Show-Header
    Write-Host "[ Feature Title ]"
    Write-Host ""

    try {
        # 1. Check class availability
        if (-not (Test-WmiClass -ClassName "Win32_Something")) {
            Write-Host "Not available." -ForegroundColor Yellow
            Read-Host "Press ENTER"; return
        }

        # 2. Query
        $items = @(Get-CimInstance -CimSession $script:CimSession `
                       -ClassName Win32_Something -ErrorAction SilentlyContinue)

        # 3. Guard empty result
        if (-not $items -or $items.Count -eq 0) {
            Write-Host "No data found." -ForegroundColor Yellow
        } else {
            # 4. Display with safe property access
            foreach ($item in $items) {
                Write-Host "Property : $(Get-SafeWmiProperty $item 'PropertyName')"
            }
        }
    } catch {
        Log-Error "root\cimv2" "Win32_Something" $_
        Write-Host "Query failed." -ForegroundColor Red
    }

    Read-Host "Press ENTER"
}
```

### Colour Conventions

| Colour | Used for |
|---|---|
| `White` | Normal / OK output |
| `Cyan` | Section headers, attention items |
| `Green` | Success, good status |
| `Yellow` | Warning, not available, soft failure |
| `Magenta` | Poor / serious degradation |
| `Red` | Critical / hard failure |
| `DarkGray` | Metadata, source labels, secondary info |

### Key Rules

- Always wrap multi-instance queries in `@()` to force array
- Always use `Get-WmiObject` (not `Get-CimInstance`) for `BatteryStaticData`
- Always call `Test-WmiClass` before querying optional classes
- Always call `Test-LenovoNamespace` before any `root\Lenovo` query
- Always use `TryParse` instead of direct casts for numeric WMI properties
- Always pass `-CimSession $script:CimSession` to `Get-CimInstance`
- Always start feature functions with `Show-Header`
- Always end feature functions with `Read-Host "Press ENTER"`
- Never let a `catch` block be empty — always call `Log-Error` inside it