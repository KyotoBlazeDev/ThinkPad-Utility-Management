# ===============================================
#  THINKPAD DEPLOYMENT CHECK
#  Standalone buyer-facing readiness tool
#  Part of ThinkPad Utility Management
#  github.com/KyotoBlazeDev/ThinkPad-Utility-Management
# ===============================================
#
# Runs a single deployment readiness check and shows
# a clear READY / REVIEW NEEDED / NOT READY verdict.
# Saves a one-page report to the Desktop automatically.
# No menu, no setup required.

# ── Administrator elevation ───────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        $scriptPath = $MyInvocation.MyCommand.Path
        if (-not $scriptPath) { throw "No path" }
        Start-Process -FilePath "powershell.exe" `
                      -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" `
                      -Verb RunAs -ErrorAction Stop
        Exit
    }
    catch {
        Clear-Host
        Write-Host ""
        Write-Host "  ADMINISTRATOR ACCESS REQUIRED" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Right-click this script and choose 'Run as Administrator'." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "  Press ENTER to exit"
        Exit
    }
}

Clear-Host

# ── Shared CIM session ────────────────────────────────────────────────────────
$script:CimSession = $null
try { $script:CimSession = New-CimSession -ErrorAction Stop } catch {}

# ══════════════════════════════════════════════════════════════════════════════
# Helper functions (self-contained — no dependency on main script)
# ══════════════════════════════════════════════════════════════════════════════

function Test-WmiClass {
    param([string]$Namespace = "root\cimv2", [string]$ClassName)
    try { return $null -ne (Get-CimClass -Namespace $Namespace -ClassName $ClassName -ErrorAction SilentlyContinue) }
    catch { return $false }
}

function Get-SafeWmiProperty {
    param([object]$Object, [string]$PropertyName, [object]$DefaultValue = "Unavailable")
    try { if ($Object -and $null -ne $Object.$PropertyName) { return $Object.$PropertyName } }
    catch {}
    return $DefaultValue
}

function Test-LenovoNamespace {
    try {
        $ns = Get-CimInstance -CimSession $script:CimSession -Namespace root -ClassName __Namespace `
                  -ErrorAction SilentlyContinue | Where-Object Name -eq "Lenovo"
        return $null -ne $ns
    }
    catch {
        try { $null = Get-CimClass -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue; return $true }
        catch { return $false }
    }
}

function Get-DesignCapacity {
    param([int]$Index = 0)
    try {
        $items = @(Get-WmiObject -Namespace root\wmi -Class BatteryStaticData -ErrorAction SilentlyContinue)
        if ($items -and $Index -lt $items.Count -and $items[$Index].DesignedCapacity -gt 0) {
            return [string]$items[$Index].DesignedCapacity
        }
    } catch {}
    try {
        $xmlPath = Join-Path $env:TEMP "dc_check_$PID.xml"
        $null = & powercfg /batteryreport /XML /OUTPUT $xmlPath 2>$null
        if (Test-Path $xmlPath) {
            [xml]$r = Get-Content $xmlPath -ErrorAction SilentlyContinue
            Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
            $bats = @($r.BatteryReport.Batteries.Battery)
            if ($bats -and $Index -lt $bats.Count -and $bats[$Index].DesignCapacity -gt 0) {
                return [string]$bats[$Index].DesignCapacity
            }
        }
    } catch {}
    return $null
}

function Get-MemoryState {
    $result = [PSCustomObject]@{ TotalRAMGB = 0; ModuleCount = 0; SpeedMismatch = $false; Modules = @() }
    try {
        $modules = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue)
        if (-not $modules -or $modules.Count -eq 0) { return $result }
        $result.ModuleCount = $modules.Count
        $speeds = @(); $totalBytes = 0
        foreach ($m in $modules) {
            [int64]$cap = 0
            if ([int64]::TryParse((Get-SafeWmiProperty $m "Capacity").ToString(), [ref]$cap)) { $totalBytes += $cap }
            [int32]$spd = 0
            if ([int32]::TryParse((Get-SafeWmiProperty $m "Speed").ToString(), [ref]$spd)) { $speeds += $spd }
        }
        if ($totalBytes -gt 0) { $result.TotalRAMGB = [math]::Round($totalBytes / 1GB, 2) }
        if ($speeds.Count -gt 1 -and ($speeds | Select-Object -Unique).Count -gt 1) { $result.SpeedMismatch = $true }
    } catch {}
    return $result
}

function Get-CdrtData {
    $result = [PSCustomObject]@{ Available = $false; ShockEvents = $null; ThermalEvents = $null }
    try {
        if (-not (Test-LenovoNamespace)) { return $result }
        $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo `
                   -ClassName Lenovo_Odometer -ErrorAction Stop
        $props = $odo.CimInstanceProperties | Select-Object -ExpandProperty Name
        $shockProp   = $props | Where-Object { $_ -match "(?i)shock" }   | Select-Object -First 1
        $thermalProp = $props | Where-Object { $_ -match "(?i)thermal" } | Select-Object -First 1
        if ($shockProp -or $thermalProp) {
            $result.Available     = $true
            $result.ShockEvents   = if ($shockProp)   { $odo.$shockProp   } else { $null }
            $result.ThermalEvents = if ($thermalProp) { $odo.$thermalProp } else { $null }
        }
    } catch {}
    return $result
}

function Get-BatterySerialMismatch {
    <#
    .SYNOPSIS
    Compares battery serial numbers reported by Lenovo_Battery (EC firmware)
    and Win32_Battery (ACPI layer) to detect barcode/serial mismatches.

    .DESCRIPTION
    Genuine OEM batteries report consistent serial numbers across both the
    Lenovo EC WMI interface (root\Lenovo\Lenovo_Battery) and the standard
    Windows ACPI battery interface (Win32_Battery). A mismatch between the
    two may indicate a third-party replacement, a refurbished cell, or a
    firmware/label inconsistency worth noting before deployment.

    This check replaces the previous manufacturer name allowlist approach.
    Under EU Battery Regulation (effective Feb 18 2027), third-party battery
    replacements are a consumer right — the mismatch is surfaced as WARN
    (informational flag) rather than FAIL, so buyers can make an informed
    decision without the tool implying the battery is unsafe or defective.

    Returns a PSCustomObject:
      Available   : $true if both sources were readable
      Mismatch    : $true if serials differ
      ECSerial    : serial from Lenovo_Battery (EC layer)
      ACPISerial  : serial from Win32_Battery (ACPI layer)
      Detail      : human-readable summary string
    #>
    $result = [PSCustomObject]@{
        Available  = $false
        Mismatch   = $false
        ECSerial   = $null
        ACPISerial = $null
        Detail     = "Unavailable"
    }

    # ── Source 1: EC layer via Lenovo_Battery ─────────────────────────────
    $ecSerial = $null
    try {
        if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
            $lbs = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo `
                         -ClassName Lenovo_Battery -ErrorAction Stop)
            if ($lbs -and $lbs.Count -gt 0) {
                $raw = Get-SafeWmiProperty -Object $lbs[0] -PropertyName "BatteryID" -DefaultValue $null
                if ($raw -and $raw -ne "Unavailable") {
                    $ecSerial = $raw.ToString().Trim()
                }
            }
        }
    } catch {}

    # ── Source 2: ACPI layer via Win32_Battery ────────────────────────────
    $acpiSerial = $null
    try {
        $wb = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery `
                  -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($wb) {
            $raw = Get-SafeWmiProperty -Object $wb -PropertyName "DeviceID" -DefaultValue $null
            if ($raw -and $raw -ne "Unavailable") {
                $acpiSerial = $raw.ToString().Trim()
            }
        }
    } catch {}

    if (-not $ecSerial -and -not $acpiSerial) { return $result }

    $result.Available  = $true
    $result.ECSerial   = $ecSerial
    $result.ACPISerial = $acpiSerial

    # Normalise before comparison — strip whitespace, compare case-insensitively
    $ecNorm   = if ($ecSerial)   { $ecSerial.ToUpper()   -replace '\s','' } else { "" }
    $acpiNorm = if ($acpiSerial) { $acpiSerial.ToUpper() -replace '\s','' } else { "" }

    if ($ecNorm -and $acpiNorm -and $ecNorm -ne $acpiNorm) {
        $result.Mismatch = $true
        $result.Detail   = "EC: $ecSerial  |  ACPI: $acpiSerial"
    } elseif ($ecNorm -and $acpiNorm) {
        $result.Detail = $ecSerial
    } elseif ($ecNorm) {
        $result.Detail = "$ecSerial (ACPI not available)"
    } else {
        $result.Detail = "$acpiSerial (EC not available)"
    }

    return $result
}

function Get-SystemInfo {
    $model = "Unknown"; $biosVersion = "Unknown"; $biosDate = "Unknown"
    try {
        $cs = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) { $model = "$($cs.Manufacturer) $($cs.Model)".Trim() }
    } catch {}
    try {
        $b = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($b) {
            $biosVersion = $b.SMBIOSBIOSVersion
            if ($b.ReleaseDate) {
                $biosDate = ([System.Management.ManagementDateTimeConverter]::ToDateTime($b.ReleaseDate)).ToString("yyyy-MM-dd")
            }
        }
    } catch {}
    $env = "Standalone"
    try {
        $cs2 = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs2 -and $cs2.PartOfDomain) { $env = "Domain-Joined$(if ($cs2.Domain) { " ($($cs2.Domain))" })" }
    } catch {}
    return [PSCustomObject]@{ Model = $model; BIOSVersion = $biosVersion; BIOSDate = $biosDate; Environment = $env }
}

# ══════════════════════════════════════════════════════════════════════════════
# Deployment check logic
# ══════════════════════════════════════════════════════════════════════════════

function Get-DeploymentResult {
    $checks   = [ordered]@{}
    $hardFail = $false

    # ── Battery health ────────────────────────────────────────────────────
    $battPct = $null
    try {
        if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
            $lbs = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop)
            $healths = @()
            foreach ($lb in $lbs) {
                $dcStr = ((Get-SafeWmiProperty $lb "DesignCapacity")    -replace '[^0-9,\.]','') -replace ',','.'
                $fcStr = ((Get-SafeWmiProperty $lb "FullChargeCapacity") -replace '[^0-9,\.]','') -replace ',','.'
                [double]$dc = 0; [double]$fc = 0
                if ([double]::TryParse($dcStr,[System.Globalization.NumberStyles]::Any,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$dc) -and
                    [double]::TryParse($fcStr,[System.Globalization.NumberStyles]::Any,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$fc) -and
                    $dc -gt 0) { $healths += [math]::Round(($fc/$dc)*100,1) }
            }
            if ($healths.Count -gt 0) { $battPct = ($healths | Measure-Object -Minimum).Minimum }
        }
    } catch {}
    if ($null -eq $battPct) {
        try {
            $acpi = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
            $healths = @()
            for ($i = 0; $i -lt $acpi.Count; $i++) {
                $fc = $acpi[$i].FullChargedCapacity; $dc = Get-DesignCapacity -Index $i
                [int64]$fcN = 0; [int64]$dcN = 0
                if ($fc -and $dc -and [int64]::TryParse($fc.ToString(),[ref]$fcN) -and
                    [int64]::TryParse($dc.ToString(),[ref]$dcN) -and $dcN -gt 0) {
                    $healths += [math]::Round(($fcN/$dcN)*100,1)
                }
            }
            if ($healths.Count -gt 0) { $battPct = ($healths | Measure-Object -Minimum).Minimum }
        } catch {}
    }
    if ($null -ne $battPct) {
        $s = if ($battPct -ge 80) { "PASS" } elseif ($battPct -ge 70) { "WARN" } else { "FAIL" }
        if ($s -eq "FAIL") { $hardFail = $true }
        $checks["Battery Health"] = [PSCustomObject]@{ Value = "$battPct%"; Status = $s }
    } else {
        $checks["Battery Health"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "WARN" }
    }

    # ── Cycle count ───────────────────────────────────────────────────────
    $cycles = $null
    try {
        if (Test-LenovoNamespace) {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            [int64]$c = 0
            if ([int64]::TryParse(([string]$odo.Battery_cycles -replace '\D',''),[ref]$c)) { $cycles = $c }
        }
    } catch {}
    if ($null -ne $cycles) {
        $s = if ($cycles -lt 300) { "PASS" } elseif ($cycles -lt 500) { "WARN" } else { "FAIL" }
        if ($s -eq "FAIL") { $hardFail = $true }
        $checks["Cycle Count"] = [PSCustomObject]@{ Value = "$cycles cycles"; Status = $s }
    } else {
        $checks["Cycle Count"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "INFO" }
    }

    # ── Memory ────────────────────────────────────────────────────────────
    try {
        $mem = Get-MemoryState
        if ($mem.ModuleCount -gt 0) {
            $s = if ($mem.SpeedMismatch) { "WARN" } else { "PASS" }
            $checks["Memory"] = [PSCustomObject]@{ Value = "$($mem.TotalRAMGB) GB  ($($mem.ModuleCount) module(s))"; Status = $s }
        } else {
            $checks["Memory"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "INFO" }
        }
    } catch {
        $checks["Memory"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "INFO" }
    }

    # ── BIOS ──────────────────────────────────────────────────────────────
    $si = Get-SystemInfo
    $s  = if ($si.BIOSVersion -ne "Unknown") { "PASS" } else { "WARN" }
    $checks["BIOS Version"] = [PSCustomObject]@{ Value = "$($si.BIOSVersion)  ($($si.BIOSDate))"; Status = $s }

    # ── CDRT shock / thermal ──────────────────────────────────────────────
    $cdrt = Get-CdrtData
    if ($cdrt.Available) {
        $shockVal = if ($null -ne $cdrt.ShockEvents) { "$($cdrt.ShockEvents)" } else { "N/A" }
        $shockS   = if ($null -eq $cdrt.ShockEvents) { "INFO" } elseif ($cdrt.ShockEvents -le 60000) { "PASS" } else { "WARN" }
        $checks["Shock Events"] = [PSCustomObject]@{ Value = $shockVal; Status = $shockS }

        $thermVal = if ($null -ne $cdrt.ThermalEvents) { "$($cdrt.ThermalEvents)" } else { "N/A" }
        $thermS   = if ($null -eq $cdrt.ThermalEvents) { "INFO" } elseif ($cdrt.ThermalEvents -le 200) { "PASS" } else { "WARN" }
        $checks["Thermal Events"] = [PSCustomObject]@{ Value = $thermVal; Status = $thermS }
    } else {
        $checks["Shock / Thermal"] = [PSCustomObject]@{ Value = "Lenovo Vantage not installed — skipped"; Status = "INFO" }
    }

    # ── Battery serial / barcode mismatch ────────────────────────────────
    # Compares EC firmware serial (Lenovo_Battery) against ACPI serial
    # (Win32_Battery). A mismatch may indicate a third-party replacement.
    # Surfaced as WARN only — under EU Battery Regulation (effective
    # Feb 18 2027) third-party replacements are a consumer right and must
    # not be treated as a defect or deployment blocker.
    $serialCheck = Get-BatterySerialMismatch
    if ($serialCheck.Available) {
        if ($serialCheck.Mismatch) {
            $checks["Battery Serial"] = [PSCustomObject]@{
                Value  = "Mismatch detected — may be third-party replacement  |  $($serialCheck.Detail)"
                Status = "WARN"
            }
        } else {
            $checks["Battery Serial"] = [PSCustomObject]@{
                Value  = $serialCheck.Detail
                Status = "PASS"
            }
        }
    } else {
        $checks["Battery Serial"] = [PSCustomObject]@{
            Value  = "Unavailable"
            Status = "INFO"
        }
    }

    # ── Verdict ───────────────────────────────────────────────────────────
    $anyWarn    = $checks.Values | Where-Object { $_.Status -eq "WARN" }
    $status     = if ($hardFail) { "NOT READY" } elseif ($anyWarn) { "REVIEW NEEDED" } else { "READY" }
    $statusColor = if ($hardFail) { "Red" } elseif ($anyWarn) { "Yellow" } else { "Green" }

    return [PSCustomObject]@{
        Checks      = $checks
        Status      = $status
        StatusColor = $statusColor
        SystemInfo  = $si
        HardFail    = $hardFail
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# Run and display
# ══════════════════════════════════════════════════════════════════════════════

Write-Host "==============================================="
Write-Host "  THINKPAD DEPLOYMENT CHECK" -ForegroundColor Cyan
Write-Host "==============================================="
Write-Host ""
Write-Host "Checking device..." -ForegroundColor DarkGray
Write-Host ""

$result = Get-DeploymentResult
$si     = $result.SystemInfo

# ── Screen output ─────────────────────────────────────────────────────────────
Write-Host "  Device      : $env:COMPUTERNAME"
Write-Host "  Model       : $($si.Model)"
Write-Host "  BIOS        : $($si.BIOSVersion)  ($($si.BIOSDate))"
Write-Host "  Environment : $($si.Environment)"
Write-Host ""

Write-Host "-----------------------------------------------" -ForegroundColor DarkGray
foreach ($key in $result.Checks.Keys) {
    $chk    = $result.Checks[$key]
    $badge  = switch ($chk.Status) {
        "PASS" { "✔ PASS" }; "WARN" { "⚠ WARN" }; "FAIL" { "✘ FAIL" }; "INFO" { "  INFO" }
    }
    $bColor = switch ($chk.Status) {
        "PASS" { "Green" }; "WARN" { "Yellow" }; "FAIL" { "Red" }; "INFO" { "DarkGray" }
    }
    $label = $key.PadRight(20)
    Write-Host "  $label : $($chk.Value)" -NoNewline
    Write-Host "  [$badge]" -ForegroundColor $bColor
}
Write-Host "-----------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

Write-Host "  Result : " -NoNewline
Write-Host $result.Status -ForegroundColor $result.StatusColor
Write-Host ""

Write-Host "  Criteria:" -ForegroundColor DarkGray
Write-Host "    PASS  Battery >= 80%,  Cycles < 300,  Memory matched,  Serial consistent" -ForegroundColor DarkGray
Write-Host "    WARN  Battery 70-79%,  Cycles 300-499,  Memory mismatch,  Serial mismatch (third-party)" -ForegroundColor DarkGray
Write-Host "    FAIL  Battery < 70%,   Cycles >= 500" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Note: A battery serial mismatch indicates a possible third-party replacement." -ForegroundColor DarkGray
Write-Host "  Under EU Battery Regulation (2027), third-party replacements are a consumer" -ForegroundColor DarkGray
Write-Host "  right and are not treated as a defect. Verify battery condition independently." -ForegroundColor DarkGray
Write-Host ""

# ── Save report to Desktop ────────────────────────────────────────────────────
$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
if (-not $desktopPath -or -not (Test-Path $desktopPath)) {
    $desktopPath = "$env:USERPROFILE\Desktop"
}
$reportPath = Join-Path $desktopPath "ThinkPad_DeploymentCheck.txt"

$lines = @()
$lines += "==============================================="
$lines += " THINKPAD DEPLOYMENT CHECK REPORT"
$lines += " Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$lines += " Device     : $env:COMPUTERNAME"
$lines += " Model      : $($si.Model)"
$lines += " BIOS       : $($si.BIOSVersion)  ($($si.BIOSDate))"
$lines += " Environment: $($si.Environment)"
$lines += "==============================================="
$lines += ""
$lines += "[ DEPLOYMENT READINESS ]"
foreach ($key in $result.Checks.Keys) {
    $chk   = $result.Checks[$key]
    $label = $key.PadRight(20)
    $lines += "  $label : $($chk.Value)  [$($chk.Status)]"
}
$lines += ""
$lines += "  Result : $($result.Status)"
$lines += ""
$lines += "-----------------------------------------------"
$lines += "  PASS  Battery >= 80%,  Cycles < 300,  Memory matched,  Serial consistent"
$lines += "  WARN  Battery 70-79%,  Cycles 300-499,  Memory mismatch,  Serial mismatch (third-party)"
$lines += "  FAIL  Battery < 70%,   Cycles >= 500"
$lines += ""
$lines += "  Note: A battery serial mismatch indicates a possible third-party replacement."
$lines += "  Under EU Battery Regulation (2027), third-party replacements are a consumer"
$lines += "  right and are not treated as a defect. Verify battery condition independently."
$lines += "-----------------------------------------------"

$lines | Out-File $reportPath -Encoding UTF8 -Force

Write-Host "  Report saved to:" -ForegroundColor DarkGray
Write-Host "  $reportPath" -ForegroundColor DarkGray
Write-Host ""

# ── Tear down ─────────────────────────────────────────────────────────────────
if ($script:CimSession) {
    Remove-CimSession -CimSession $script:CimSession -ErrorAction SilentlyContinue
}

Read-Host "Press ENTER to close"