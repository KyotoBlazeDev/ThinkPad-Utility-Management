Clear-Host

# ── Administrator privilege check with auto-elevation ────────────────────
# The root\wmi and root\Lenovo WMI namespaces communicate directly with the
# system BIOS and Embedded Controller (EC). Windows restricts access to these
# namespaces to Administrator-level processes to prevent unprivileged software
# from reading or modifying firmware-level hardware state. Without elevation,
# every WMI call in this script will fail mid-session with opaque access-denied
# errors rather than surfacing a clear message to the user.
#
# When not running as Administrator, attempt to relaunch this script elevated
# via a UAC prompt. If UAC is declined or unavailable, display a clear
# explanation of why elevation is required before exiting.

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    try {
        $scriptPath = $MyInvocation.MyCommand.Path

        # A resolved path is required to relaunch — the script must be saved to disk.
        # Scripts piped into PowerShell or pasted into a console have no path and
        # cannot be relaunched automatically.
        if (-not $scriptPath) {
            throw "Script path is unavailable. Save the script to a .ps1 file and run it directly."
        }

        # Trigger a UAC elevation prompt and relaunch this script in a new
        # elevated PowerShell process. The current non-elevated instance exits
        # immediately so two copies never run side by side.
        Start-Process -FilePath "powershell.exe" `
                      -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" `
                      -Verb RunAs `
                      -ErrorAction Stop
        Exit

    }
    catch {
        # Reached when: UAC was declined, UAC is disabled by policy, or the
        # script path could not be resolved (pipe / unsaved buffer scenario).
        Clear-Host
        Write-Host ""
        Write-Host "  ADMINISTRATOR ACCESS REQUIRED" -ForegroundColor Red
        Write-Host ""

        if ($_.Exception.Message -match "canceled by the user|was canceled") {
            Write-Host "  You clicked No on the permission prompt." -ForegroundColor Yellow
        }
        else {
            Write-Host "  Could not request administrator access automatically:" -ForegroundColor Yellow
            Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Right-click PowerShell and choose" -ForegroundColor Gray
            Write-Host "  'Run as Administrator', then try again." -ForegroundColor Gray
        }

        Write-Host ""
        Write-Host "  Why is this needed?" -ForegroundColor Cyan
        Write-Host "  This tool reads battery and hardware data directly" -ForegroundColor Gray
        Write-Host "  from your device. Windows requires administrator" -ForegroundColor Gray
        Write-Host "  access to read this information — the tool does" -ForegroundColor Gray
        Write-Host "  not make any changes to your device." -ForegroundColor Gray
        Write-Host ""
        Read-Host "  Press ENTER to exit"
        Exit
    }
}

$script:ErrorLog    = @()
$script:SystemInfo  = $null
$script:StorageInfo = $null   # Cached storage info — populated by Get-StorageInfo
$script:CimSession  = $null   # Shared CIM session — initialised after Show-Disclaimer
$script:BatteryAlert = $null  # Cached battery alert state — populated by Get-BatteryAlertState
$script:MenuLoopCount = 0     # Counts menu iterations — used to trigger guided battery analysis
$script:GuidedAnalysisShown = $false  # Ensures auto-prompt fires only once per session


function New-ScriptCimSession {
    <#
    .SYNOPSIS
    Creates the shared CIM session stored in $script:CimSession and returns it.

    .DESCRIPTION
    A CIM session is a persistent, reusable connection to the local (or remote)
    CIM infrastructure. By creating one session at startup and passing it via
    -CimSession to every Get-CimInstance call, the script avoids the per-call
    overhead of implicitly opening and closing a DCOM/WinRM transport each time.

    Benefits:
      - Faster repeated queries within the same script session
      - Cleaner enterprise execution (single authenticated connection)
      - Forward-compatible with remote Lenovo device management (change
        New-CimSession -ComputerName to target a remote machine)

    Graceful degradation: if New-CimSession fails for any reason (policy
    restriction, WMI service issue), $script:CimSession is left as $null and
    every Get-CimInstance call falls back to its implicit local session.
    No function in this script checks for a null session before calling
    Get-CimInstance — passing $null to -CimSession is equivalent to omitting it.
    #>
    try {
        $script:CimSession = New-CimSession -ErrorAction Stop
    }
    catch {
        # Non-fatal: log the failure and continue without a shared session.
        # All Get-CimInstance calls will fall back to implicit local sessions.
        $script:ErrorLog += [PSCustomObject]@{
            TimeStamp        = Get-Date
            Namespace        = 'local'
            Class            = 'CimSession'
            Message          = "New-CimSession failed: $($_.Exception.Message). Falling back to implicit per-call sessions."
            StackTrace       = $_.Exception.StackTrace
            ScriptName       = $_.InvocationInfo.ScriptName
            ScriptLineNumber = $_.InvocationInfo.ScriptLineNumber
            Context          = 'Session initialisation'
            BatteryDetails   = $null
            FullException    = $_
        }
    }
}

function Get-GenuineManufacturerList {
    <#
    .SYNOPSIS
    Returns the list of known OEM battery manufacturer name tokens.

    .DESCRIPTION
    DEPRECATED — no longer used as a pass/fail gate in Get-BatteryClassification.

    Historical context: This function was introduced when Lenovo actively enforced
    a battery whitelist via BIOS firmware (circa 2016), blocking third-party cells
    from charging. The "non-genuine" warning was intended to surface that lockdown
    to users.

    As of EU Battery Regulation (Regulation (EU) 2023/1542, effective Feb 18 2027),
    consumers have an explicit legal right to replace batteries with third-party
    cells. Treating a non-OEM manufacturer name as a WARN or FAIL condition is
    therefore inappropriate in an EU-facing tool.

    This function is retained for reference and potential informational display
    but must not be used to produce warnings, failures, or any output implying
    a third-party battery is unsafe or non-compliant.

    Scheduled for full removal in a future cleanup pass post-v2.0 release.

    .NOTES
    v2.0 — Deprecated. See GitHub Issue #1 (EU Compliance milestone).
    #>

    # Built-in fallback list -- always used if JSON load fails for any reason
    $builtIn = @(
        "LENOVO",    # Lenovo-branded packs
        "PANASONIC", # Historic OEM supplier (IBM/early Lenovo era)
        "SANYO",     # Historic OEM supplier
        "SONY",      # Historic OEM supplier
        "MURATA",    # Sony battery division acquired by Murata in 2017
        "LGC",       # LG Chem - common modern OEM supplier
        "LG",        # LG Chem alternate short form reported by some firmware
        "SDI",       # Samsung SDI
        "SAMSUNG",   # Samsung SDI alternate form
        "CELXPERT",  # Confirmed Lenovo-authorized supplier
        "ATL",       # Amperex Technology Limited
        "COSMX",     # Authorized supplier
        "BYD",       # BYD battery division
        "NVT",       # Authorized supplier
        "SUNWODA",   # Confirmed Lenovo-authorized supplier
        "SMP"        # Authorized supplier
    )

    # Resolve JSON path relative to this script file
    $jsonPath = Join-Path $PSScriptRoot "manufacturers.json"

    if (-not (Test-Path $jsonPath)) {
        return $builtIn
    }

    try {
        $data = Get-Content -Path $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        $names = @($data.genuineManufacturers |
                   Where-Object { $_.name -and $_.name.Trim() -ne "" } |
                   Select-Object -ExpandProperty name)

        if ($names.Count -gt 0) {
            return $names
        }
        else {
            Write-Warning "manufacturers.json contained no valid entries -- using built-in manufacturer list."
            return $builtIn
        }
    }
    catch {
        Write-Warning "Failed to load manufacturers.json ($($_.Exception.Message)) -- using built-in manufacturer list."
        return $builtIn
    }
}

function Get-BatteryClassification {
    param(
        [object]$DesignCapacity,
        [object]$FullChargeCapacity,
        [object]$CycleCount,
        [string]$BatteryStatus = "Unknown",
        [string]$Manufacturer = "Unknown",

        # Optional: pass a pre-computed health percentage to bypass the internal
        # int64 capacity parse. Used by callers (e.g. Lenovo_Battery path) that
        # receive capacity as localised decimal strings ("39,96Wh") and have
        # already calculated the percentage themselves with double arithmetic.
        # When supplied and > 0, DesignCapacity / FullChargeCapacity are ignored
        # for the health calculation — all other logic (genuineness, cycle count,
        # classification thresholds) still runs normally.
        [double]$PrecomputedHealthPercent = 0
    )

    $HealthPercent = 0
    $Classification = "Unknown"
    $Severity = 3
    $IsGenuine = $true
    $FailureFlag = $false
    $Message = ""

    # Calculate health percentage
    # If a pre-computed value was supplied, use it directly and skip the
    # int64 capacity parse which would fail on decimal Wh strings.
    if ($PrecomputedHealthPercent -gt 0) {
        $HealthPercent = [math]::Round($PrecomputedHealthPercent, 2)
    }
    else {
        [int64]$fc = 0
        [int64]$dc = 0

        if ($FullChargeCapacity -and $DesignCapacity -and
        [int64]::TryParse([string]$FullChargeCapacity, [ref]$fc) -and
        [int64]::TryParse([string]$DesignCapacity, [ref]$dc) -and
        $dc -gt 0) {
            $HealthPercent = [math]::Round(($fc / $dc) * 100, 2)
        }
        else {
            $Message = "Unable to calculate health: invalid capacity values"
            $Severity = 3
            $FailureFlag = $true
            return [PSCustomObject]@{
                HealthPercent  = $HealthPercent
                Classification = $Classification
                Severity       = $Severity
                IsGenuine      = $IsGenuine
                FailureFlag    = $FailureFlag
                Message        = $Message
            }
        }
    }

    # DEPRECATED (v2.0 — EU Compliance): Manufacturer allowlist check removed.
    # Get-GenuineManufacturerList was previously used here to flag third-party
    # batteries as non-genuine. Under EU Battery Regulation (effective Feb 18 2027)
    # this is no longer appropriate — third-party replacements are a consumer right.
    # $IsGenuine is preserved as $true to suppress legacy warning display sites
    # that have not yet been removed. See GitHub Issue #1.
    $IsGenuine = $true

    # Parse cycle count
    [int64]$cycles = 0
    if ($CycleCount) {
        $cycleStr = $CycleCount.ToString() -replace '\D', ''
        [int64]::TryParse($cycleStr, [ref]$cycles) | Out-Null
    }

    # Classify battery health and determine severity
    # Check for critical conditions — any status outside the known-good set
    # (or an explicit failure keyword) is treated as critical.
    $isCriticalStatus = if ($BatteryStatus -and $BatteryStatus -ne "Unknown" -and $BatteryStatus -ne "N/A") {
        $normalStates = @("OK", "Discharging", "Charging", "Unknown")
        (-not ($normalStates -contains $BatteryStatus)) -or ($BatteryStatus -match "Failed|Critical|Error|Degraded")
    } else {
        $false
    }

    if ($isCriticalStatus -or $HealthPercent -lt 20 -or $cycles -gt 1500) {
        $Classification = "Critical"
        $Severity = 4
        $FailureFlag = $true
        $Message = "The battery reports a failure and needs to be replaced as soon as possible."
    }
    elseif ($HealthPercent -lt 40 -or ($cycles -gt 1000 -and $HealthPercent -lt 60)) {
        $Classification = "Poor"
        $Severity = 3
        $Message = "Battery health is poor. Consider replacement soon."
    }
    elseif ($HealthPercent -lt 60 -or ($cycles -gt 500 -and $HealthPercent -lt 75)) {
        $Classification = "Fair"
        $Severity = 2
        $Message = "The battery is functioning correctly, but due to normal aging of the battery, the battery life between charges is now significantly shorter than when it was new."
    }
    elseif ($HealthPercent -lt 80) {
        # Warning tier — aligns with Lenovo Vantage's updated alert threshold.
        # Vantage now flags batteries below 80% as needing attention.
        # This tier surfaces the same early signal while keeping the existing
        # Fair / Poor / Critical thresholds intact for backwards compatibility.
        $Classification = "Warning"
        $Severity = 1
        $Message = "Battery health is below 80%. Lenovo Vantage recommends attention at this level. Monitor battery life and consider replacement if runtime becomes insufficient."
    }
    else {
        $Classification = "Good"
        $Severity = 0
        $Message = "Battery health is good."
    }

    # Add genuineness warning to message
    # DEPRECATED (v2.0): Non-genuine warning message suppressed.
    # $IsGenuine is always $true — this block is inert and retained only until
    # the IsGenuine field is fully removed from the return object in a future pass.

    return [PSCustomObject]@{
        HealthPercent = $HealthPercent
        Classification = $Classification
        Severity = $Severity
        IsGenuine = $IsGenuine
        FailureFlag = $FailureFlag
        Message = $Message
    }
}

function Log-Error {
    param($Namespace,$Class,$Exception)

    $ExceptionMessage = $Exception.Exception.Message
    $StackTrace = $Exception.Exception.StackTrace
    $InvocationInfo = $Exception.InvocationInfo

    $script:ErrorLog += [PSCustomObject]@{
        TimeStamp       = Get-Date
        Namespace       = $Namespace
        Class           = $Class
        Message         = $ExceptionMessage
        StackTrace      = $StackTrace
        ScriptName      = $InvocationInfo.ScriptName
        ScriptLineNumber = $InvocationInfo.ScriptLineNumber
        Context         = $null
        BatteryDetails  = $null
        FullException   = $Exception
    }
}

function Log-BatteryError {
    param($Namespace, $Class, $Exception, $BatteryIndex, $BatteryID, $BatteryDetails)

    $ExceptionMessage = $Exception.Exception.Message
    $StackTrace = $Exception.Exception.StackTrace
    $InvocationInfo = $Exception.InvocationInfo

    $contextString = "Battery Index: $BatteryIndex | Battery ID: $BatteryID"
    if ($BatteryDetails) {
        $contextString += " | InstanceName: $($BatteryDetails.InstanceName)"
    }

    $script:ErrorLog += [PSCustomObject]@{
        TimeStamp        = Get-Date
        Namespace        = $Namespace
        Class            = $Class
        Message          = $ExceptionMessage
        StackTrace       = $StackTrace
        ScriptName       = $InvocationInfo.ScriptName
        ScriptLineNumber = $InvocationInfo.ScriptLineNumber
        Context          = $contextString
        BatteryDetails   = $BatteryDetails
        FullException    = $Exception
    }
}


function Get-SystemInfo {
    <#
    .SYNOPSIS
    Queries Win32_ComputerSystem and Win32_BIOS once and caches the result
    in $script:SystemInfo. Subsequent calls return the cached object.
    #>
    if ($script:SystemInfo) { return $script:SystemInfo }

    $model       = "Unknown"
    $biosVersion = "Unknown"
    $biosDate    = "Unknown"

    try {
        $cs = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            # Win32_ComputerSystem.Model is the most reliable field for ThinkPad model strings
            $model = "$($cs.Manufacturer) $($cs.Model)".Trim()
        }
    } catch {}

    try {
        $bios = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) {
            $biosVersion = $bios.SMBIOSBIOSVersion
            # ReleaseDate is a CIM datetime — format to readable date
            if ($bios.ReleaseDate) {
                $biosDate = ([System.Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)).ToString("yyyy-MM-dd")
            }
        }
    } catch {}

    $script:SystemInfo = [PSCustomObject]@{
        Model       = $model
        BIOSVersion = $biosVersion
        BIOSDate    = $biosDate
    }

    return $script:SystemInfo
}


function Get-StorageInfo {
    <#
    .SYNOPSIS
    Returns the primary drive type (SSD/HDD/NVMe), size in GB, and model string.
    Queries MSFT_PhysicalDisk first; falls back to Win32_DiskDrive if unavailable.
    Result is cached in $script:StorageInfo after the first call.
    #>
    if ($script:StorageInfo) { return $script:StorageInfo }

    $result = [PSCustomObject]@{
        Type      = "Unknown"
        SizeGB    = "Unknown"
        Model     = "Unknown"
        Available = $false
    }

    # Primary source: MSFT_PhysicalDisk (Storage module — most reliable for MediaType)
    try {
        $disks = @(Get-PhysicalDisk -ErrorAction SilentlyContinue |
                   Where-Object { $_.MediaType -ne "Unspecified" } |
                   Sort-Object -Property @{Expression={$_.MediaType -eq "HDD"}; Ascending=$true})
        # Prefer SSD/NVMe — sort HDD last so if there's an SSD it appears first
        $disk = $disks | Select-Object -First 1
        if ($disk) {
            $typeMap = @{ "SSD" = "SSD"; "HDD" = "HDD"; "SCM" = "NVMe / SCM" }
            $result.Type      = if ($disk.MediaType -eq "SSD" -and $disk.BusType -eq "NVMe") { "NVMe SSD" }
                                 elseif ($typeMap.ContainsKey($disk.MediaType)) { $typeMap[$disk.MediaType] }
                                 else { $disk.MediaType }
            $result.SizeGB    = if ($disk.Size -gt 0) { [math]::Round($disk.Size / 1GB) } else { "Unknown" }
            $result.Model     = if ($disk.FriendlyName) { $disk.FriendlyName } else { "Unknown" }
            $result.Available = $true
            $script:StorageInfo = $result
            return $result
        }
    } catch {}

    # Fallback: Win32_DiskDrive (always present but lacks MediaType)
    try {
        $drive = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_DiskDrive `
                     -ErrorAction SilentlyContinue |
                 Sort-Object Size -Descending |
                 Select-Object -First 1
        if ($drive) {
            $result.Type      = "Unknown (drive found)"
            $result.SizeGB    = if ($drive.Size -gt 0) { [math]::Round($drive.Size / 1GB) } else { "Unknown" }
            $result.Model     = if ($drive.Model) { $drive.Model } else { "Unknown" }
            $result.Available = $true
        }
    } catch {}

    $script:StorageInfo = $result
    return $result
}

function Get-LenovoDeviceFamily {
    # Consumer edition: root\Lenovo is not available on consumer Lenovo devices.
    # Always returns "Other" so callers degrade gracefully without SIF checks.
    return "Other"
}

function Get-LenovoNamespaceUnavailableMessage {
    <#
    .SYNOPSIS
    Writes a device-aware explanation of why root\Lenovo is unavailable.

    .DESCRIPTION
    If this is a ThinkPad, the namespace should exist — absence means a driver
    problem, so repair advice is shown.

    If this is any other Lenovo device, root\Lenovo is absent by firmware design.
    No driver install or WMI repair will fix this. The user is told clearly that
    the feature is ThinkPad-exclusive and pointed to Lenovo Vantage instead.
    #>
    param(
        [string]$Feature = "this feature"
    )

    $family       = Get-LenovoDeviceFamily
    $sifInstalled = Test-SifInstalled
    $si           = Get-SystemInfo

    switch ($family) {
        "ThinkPad" {
            Write-Host "Extended battery data is not available right now." -ForegroundColor Red
            Write-Host ""
            if ($sifInstalled) {
                Write-Host "The Lenovo driver for $Feature appears to be installed," -ForegroundColor Yellow
                Write-Host "but could not be read. Try restarting your device." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "If the problem persists, reinstall Lenovo drivers from:" -ForegroundColor Yellow
                Write-Host "  support.lenovo.com  →  Drivers & Software"
            } else {
                Write-Host "The Lenovo driver required for $Feature is not installed." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Install it from Lenovo Support:" -ForegroundColor Cyan
                Write-Host "  support.lenovo.com  →  Drivers & Software"
                Write-Host "  Search: 'System Interface Foundation'"
            }
        }

        "Other" {
            Write-Host "$Feature is not available on this device." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Device: $($si.Model)" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "This feature is only supported on Lenovo ThinkPad devices." -ForegroundColor Yellow
            Write-Host "It is not available on other Lenovo models by design." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "For battery and system information on this device," -ForegroundColor Cyan
            Write-Host "try the Lenovo Vantage app (available in the Microsoft Store)." -ForegroundColor Cyan
        }

        default {
            Write-Host "$Feature is not available on this device." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "This feature is only supported on Lenovo ThinkPad devices." -ForegroundColor Yellow
        }
    }
}

function Get-BatteryAlertState {
    <#
    .SYNOPSIS
    Queries all batteries at startup and caches per-battery health results in
    $script:BatteryAlert. Subsequent calls return the cached object.

    .DESCRIPTION
    Most users only open this script after something has already gone wrong —
    the laptop dies unexpectedly, shuts off mid-use, or won't charge past a
    low percentage. By querying health at startup and surfacing a summary in
    the header and menu on every screen, the alert is impossible to miss even
    for users who don't know which menu option to pick.

    Devices with two batteries (e.g. PowerBridge models) have an internal
    and external battery — an internal fixed cell and a hot-swappable external bay battery.
    Both are tracked individually so degradation on either one is reported
    correctly rather than masked by the other.

    Sources tried in order (mirrors the FullCharge function priority):
      1. Lenovo_Battery (root\Lenovo) — richest data, preferred when SIF present
      2. BatteryFullChargedCapacity + BatteryStaticData (root\wmi ACPI fallback)

    Returns a PSCustomObject with:
      Batteries      : array of per-battery results (Index, BatteryID,
                       Severity, Classification, HealthPercent)
      WorstSeverity  : highest Severity across all batteries
      SummaryLine    : pre-formatted summary string for Show-Header
      BatteryCount   : total number of batteries detected
      Available      : $true if health data was found for at least one battery
    #>
    if ($script:BatteryAlert) { return $script:BatteryAlert }

    $result = [PSCustomObject]@{
        Batteries     = @()
        WorstSeverity = 0
        SummaryLine   = ""
        BatteryCount  = 0
        Available     = $false
    }

    # ── Source 1: Lenovo_Battery ─────────────────────────────────────────
    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
        try {
            $lbBatteries = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop)
            if ($lbBatteries -and $lbBatteries.Count -gt 0) {
                $index = 0
                foreach ($lb in $lbBatteries) {
                    $batteryID = Get-SafeWmiProperty -Object $lb -PropertyName "BatteryID"
                    $lbDesign  = Get-SafeWmiProperty -Object $lb -PropertyName "DesignCapacity"
                    $lbFull    = Get-SafeWmiProperty -Object $lb -PropertyName "FullChargeCapacity"
                    $lbCycles  = Get-SafeWmiProperty -Object $lb -PropertyName "CycleCount"
                    $mfr       = Get-SafeWmiProperty -Object $lb -PropertyName "Manufacturer"

                    $dcStr = ($lbDesign -replace '[^0-9,\.]','') -replace ',','.'
                    $fcStr = ($lbFull   -replace '[^0-9,\.]','') -replace ',','.'
                    [double]$dcNum = 0; [double]$fcNum = 0
                    $healthPct = 0
                    if ([double]::TryParse($dcStr, [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dcNum) -and
                        [double]::TryParse($fcStr, [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fcNum) -and
                        $dcNum -gt 0) {
                        $healthPct = [math]::Round(($fcNum / $dcNum) * 100, 2)
                    }

                    $c = Get-BatteryClassification `
                        -PrecomputedHealthPercent $healthPct `
                        -CycleCount               $lbCycles `
                        -Manufacturer             $(if ($mfr -and $mfr -ne "Unavailable") { $mfr } else { "Unknown" })

                    $result.Batteries += [PSCustomObject]@{
                        Index          = $index
                        BatteryID      = $batteryID
                        Severity       = $c.Severity
                        Classification = $c.Classification
                        HealthPercent  = $healthPct
                    }

                    if ($c.Severity -gt $result.WorstSeverity) {
                        $result.WorstSeverity = $c.Severity
                    }

                    $index++
                }

                if ($result.Batteries.Count -gt 0) {
                    $result.BatteryCount = $result.Batteries.Count
                    $result.Available    = $true
                    $script:BatteryAlert = $result
                    $script:BatteryAlert = Set-BatteryAlertSummary $result
                    return $script:BatteryAlert
                }
            }
        } catch {}
    }

    # ── Source 2: ACPI BatteryFullChargedCapacity ────────────────────────
    try {
        $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
        if ($batteries -and $batteries.Count -gt 0) {
            $index = 0
            foreach ($bat in $batteries) {
                $fc  = Get-SafeWmiProperty -Object $bat -PropertyName "FullChargedCapacity"
                $dc  = Get-DesignCapacity -Index $index
                $mfr = Get-BatteryManufacturer -Index $index

                $c = Get-BatteryClassification `
                    -DesignCapacity     $dc `
                    -FullChargeCapacity $fc `
                    -CycleCount         $null `
                    -Manufacturer       $(if ($mfr) { $mfr } else { "Unknown" })

                $result.Batteries += [PSCustomObject]@{
                    Index          = $index
                    BatteryID      = "Battery $index"
                    Severity       = $c.Severity
                    Classification = $c.Classification
                    HealthPercent  = $c.HealthPercent
                }

                if ($c.Severity -gt $result.WorstSeverity) {
                    $result.WorstSeverity = $c.Severity
                }

                $index++
            }

            if ($result.Batteries.Count -gt 0) {
                $result.BatteryCount = $result.Batteries.Count
                $result.Available    = $true
            }
        }
    } catch {}

    $script:BatteryAlert = Set-BatteryAlertSummary $result
    return $script:BatteryAlert
}

function Set-BatteryAlertSummary {
    <#
    .SYNOPSIS
    Builds the summary line for Show-Header from the per-battery results array.

    .DESCRIPTION
    Single battery  : "Battery 0: Good (91%)"
    Dual battery    : "Battery 0: Good (91%)  |  Battery 1: Poor (52%)"
    Any count       : each battery appended with " | " separator

    Only batteries with Severity >= 1 are coloured in the header — the summary
    string itself is plain text; colouring is applied by Show-Header per token.
    #>
    param([object]$AlertState)

    if (-not $AlertState.Available -or $AlertState.Batteries.Count -eq 0) {
        $AlertState.SummaryLine = ""
        return $AlertState
    }

    $parts = @()
    foreach ($b in $AlertState.Batteries) {
        $label = if ($b.BatteryID -and $b.BatteryID -ne "Unavailable") {
            $b.BatteryID
        } else {
            "Battery $($b.Index)"
        }
        $parts += "$label`: $($b.Classification) ($($b.HealthPercent)%)"
    }

    $AlertState.SummaryLine = $parts -join "  |  "
    return $AlertState
}



function Show-Header {
    Clear-Host
    $si  = Get-SystemInfo
    $sto = Get-StorageInfo
    Write-Host "==============================================="
    Write-Host "  LENOVO DEVICE HEALTH CHECK  V.1.1"
    Write-Host "==============================================="
    Write-Host "Device  : $env:COMPUTERNAME"
    Write-Host "Model   : $($si.Model)"
    Write-Host "BIOS    : $($si.BIOSVersion)  ($($si.BIOSDate))"

    # Storage one-liner
    if ($sto.Available) {
        Write-Host "Storage : $($sto.Type)  $($sto.SizeGB) GB  " -NoNewline
        Write-Host "($($sto.Model))" -ForegroundColor DarkGray
    } else {
        Write-Host "Storage : Could not read drive info" -ForegroundColor DarkGray
    }

    # ── Battery health one-liner ─────────────────────────────────────────
    if ($script:BatteryAlert -and $script:BatteryAlert.Available) {
        $worst  = $script:BatteryAlert.WorstSeverity
        $emoji  = @{0="  OK"; 1="⚠"; 2="⚠"; 3="❗"; 4="🔥"}[$worst]
        $phrase = @{0="Healthy"; 1="Below 80%"; 2="Poor"; 3="Critical"; 4="FAILURE"}[$worst]
        $color  = @{0="Green"; 1="Cyan"; 2="Yellow"; 3="Magenta"; 4="Red"}[$worst]
        Write-Host "Battery : " -NoNewline -ForegroundColor White
        Write-Host "$emoji $phrase" -ForegroundColor $color -NoNewline
        if ($worst -ge 1) {
            Write-Host "  → see option 3" -ForegroundColor DarkGray
        } else {
            Write-Host ""
        }
    } else {
        Write-Host "Battery : " -NoNewline -ForegroundColor White
        Write-Host "Could not read battery info" -ForegroundColor DarkGray
    }

    # ── Overall Status line ──────────────────────────────────────────────
    # Derived from battery severity only (primary health indicator).
    # Shown on every screen so the user always has a top-level verdict.
    if ($script:BatteryAlert -and $script:BatteryAlert.Available) {
        $worst = $script:BatteryAlert.WorstSeverity
        $overallLabel = @{0="Good"; 1="Fair"; 2="Fair"; 3="Attention Needed"; 4="Attention Needed"}[$worst]
        $overallColor = @{0="Green"; 1="Cyan"; 2="Yellow"; 3="Red"; 4="Red"}[$worst]
        Write-Host "Overall : " -NoNewline -ForegroundColor White
        Write-Host $overallLabel -ForegroundColor $overallColor
    }

    if ($script:ErrorLog.Count -gt 0) {
        Write-Host "Issues  : " -NoNewline
        Write-Host "$($script:ErrorLog.Count) issue(s) noted this session  [Option 8 to view]" -ForegroundColor Red
    }

    # ── Per-battery detail banner ────────────────────────────────────────
    # Displayed whenever any battery has Severity >= 1 (health below 80%).
    # Shows each affected battery individually on its own line with its
    # classification and percentage, so dual-battery laptops report both.
    if ($script:BatteryAlert -and $script:BatteryAlert.Available -and $script:BatteryAlert.WorstSeverity -ge 1) {
        $bannerColor = switch ($script:BatteryAlert.WorstSeverity) {
            1 { "Cyan" }; 2 { "Yellow" }; 3 { "Magenta" }; 4 { "Red" }
            default { "Yellow" }
        }
        $bannerIcon = switch ($script:BatteryAlert.WorstSeverity) {
            1 { "⚠" }; 2 { "⚠" }; 3 { "❗" }; 4 { "🔥" }
            default { "⚠" }
        }
        $bannerLabel = switch ($script:BatteryAlert.WorstSeverity) {
            1 { "BATTERY ATTENTION RECOMMENDED" }
            2 { "BATTERY HEALTH FAIR — MONITOR CLOSELY" }
            3 { "BATTERY HEALTH POOR — REPLACEMENT ADVISED" }
            4 { "BATTERY CRITICAL — REPLACE IMMEDIATELY" }
            default { "BATTERY ALERT" }
        }
        Write-Host "-----------------------------------------------" -ForegroundColor $bannerColor
        Write-Host "  $bannerIcon  $bannerLabel  $bannerIcon" -ForegroundColor $bannerColor
        foreach ($b in $script:BatteryAlert.Batteries) {
            if ($b.Severity -ge 1) {
                $bLabel = if ($b.BatteryID -and $b.BatteryID -ne "Unavailable") { $b.BatteryID } else { "Battery $($b.Index)" }
                Write-Host "     $bLabel`: $($b.Classification) — $($b.HealthPercent)%  [Select Option 3 for full analysis]" -ForegroundColor $bannerColor
            }
        }
        Write-Host "-----------------------------------------------" -ForegroundColor $bannerColor
    }

    Write-Host ""
}

function Show-Disclaimer {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Yellow
    Write-Host "BEFORE YOU CONTINUE" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This tool reads information directly from your Lenovo device — things like battery health, memory, and BIOS version. It does not change any settings or modify your device in any way."
    Write-Host ""
    Write-Host "Some features may point you to external websites, community pages, or third-party resources. Lenovo does not control or maintain those external sites."
    Write-Host ""
    Write-Host "This tool is not made by or affiliated with Lenovo. It is a free, independent diagnostic tool for personal use."
    Write-Host ""
    Write-Host "Warranty note: Under EU Battery Regulation (effective Feb 2027), you have the right to replace batteries with third-party cells. This tool does not install anything or make any changes."
    Write-Host ""
    Write-Host ""
    Write-Host "You must accept to continue. You may exit without accepting if you do not agree."
    Write-Host ""
    do {
        $response = Read-Host "Type 'A' to Accept and continue, or 'E' to Exit"
        if ($null -ne $response) {
            switch ($response.Trim().ToUpper()) {
                "A" { return }          # proceed with script
                "E" { Clear-Host; Write-Host "Exiting..."; Exit } # exit immediately
                default { Write-Host "Please enter 'A' to accept or 'E' to exit." -ForegroundColor Yellow }
            }
        }
        else {
            Write-Host "Please enter 'A' to accept or 'E' to exit." -ForegroundColor Yellow
        }
    } while ($true)
}

function Get-DesignCapacity {
    <#
    .SYNOPSIS
    Retrieves battery design capacity using multiple fallback sources.
    Returns the design capacity as a string, or $null if not found.

    Sources tried in order:
      1. BatteryStaticData via Get-WmiObject (MUST use WmiObject - CimInstance fails on this class)
      2. powercfg /batteryreport /XML (most reliable universal fallback)
    #>
    param([int]$Index = 0)

    # Source 1: BatteryStaticData - MUST use Get-WmiObject, Get-CimInstance throws generic failure on this class
    try {
        $staticItems = @(Get-WmiObject -Namespace root\wmi -Class BatteryStaticData -ErrorAction SilentlyContinue)
        if ($staticItems -and $Index -lt $staticItems.Count) {
            $val = $staticItems[$Index].DesignedCapacity
            if ($null -ne $val -and $val -gt 0) {
                return [string]$val
            }
        }
    } catch {}

    # Source 2: powercfg /batteryreport /XML
    try {
        $xmlPath = Join-Path $env:TEMP "battery_tmp_$PID.xml"
        $null = & powercfg /batteryreport /XML /OUTPUT $xmlPath 2>$null
        if (Test-Path $xmlPath) {
            [xml]$report = Get-Content $xmlPath -ErrorAction SilentlyContinue
            Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
            $batteries = $report.BatteryReport.Batteries.Battery
            # batteries may be single object or array
            if ($batteries) {
                $batArray = @($batteries)
                if ($Index -lt $batArray.Count) {
                    $val = $batArray[$Index].DesignCapacity
                    if ($val -and [int64]$val -gt 0) {
                        return [string]$val
                    }
                }
            }
        }
    } catch {}

    return $null
}

function Get-BatteryManufacturer {
    <#
    .SYNOPSIS
    Retrieves battery manufacturer using multiple fallback sources.
    Returns the manufacturer string, or $null if not found.

    Sources tried in order:
      1. BatteryStaticData.ManufactureName via Get-WmiObject (MUST use WmiObject - CimInstance fails on this class)
      2. Win32_Battery.Manufacturer via Get-CimInstance (widely available fallback)
    #>
    param([int]$Index = 0)

    # Source 1: BatteryStaticData - MUST use Get-WmiObject
    try {
        $staticItems = @(Get-WmiObject -Namespace root\wmi -Class BatteryStaticData -ErrorAction SilentlyContinue)
        if ($staticItems -and $Index -lt $staticItems.Count) {
            $val = $staticItems[$Index].ManufactureName
            if ($val -and $val.Trim() -ne "") {
                return $val.Trim()
            }
        }
    } catch {}

    # Source 2: Win32_Battery.Manufacturer
    try {
        $win32Items = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery -ErrorAction SilentlyContinue)
        if ($win32Items -and $Index -lt $win32Items.Count) {
            $val = $win32Items[$Index].Manufacturer
            if ($val -and $val.Trim() -ne "") {
                return $val.Trim()
            }
        }
    } catch {}

    return $null
}

function BatteryStaticData {
    Show-Header
    Write-Host "[ Battery Details ]"
    Write-Host ""

    try {
        if (-not (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryStaticData")) {
            Write-Host "Battery details are not available on this device." -ForegroundColor Yellow
            Read-Host "Press ENTER"
            return
        }

        $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryStaticData -ErrorAction SilentlyContinue)

        if (-not $batteries -or $batteries.Count -eq 0) {
            Write-Host "No batteries found." -ForegroundColor Yellow
        }
        else {
            $index = 0
            foreach ($bat in $batteries) {
                try {
                    $deviceID = Get-SafeWmiProperty -Object $bat -PropertyName "DeviceName"
                    $manufacturer = Get-SafeWmiProperty -Object $bat -PropertyName "ManufactureName"
                    $serial = Get-SafeWmiProperty -Object $bat -PropertyName "SerialNumber"
                    $designCapacity = Get-SafeWmiProperty -Object $bat -PropertyName "DesignedCapacity"

                    Write-Host "Battery: $deviceID"
                    Write-Host ("-" * 50)
                    Write-Host "Manufacturer         : $manufacturer"
                    Write-Host "Serial Number        : $serial"
                    Write-Host "Original Capacity    : $designCapacity mWh"
                    
                    Write-Host ""
                    $index++
                }
                catch {
                    Log-BatteryError "root\wmi" "BatteryStaticData" $_ $index $($bat.DeviceName) $null
                    Write-Host "Could not read this battery." -ForegroundColor Red
                    Write-Host ""
                    $index++
                }
            }
        }
    }
    catch {
        Log-Error "root\wmi" "BatteryStaticData" $_
        Write-Host "Could not read battery information." -ForegroundColor Red
    }

    # ── Battery Replacement Safety Warning ───────────────────────────────────
    Write-Host "[ Replacement Safety Warning ]" -ForegroundColor Yellow
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    Write-Host "EN  " -ForegroundColor Cyan -NoNewline
    Write-Host "Replace with same type only. Use of another battery"
    Write-Host "    may present a risk of fire or explosion."
    Write-Host "    PLEASE REFER TO USER MANUAL OR FOLLOW LOCAL ORDINANCES AND/OR REGULATIONS FOR DISPOSAL"
    Write-Host ""
    Write-Host "FR  " -ForegroundColor Cyan -NoNewline
    Write-Host "Remplacer par le même type uniquement. L'utilisation"
    Write-Host "    d'un autre type peut provoquer un incendie ou une explosion."
    Write-Host "    Mettre au rebut les batteries usagées selon les ordonnances et réglementations locales."
    Write-Host ""
    Write-Host "ID  " -ForegroundColor Cyan -NoNewline
    Write-Host "Ganti hanya dengan jenis yang sama. Penggunaan baterai"
    Write-Host "    lain dapat menimbulkan risiko kebakaran atau ledakan."
    Write-Host "    Buang baterai bekas sesuai peraturan dan ketentuan setempat."
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    Write-Host ""

    Read-Host "Press ENTER"
}

function FullCharge {
    Show-Header
    Write-Host "[ Battery Information ]"
    Write-Host ""

    # ── Source priority ──────────────────────────────────────────────────
    # Lenovo_Battery (root\Lenovo) is tried first. It is available on systems
    # with SIF installed and exposes the full set of Lenovo battery properties:
    # identity, health classification, charge state, electrical readings, and
    # firmware metadata in one class.
    #
    # If Lenovo_Battery is unavailable (SIF not installed, older firmware, or
    # non-Lenovo battery controller), the function falls back to the original
    # ACPI-based BatteryFullChargedCapacity path which works on any Windows system.

    $usedLenovoBattery = $false

    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
        try {
            $lbBatteries = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop)

            if ($lbBatteries -and $lbBatteries.Count -gt 0) {
                $usedLenovoBattery = $true
                Write-Host ""

                $index = 0
                foreach ($lb in $lbBatteries) {
                    try {
                        Write-Host "Battery: $($lb.BatteryID)"
                        Write-Host ("=" * 50)

                        # ── Identity ─────────────────────────────────────────
                        Write-Host "[ Identity ]"
                        Write-Host "  Manufacturer         : $(Get-SafeWmiProperty -Object $lb -PropertyName 'Manufacturer')"
                        Write-Host "  Part Number          : $(Get-SafeWmiProperty -Object $lb -PropertyName 'FRUPartNumber')"
                        Write-Host "  Barcode              : $(Get-SafeWmiProperty -Object $lb -PropertyName 'BarCode')"
                        Write-Host "  Battery Type         : $(Get-SafeWmiProperty -Object $lb -PropertyName 'DeviceChemistry')"
                        Write-Host "  Firmware Version     : $(Get-SafeWmiProperty -Object $lb -PropertyName 'FirmwareVersion')"
                        Write-Host "  Manufacture Date     : $(Get-SafeWmiProperty -Object $lb -PropertyName 'ManufactureDate')"
                        Write-Host "  First Used           : $(Get-SafeWmiProperty -Object $lb -PropertyName 'FirstUseDate')"
                        Write-Host ""

                        # ── Health ───────────────────────────────────────────
                        Write-Host "[ Health ]"
                        $lbHealth     = Get-SafeWmiProperty -Object $lb -PropertyName "BatteryHealth"
                        $lbCondition  = Get-SafeWmiProperty -Object $lb -PropertyName "Condition"
                        $lbCycles     = Get-SafeWmiProperty -Object $lb -PropertyName "CycleCount"
                        $lbDesign     = Get-SafeWmiProperty -Object $lb -PropertyName "DesignCapacity"
                        $lbFull       = Get-SafeWmiProperty -Object $lb -PropertyName "FullChargeCapacity"

                        $healthColor = switch -Wildcard ($lbHealth.ToString().ToLower()) {
                            "green"  { "Green"   }
                            "yellow" { "Yellow"  }
                            "red"    { "Red"     }
                            default  { "White"   }
                        }

                        Write-Host "  Battery Health       : " -NoNewline
                        Write-Host $lbHealth -ForegroundColor $healthColor
                        Write-Host "  Condition            : $lbCondition"
                        Write-Host "  Charge Cycles        : $lbCycles"
                        Write-Host "  Original Capacity    : $lbDesign"
                        Write-Host "  Max Charge Now       : $lbFull"
                        Write-Host "  Design Voltage       : $(Get-SafeWmiProperty -Object $lb -PropertyName 'DesignVoltage')"

                        # Calculate numeric health % if both capacity values are parseable
                        # Lenovo_Battery reports capacity as strings like "39,96Wh" — strip non-numeric
                        $healthPct = 0
                        $dcStr = ($lbDesign  -replace '[^0-9,\.]','') -replace ',','.'
                        $fcStr = ($lbFull    -replace '[^0-9,\.]','') -replace ',','.'
                        [double]$dcNum = 0
                        [double]$fcNum = 0
                        if ([double]::TryParse($dcStr, [System.Globalization.NumberStyles]::Any,
                                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dcNum) -and
                            [double]::TryParse($fcStr, [System.Globalization.NumberStyles]::Any,
                                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fcNum) -and
                            $dcNum -gt 0) {
                            $healthPct = [math]::Round(($fcNum / $dcNum) * 100, 2)
                            Write-Host "  Health %             : " -NoNewline
                            Write-Host "$healthPct%" -ForegroundColor $(if ($healthPct -ge 80) { "Green" } elseif ($healthPct -ge 60) { "Cyan" } elseif ($healthPct -ge 40) { "Yellow" } elseif ($healthPct -ge 20) { "Magenta" } else { "Red" })
                        }
                        Write-Host ""

                        # ── Charge State ─────────────────────────────────────
                        Write-Host "[ Charge State ]"
                        $lbStatus  = Get-SafeWmiProperty -Object $lb -PropertyName "Status"
                        $statusColor = switch -Wildcard ($lbStatus.ToString().ToLower()) {
                            "charging"     { "Green"  }
                            "discharging"  { "Yellow" }
                            "idle"         { "Cyan"   }
                            default        { "White"  }
                        }
                        Write-Host "  Status               : " -NoNewline
                        Write-Host $lbStatus -ForegroundColor $statusColor
                        Write-Host "  Remaining Capacity   : $(Get-SafeWmiProperty -Object $lb -PropertyName 'RemainingCapacity')"
                        Write-Host "  Charge Level         : $(Get-SafeWmiProperty -Object $lb -PropertyName 'RemainingPercentage')"
                        Write-Host "  Time Remaining       : $(Get-SafeWmiProperty -Object $lb -PropertyName 'RemainingTime')"
                        Write-Host "  Full Charge By       : $(Get-SafeWmiProperty -Object $lb -PropertyName 'ChargeCompletionTime')"
                        Write-Host "  Charger              : $(Get-SafeWmiProperty -Object $lb -PropertyName 'Adapter')"
                        Write-Host ""

                        # ── Electrical ───────────────────────────────────────
                        Write-Host "[ Electrical ]"
                        Write-Host "  Voltage              : $(Get-SafeWmiProperty -Object $lb -PropertyName 'Voltage')"
                        Write-Host "  Wattage              : $(Get-SafeWmiProperty -Object $lb -PropertyName 'Wattage')"

                        # Temperature - the raw value from Lenovo_Battery varies by firmware
                        # version and locale. Observed formats include:
                        #   "28"    plain integer (most common)
                        #   "28.5"  decimal with period
                        #   "28,5"  decimal with locale comma
                        #   "28 C"  integer with unit suffix
                        #   "285"   tenths-of-a-degree encoding (divide by 10)
                        # Lenovo_Battery carries no timestamp - the value reflects
                        # whatever the EC last wrote to the WMI data block.
                        #
                        # Colour tiers:
                        #   <=40 C  Normal    White
                        #   41-50 C Warning   Yellow
                        #   51-60 C Elevated  Red
                        #   >60 C   Critical  Red
                        $tempRaw = Get-SafeWmiProperty -Object $lb -PropertyName "Temperature"
                        Write-Host "  Temperature          : " -NoNewline
                        $tempNum = $null
                        if ($tempRaw -and $tempRaw -ne "Unavailable") {
                            $tempStripped = (([string]$tempRaw) -replace '[^0-9,\.]','') -replace ',','.'
                            [double]$parsed = 0
                            if ($tempStripped -ne "" -and [double]::TryParse($tempStripped,
                                    [System.Globalization.NumberStyles]::Any,
                                    [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
                                if ($parsed -ge 200 -and ([string]$tempRaw) -notmatch '[Cc]') {
                                    $parsed = [math]::Round($parsed / 10, 1)
                                }
                                $tempNum = $parsed
                            }
                        }
                        if ($null -ne $tempNum) {
                            $tempColor = if     ($tempNum -gt 60) { "Red"    }
                                         elseif ($tempNum -ge 51) { "Red"    }
                                         elseif ($tempNum -ge 41) { "Yellow" }
                                         else                     { "White"  }
                            $tempLabel = if     ($tempNum -gt 60) { "  [!] CRITICAL — Exceeds safe limits" }
                                         elseif ($tempNum -ge 51) { "  [!] HIGH — Getting very warm (51–60 °C)" }
                                         elseif ($tempNum -ge 41) { "  [!] WARM — Higher than normal (41–50 °C)" }
                                         else                     { "  Normal" }
                            Write-Host "$tempNum C" -ForegroundColor $tempColor -NoNewline
                            Write-Host $tempLabel -ForegroundColor $tempColor
                        }
                        else {
                            Write-Host $tempRaw
                        }
                        Write-Host ""

                        # ── Classification ───────────────────────────────────
                        # Feed Lenovo_Battery data into the shared classification
                        # engine so severity, genuineness, and alert flags are
                        # consistent with the rest of the script.
                        $mfr = Get-SafeWmiProperty -Object $lb -PropertyName "Manufacturer"
                        $classification = Get-BatteryClassification `
                            -PrecomputedHealthPercent $healthPct `
                            -CycleCount               $lbCycles `
                            -Manufacturer             $(if ($mfr -and $mfr -ne "Unavailable") { $mfr } else { "Unknown" })

                        $severityColor = switch ($classification.Severity) {
                            0 { "Green" }; 1 { "Cyan" }; 2 { "Yellow" }; 3 { "Magenta" }; 4 { "Red" }
                            default { "White" }
                        }

                        Write-Host "[ Health Assessment ]"
                        Write-Host "  Rating               : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "  Summary              : " -NoNewline
                        Write-Host $classification.Message -ForegroundColor $severityColor

                        # DEPRECATED (v2.0): Non-genuine warning removed — EU Battery Regulation compliance.
                        if ($classification.FailureFlag) {
                            Write-Host "  ALERT: Battery replacement needed immediately!" -ForegroundColor Red
                        }

                        Write-Host ""
                        $index++
                    }
                    catch {
                        Log-BatteryError "root\Lenovo" "Lenovo_Battery" $_ $index $($lb.BatteryID) $null
                        Write-Host "Battery: $index - Could not read this battery" -ForegroundColor Red
                        Write-Host ""
                        $index++
                    }
                }
            }
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Battery" $_
            Write-Host "Battery information could not be loaded — trying another source." -ForegroundColor Yellow
            Write-Host ""
        }
    }

    # ── ACPI fallback ────────────────────────────────────────────────────
    # Used when Lenovo_Battery is unavailable: SIF not installed, older
    # firmware, or the Lenovo_Battery query above failed.
    if (-not $usedLenovoBattery) {
        try {
            if (-not (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryFullChargedCapacity")) {
                Write-Host "Battery information is not available on this device." -ForegroundColor Yellow
                Write-Host "Try installing the latest Lenovo drivers from support.lenovo.com." -ForegroundColor Yellow
                Read-Host "Press ENTER"
                return
            }

            $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

            if (-not $batteries -or $batteries.Count -eq 0) {
                Write-Host "No battery found." -ForegroundColor Yellow
            }
            else {
                Write-Host ""

                $index = 0
                foreach ($bat in $batteries) {
                    try {
                        $FullCharge     = Get-SafeWmiProperty -Object $bat -PropertyName "FullChargedCapacity"
                        $DesignCapValue = Get-DesignCapacity -Index $index
                        $Manufacturer   = Get-BatteryManufacturer -Index $index

                        $DesignCap = if ($DesignCapValue) { $DesignCapValue } else { "Unavailable" }
                        $Health    = "Unavailable"

                        [int64]$fc = 0
                        [int64]$dc = 0
                        if ($null -ne $DesignCapValue -and
                            [int64]::TryParse($FullCharge.ToString(), [ref]$fc) -and
                            [int64]::TryParse($DesignCapValue.ToString(), [ref]$dc) -and
                            $dc -gt 0) {
                            $Health = [math]::Round(($fc / $dc) * 100, 2)
                        }

                        Write-Host "Battery $($index + 1)"
                        Write-Host ("-" * 50)
                        Write-Host "Manufacturer         : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })"
                        Write-Host "Max Charge Now       : $FullCharge mWh"
                        Write-Host "Original Capacity    : $DesignCap mWh"
                        Write-Host "Health               : $Health%"

                        $classification = Get-BatteryClassification `
                            -DesignCapacity     $DesignCapValue `
                            -FullChargeCapacity $FullCharge `
                            -CycleCount         $null `
                            -Manufacturer       $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })

                        $severityColor = switch ($classification.Severity) {
                            0 { "Green" }; 1 { "Cyan" }; 2 { "Yellow" }; 3 { "Magenta" }; 4 { "Red" }
                            default { "White" }
                        }

                        Write-Host "Rating               : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "Summary              : " -NoNewline
                        Write-Host $classification.Message -ForegroundColor $severityColor

                        # DEPRECATED (v2.0): Non-genuine warning removed — EU Battery Regulation compliance.
                        if ($classification.FailureFlag) {
                            Write-Host "ALERT: Battery replacement needed immediately!" -ForegroundColor Red
                        }

                        Write-Host ""
                        $index++
                    }
                    catch {
                        Log-BatteryError "root\wmi" "BatteryFullChargedCapacity" $_ $index $($bat.InstanceName) $null
                        Write-Host "Battery $($index + 1) - Could not read this battery" -ForegroundColor Red
                        Write-Host ""
                        $index++
                    }
                }
            }
        }
        catch {
            Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
            Write-Host "Battery information is not available right now." -ForegroundColor Yellow
        }
    }

    Read-Host "Press ENTER"
}

function DiagnosticInfo {
    Show-Header
    Write-Host "[ What This Tool Can See On Your Device ]"
    Write-Host ""
    Write-Host "Checking what information is available..." -ForegroundColor Cyan
    Write-Host ""

    $sifInstalled  = Test-SifInstalled
    $deviceFamily  = Get-LenovoDeviceFamily
    $hasLenovoNS   = Test-LenovoNamespace
    $cvInstalled   = Test-CommercialVantage

    # ── Battery ──────────────────────────────────────────────────────────
    Write-Host "[ Battery ]" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "  Battery details (manufacturer, serial)  : " -NoNewline
    if (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryStaticData") {
        Write-Host "Available" -ForegroundColor Green
    } else {
        Write-Host "Not available on this device" -ForegroundColor Yellow
    }

    Write-Host "  Battery health and capacity             : " -NoNewline
    $dcVal = Get-DesignCapacity -Index 0
    if ($dcVal) {
        Write-Host "Available  ($dcVal mWh original capacity)" -ForegroundColor Green
    } else {
        Write-Host "Not available" -ForegroundColor Yellow
    }

    Write-Host "  Charge cycle count                      : " -NoNewline
    if ($hasLenovoNS -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
        Write-Host "Available" -ForegroundColor Green
    } else {
        Write-Host "Not available on this device" -ForegroundColor Yellow
        if ($deviceFamily -ne "Other" -and -not $sifInstalled) {
            Write-Host "    → Install Lenovo drivers from support.lenovo.com to enable this." -ForegroundColor DarkGray
        }
    }

    Write-Host ""

    # ── Extended usage history (Lenovo Vantage) ───────────────────────────
    Write-Host "[ Extended Usage Tracking ]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  CPU uptime, drop events, heat events    : " -NoNewline
    if ($cvInstalled -and $hasLenovoNS) {
        $cdrtData = Get-CdrtOdometerData
        if ($cdrtData.Available) {
            Write-Host "Available" -ForegroundColor Green
            Write-Host "    CPU time   : " -NoNewline
            Write-Host $(if ($null -ne $cdrtData.CpuUptimeMinutes) { "~$([math]::Round($cdrtData.CpuUptimeMinutes / 60, 1)) hours" } else { "Not found" }) -ForegroundColor $(if ($null -ne $cdrtData.CpuUptimeMinutes) { "Green" } else { "Yellow" })
            Write-Host "    Drop events: " -NoNewline
            Write-Host $(if ($null -ne $cdrtData.ShockEvents) { $cdrtData.ShockEvents } else { "Not found" }) -ForegroundColor $(if ($null -ne $cdrtData.ShockEvents) { "Green" } else { "Yellow" })
            Write-Host "    Heat events: " -NoNewline
            Write-Host $(if ($null -ne $cdrtData.ThermalEvents) { $cdrtData.ThermalEvents } else { "Not found" }) -ForegroundColor $(if ($null -ne $cdrtData.ThermalEvents) { "Green" } else { "Yellow" })
        } else {
            Write-Host "Not available" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Not available on this device" -ForegroundColor Yellow
        Write-Host "    (Requires Lenovo Vantage to be installed)" -ForegroundColor DarkGray
    }

    Write-Host ""

    # ── Memory ────────────────────────────────────────────────────────────
    Write-Host "[ Memory (RAM) ]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Memory module info                      : " -NoNewline
    if (Test-WmiClass -ClassName "Win32_PhysicalMemory") {
        Write-Host "Available" -ForegroundColor Green
    } else {
        Write-Host "Not available on this device" -ForegroundColor Yellow
    }

    Write-Host ""

    # ── BIOS ──────────────────────────────────────────────────────────────
    Write-Host "[ BIOS / Firmware ]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Installed BIOS version                  : " -NoNewline
    $si = Get-SystemInfo
    if ($si.BIOSVersion -and $si.BIOSVersion -ne "Unknown") {
        Write-Host "$($si.BIOSVersion)  ($($si.BIOSDate))" -ForegroundColor Green
    } else {
        Write-Host "Not available" -ForegroundColor Yellow
    }
    Write-Host "  Online BIOS update check                : Available (requires internet)" -ForegroundColor Green

    Write-Host ""
    Read-Host "Press ENTER"
}

function ComprehensiveBatteryAnalysis {
    Show-Header
    Write-Host "[ Comprehensive Battery Analysis ]"
    Write-Host ""

    $hasLenovoData = $false
    $hasACPIData = $false
    $cycles = 0

    # Try to get Lenovo cycle data
    # Requires Lenovo System Interface Foundation (SIF) to be installed.
    # SIF registers the root\Lenovo namespace; without it this block is skipped
    # and the analysis proceeds with ACPI-only capacity data.
    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
        try {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            $cycles = ([string]$odo.Battery_cycles -replace '\D','')
            $hasLenovoData = $true
            
            Write-Host "[ Usage History ]"
            Write-Host "Charge Cycles : $cycles"
            Write-Host "Drop Events   : $($odo.Shock_events)"
            Write-Host "Heat Events   : $($odo.Thermal_events)"
            Write-Host ""
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Odometer" $_
            Write-Host "Charge cycle data is not available right now." -ForegroundColor Yellow
            Write-Host ""
        }
    }
    else {
        Write-Host "[ Usage History ]" -ForegroundColor Yellow
        Get-LenovoNamespaceUnavailableMessage -Feature "battery cycle count"
        Write-Host ""
    }

    # Try to get ACPI battery capacity data
    if (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryFullChargedCapacity") {
        try {
            $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

            if ($batteries -and $batteries.Count -gt 0) {
                $hasACPIData = $true
                Write-Host "[ Battery Health ]"
                Write-Host ""

                $index = 0
                foreach ($bat in $batteries) {
                    try {
                        $FullCharge = Get-SafeWmiProperty -Object $bat -PropertyName "FullChargedCapacity"
                        
                        $DesignCapValue = Get-DesignCapacity -Index $index
                        $Manufacturer   = Get-BatteryManufacturer -Index $index
                        
                        $DesignCap = if ($DesignCapValue) { $DesignCapValue } else { "Unavailable" }
                        
                        Write-Host "Battery $($index + 1)"
                        Write-Host ("-" * 50)
                        Write-Host "Manufacturer         : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })"
                        Write-Host "Max Charge Now       : $FullCharge mWh"
                        Write-Host "Original Capacity    : $DesignCap mWh"
                        
                        $classification = Get-BatteryClassification -DesignCapacity $DesignCapValue -FullChargeCapacity $FullCharge -CycleCount $cycles -Manufacturer $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })
                        
                        Write-Host ""
                        Write-Host "[ Health Assessment ]"
                        
                        $severityColor = switch ($classification.Severity) {
                             0 { "Green" }
                            1 { "Cyan" }
                            2 { "Yellow" }
                            3 { "Magenta" }
                            4 { "Red" }
                            default { "White" }
                        }
                        
                        Write-Host "Health               : " -NoNewline
                        Write-Host "$($classification.HealthPercent)%" -ForegroundColor $severityColor
                        
                        Write-Host "Rating               : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "Summary              : " -NoNewline
                        Write-Host $classification.Message -ForegroundColor $severityColor
                        
                        # DEPRECATED (v2.0): Non-genuine warning removed — EU Battery Regulation compliance.
                        
                        if ($classification.FailureFlag) {
                            Write-Host "ALERT: Battery replacement needed immediately!" -ForegroundColor Red
                        }
                        
                        Write-Host ""
                        $index++
                    }
                    catch {
                        Log-BatteryError "root\wmi" "BatteryFullChargedCapacity" $_ $index $($bat.InstanceName) $null
                        Write-Host "Battery $($index + 1) - Could not complete analysis" -ForegroundColor Red
                        Write-Host ""
                        $index++
                    }
                }
            }
            else {
                Write-Host "No battery data found." -ForegroundColor Yellow
                Write-Host ""
            }
        }
        catch {
            Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
            Write-Host "Battery information is not available right now." -ForegroundColor Yellow
            Write-Host ""
        }
    }

    # Try Win32_Battery as fallback
    if (-not $hasACPIData) {
        try {
            $batteries = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery -ErrorAction SilentlyContinue)

            if ($batteries -and $batteries.Count -gt 0) {
                Write-Host "[ Battery ]"
                Write-Host ""

                foreach ($bat in $batteries) {
                    try {
                        $percent = Get-SafeWmiProperty -Object $bat -PropertyName "EstimatedChargeRemaining"
                        $deviceID = Get-SafeWmiProperty -Object $bat -PropertyName "DeviceID"
                        $manufacturer = Get-SafeWmiProperty -Object $bat -PropertyName "Manufacturer"
                        $status = Get-SafeWmiProperty -Object $bat -PropertyName "Status"

                        Write-Host "Battery: $deviceID"
                        Write-Host "Current Charge   : $percent%"
                        
                        $classification = Get-BatteryClassification -DesignCapacity $null -FullChargeCapacity $percent -CycleCount $cycles -BatteryStatus $status -Manufacturer $manufacturer
                        
                        $severityColor = switch ($classification.Severity) {
                             0 { "Green" }
                            1 { "Cyan" }
                            2 { "Yellow" }
                            3 { "Magenta" }
                            4 { "Red" }
                            default { "White" }
                        }
                        
                        Write-Host "Rating           : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "Status           : $status"
                        
                        # DEPRECATED (v2.0): Non-genuine warning removed — EU Battery Regulation compliance.
                        
                        Write-Host ""
                    }
                    catch {
                        Log-BatteryError "root\cimv2" "Win32_Battery" $_ $null $($bat.DeviceID) $null
                        Write-Host "Could not read battery information." -ForegroundColor Red
                        Write-Host ""
                    }
                }
            }
        }
        catch {
            Log-Error "root\cimv2" "Win32_Battery" $_
            Write-Host "No battery information found on this device." -ForegroundColor Red
        }
    }
    
    if (-not $hasLenovoData -and -not $hasACPIData) {
        Write-Host "Battery information is not available on this device." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Try installing the latest Lenovo drivers from support.lenovo.com." -ForegroundColor Yellow
    }

    Read-Host "Press ENTER"
}

function ExportReport {

    $reportPath = "$env:USERPROFILE\Lenovo_Report.txt"
    $errorPath  = "$env:USERPROFILE\Lenovo_ErrorLog.txt"

    [System.Collections.Generic.List[string]]$Report = @()

    $si = Get-SystemInfo
    $Report += "==============================================="
    $Report += " LENOVO DEVICE HEALTH REPORT"
    $Report += " Generated : $(Get-Date)"
    $Report += " Device    : $env:COMPUTERNAME"
    $Report += " Model     : $($si.Model)"
    $Report += " BIOS      : $($si.BIOSVersion)  ($($si.BIOSDate))"
    $Report += "==============================================="
    $Report += ""

    # ---------------- Storage ----------------
    $sto = Get-StorageInfo
    $Report += "[ Storage ]"
    if ($sto.Available) {
        $Report += "Type       : $($sto.Type)"
        $Report += "Size       : $($sto.SizeGB) GB"
        $Report += "Model      : $($sto.Model)"
    } else {
        $Report += "Drive information not available."
    }
    $Report += ""

    # ---------------- Windows Battery ----------------
    try {
        if (Test-WmiClass -ClassName "Win32_Battery") {
            $batteries = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery -ErrorAction SilentlyContinue)

            $Report += "[ Battery ]"
            if (-not $batteries -or $batteries.Count -eq 0) {
                $Report += "No battery found."
            }
            else {
                foreach ($bat in $batteries) {
                    $Report += "Battery ID       : $(Get-SafeWmiProperty -Object $bat -PropertyName 'DeviceID')"
                    $Report += "Manufacturer     : $(Get-SafeWmiProperty -Object $bat -PropertyName 'Manufacturer')"
                    $Report += "Serial           : $(Get-SafeWmiProperty -Object $bat -PropertyName 'SerialNumber')"
                    $Report += "Current Charge   : $(Get-SafeWmiProperty -Object $bat -PropertyName 'EstimatedChargeRemaining')%"
                    $Report += ""
                }
            }
            $Report += ""
        }
        else {
            $Report += "[ Battery ]"
            $Report += "Battery information not available on this device."
            $Report += ""
        }
    }
    catch {
        Log-Error "root\cimv2" "Win32_Battery" $_
        $Report += "[ Battery ]"
        $Report += "Could not read battery information."
        $Report += ""
    }

    # ---------------- Lenovo Cycles ----------------
    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
        try {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            $cycles = ([string]$odo.Battery_cycles -replace '\D','')

            $Report += "[ Usage History ]"
            $Report += "Charge Cycles : $cycles"
            $Report += "Drop Events   : $($odo.Shock_events)"
            $Report += "Heat Events   : $($odo.Thermal_events)"
            $Report += ""
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Odometer" $_
            $Report += "Usage history unavailable."
            $Report += ""
        }
    }
    else {
        $Report += "[ Usage History ]"
        $Report += "Charge cycle data is not available on this device."
        $Report += ""
    }

    # ---------------- Warranty ----------------
    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_WarrantyInformation")) {
        try {
            $w = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_WarrantyInformation -ErrorAction Stop

            $Report += "[ Warranty ]"
            $Report += "Serial    : $($w.SerialNumber)"
            $Report += "Start     : $($w.StartDate)"
            $Report += "End       : $($w.EndDate)"
            $Report += "Last Sync : $($w.LastUpdateTime)"
            $Report += ""
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_WarrantyInformation" $_
            $Report += "Warranty information unavailable."
            $Report += ""
        }
    }
    else {
        $Report += "[ Warranty ]"
        $Report += "Warranty information is not available on this device."
        $Report += ""
    }

    # ---------------- Full Charge Capacity ----------------
    try {
        $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

        $Report += "[ Battery Health ]"

        if (-not $batteries -or $batteries.Count -eq 0) {
            $Report += "No battery capacity data found."
        }
        else {
            for ($i = 0; $i -lt $batteries.Count; $i++) {
                try {
                    if ($batteries[$i] -and $batteries[$i].FullChargedCapacity) {
                        $FullCharge = $batteries[$i].FullChargedCapacity
                        
                        $DesignCapValue = Get-DesignCapacity -Index $i
                        
                        $DesignCap = if ($DesignCapValue) { $DesignCapValue } else { "Unavailable" }
                        $Health     = "Unavailable"

                        [int64]$fc = 0
                        [int64]$dc = 0

                        if ($null -ne $DesignCapValue -and
                            [int64]::TryParse($FullCharge.ToString(), [ref]$fc) -and
                            [int64]::TryParse($DesignCapValue.ToString(), [ref]$dc) -and
                            $dc -gt 0)
                        {
                            $Health = [math]::Round(($fc / $dc) * 100, 2)
                        }
                        else {
                            $Health = "Unavailable"
                        }
                        $Report += "Battery $($i + 1)"
                        $Report += "Max Charge Now       : $FullCharge mWh"
                        $Report += "Original Capacity    : $DesignCap mWh"
                        $Report += "Health               : $Health%"
                        $Report += ""
                    }
                    else {
                        $Report += "Battery $($i + 1) - Data unavailable"
                        $Report += ""
                    }
                }
                catch {
                    Log-BatteryError "root\wmi" "BatteryFullChargedCapacity" $_ $i $($batteries[$i].InstanceName) $null
                    $Report += "Battery $($i + 1) - Could not read"
                    $Report += ""
                }
            }
        }
    }
    catch {
        Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
        $Report += "Battery capacity information unavailable."
    }

    # -------- Comprehensive Battery Analysis --------
    try {
        $Report += "[ Battery Health Analysis ]"
        $Report += ""
        
        $reportCycles = 0
        if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
            try {
                $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue
                $reportCycles = ([string]$odo.Battery_cycles -replace '\D','')
                $Report += "Charge Cycles : $reportCycles"
                $Report += ""
            }
            catch {
                $Report += "Charge cycle data unavailable."
                $Report += ""
            }
        }
        
        try {
            $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
            
            if ($batteries -and $batteries.Count -gt 0) {
                $index = 0
                foreach ($bat in $batteries) {
                    try {
                        $FullCharge = Get-SafeWmiProperty -Object $bat -PropertyName "FullChargedCapacity"
                        
                        $DesignCapValue  = Get-DesignCapacity -Index $index
                        $Manufacturer    = Get-BatteryManufacturer -Index $index
                        
                        $classification = Get-BatteryClassification -DesignCapacity $DesignCapValue -FullChargeCapacity $FullCharge -CycleCount $reportCycles -Manufacturer $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })
                        
                        $Report += ""
                        $Report += "Battery $($index + 1)"
                        $Report += "  Manufacturer   : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })"
                        $Report += "  Health         : $($classification.HealthPercent)%"
                        $Report += "  Rating         : $($classification.Classification)"
                        # DEPRECATED (v2.0): Genuine field removed from report — EU Battery Regulation compliance.
                        $Report += "  Manufacturer   : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })  [INFO]"
                        $Report += "  Summary        : $($classification.Message)"
                        $Report += ""
                        
                        $index++
                    }
                    catch {
                        $Report += "Battery $($index + 1) - Could not read"
                        $Report += ""
                    }
                }
            }
            else {
                $Report += "No battery data found."
                $Report += ""
            }
        }
        catch {
            $Report += "Battery analysis could not be completed."
            $Report += ""
        }
    }
    catch {
        $Report += "[ Battery Health Analysis ]"
        $Report += "Could not generate."
        $Report += ""
    }

    # Save report
    $Report | Out-File $reportPath -Encoding UTF8

    # Save issues if any exist
    if ($script:ErrorLog.Count -gt 0) {
        $ErrorOutput = @()
        foreach ($err in $script:ErrorLog) {
            $ErrorOutput += ("=" * 70)
            $ErrorOutput += "Time      : $($err.TimeStamp)"
            $ErrorOutput += "Component : $($err.Namespace) / $($err.Class)"
            $ErrorOutput += "Detail    : $($err.Message)"
            $ErrorOutput += "Location  : line $($err.ScriptLineNumber)"
            if ($err.Context) {
                $ErrorOutput += "Context   : $($err.Context)"
            }
            if ($err.BatteryDetails) {
                $ErrorOutput += "Battery   :"
                $ErrorOutput += "  Max Charge Now : $($err.BatteryDetails.FullChargedCapacity)"
                $ErrorOutput += "  Original Cap.  : $($err.BatteryDetails.DesignedCapacity)"
            }
            $ErrorOutput += "Trace     : $($err.StackTrace)"
            $ErrorOutput += ""
        }
        $ErrorOutput | Out-File $errorPath -Encoding UTF8
    }

    Show-Header
    Write-Host "Report saved to:" -ForegroundColor Green
    Write-Host $reportPath

    if ($script:ErrorLog.Count -gt 0) {
        Write-Host ""
        Write-Host "Issue log saved to:" -ForegroundColor Yellow
        Write-Host $errorPath
    }

    Read-Host "Press ENTER"
}

function Write-BatteryTrendLog {
    <#
    .SYNOPSIS
    Appends a battery health snapshot to the persistent trend log CSV.

    .DESCRIPTION
    Called silently on every script run after Get-BatteryAlertState has
    populated $script:BatteryAlert. Writes one row per battery per run to:
      $env:LOCALAPPDATA\LenovoUtility\battery_trend.csv

    Columns: Timestamp, BatteryIndex, BatteryID, FullCharge, DesignCapacity,
             HealthPct, CycleCount

    Duplicate suppression: if the most recent entry for a given BatteryIndex
    was written within the last 20 hours, the row is skipped — so rapid
    successive runs (e.g. testing) don't inflate the log.
    #>

    try {
        $logDir  = Join-Path $env:LOCALAPPDATA "LenovoUtility"
        $logPath = Join-Path $logDir "battery_trend.csv"

        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        # Load existing rows for duplicate suppression
        $existing = @()
        if (Test-Path $logPath) {
            try { $existing = @(Import-Csv $logPath) } catch {}
        }

        $now = Get-Date

        # Collect per-battery data from ACPI (always available, source-agnostic)
        $acpiBatteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

        # Also try Lenovo_Battery for cycle count
        $lbBatteries = @()
        if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
            try {
                $lbBatteries = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction SilentlyContinue)
            } catch {}
        }

        # Fallback cycle count from Lenovo_Odometer
        $odoObj = $null
        if ($lbBatteries.Count -eq 0 -and (Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
            try {
                $odoObj = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue
            } catch {}
        }

        $newRows = @()

        for ($i = 0; $i -lt $acpiBatteries.Count; $i++) {
            try {
                [int64]$fc = 0
                if (-not ([int64]::TryParse(
                    (Get-SafeWmiProperty -Object $acpiBatteries[$i] -PropertyName "FullChargedCapacity"),
                    [ref]$fc))) { continue }

                $dcRaw = Get-DesignCapacity -Index $i
                [int64]$dc = 0
                if (-not ([int64]::TryParse(($dcRaw -replace '\D',''), [ref]$dc)) -or $dc -le 0) { continue }

                # Cap at 100 — FullCharge can exceed DesignCapacity on some firmware
                # (e.g. BYD L24B4PC0 reports ~85160 mWh against 80000 mWh design).
                # Values above 100% are firmware noise, not genuine overcapacity.
                $healthPct = [math]::Min([math]::Round(($fc / $dc) * 100, 2), 100.0)

                # Cycle count — prefer Lenovo_Battery, fall back to Lenovo_Odometer
                $cycles = ""
                if ($i -lt $lbBatteries.Count) {
                    $rawCyc = Get-SafeWmiProperty -Object $lbBatteries[$i] -PropertyName "CycleCount"
                    [int64]$cyc = 0
                    if ([int64]::TryParse(($rawCyc -replace '\D',''), [ref]$cyc)) { $cycles = $cyc }
                } elseif ($odoObj) {
                    [int64]$cyc = 0
                    if ([int64]::TryParse(([string]$odoObj.Battery_cycles -replace '\D',''), [ref]$cyc)) { $cycles = $cyc }
                }

                # Battery ID from alert cache or fallback label
                $batteryID = "Battery $i"
                if ($script:BatteryAlert -and $script:BatteryAlert.Batteries -and $i -lt $script:BatteryAlert.Batteries.Count) {
                    $cachedID = $script:BatteryAlert.Batteries[$i].BatteryID
                    if ($cachedID -and $cachedID -ne "Unavailable") { $batteryID = $cachedID }
                }

                # Duplicate suppression — skip if logged within last 20 hours for this index
                $recentEntry = $existing |
                    Where-Object { $_.BatteryIndex -eq $i } |
                    Sort-Object Timestamp |
                    Select-Object -Last 1

                if ($recentEntry) {
                    try {
                        $lastTime = [datetime]::Parse($recentEntry.Timestamp)
                        if (($now - $lastTime).TotalHours -lt 20) { continue }
                    } catch {}
                }

                $newRows += [PSCustomObject]@{
                    Timestamp      = $now.ToString("yyyy-MM-dd HH:mm:ss")
                    BatteryIndex   = $i
                    BatteryID      = $batteryID
                    FullCharge     = $fc
                    DesignCapacity = $dc
                    HealthPct      = $healthPct
                    CycleCount     = $cycles
                }
            } catch {}
        }

        if ($newRows.Count -gt 0) {
            $writeHeader = -not (Test-Path $logPath)
            $newRows | Export-Csv -Path $logPath -Append -NoTypeInformation -Force
            # If file was just created, it already has a header. If appending to existing,
            # Export-Csv -Append on PowerShell 5 doesn't re-write header — correct behaviour.
        }
    } catch {}
}

function Show-BatteryHealthTrend {
    <#
    .SYNOPSIS
    Displays the battery health trend from the persistent log, including
    degradation rate and projection to the 80% replacement threshold.
    #>

    Clear-Host
    Write-Host "==============================================="
    Write-Host "  BATTERY HEALTH TREND"
    Write-Host "==============================================="
    Write-Host ""

    $logPath = Join-Path $env:LOCALAPPDATA "LenovoUtility\battery_trend.csv"

    if (-not (Test-Path $logPath)) {
        Write-Host "No history recorded yet." -ForegroundColor Yellow
        Write-Host "This tool quietly saves a snapshot each time you run it." -ForegroundColor DarkGray
        Write-Host "Run it a few more times over the coming days or weeks to build a trend." -ForegroundColor DarkGray
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    $rows = @()
    try { $rows = @(Import-Csv $logPath) } catch {
        Write-Host "Could not read health history." -ForegroundColor Red
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    if ($rows.Count -eq 0) {
        Write-Host "History file exists but contains no data yet." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    $indices = $rows | Select-Object -ExpandProperty BatteryIndex -Unique | Sort-Object

    foreach ($idx in $indices) {
        $battRows = @($rows | Where-Object { $_.BatteryIndex -eq $idx } | Sort-Object Timestamp)

        $battID = $battRows[-1].BatteryID
        Write-Host "Battery: $battID" -ForegroundColor Cyan
        Write-Host ("-" * 50)

        Write-Host ("  {0,-20} {1,8} {2,8} {3,10}" -f "Date", "Health%", "MaxChg", "Cycles")
        Write-Host ("  {0,-20} {1,8} {2,8} {3,10}" -f "--------------------", "-------", "-------", "------")

        foreach ($row in $battRows) {
            [double]$hp = 0
            [double]::TryParse($row.HealthPct, [ref]$hp) | Out-Null
            $color = if ($hp -ge 80) { "Green" } elseif ($hp -ge 70) { "Yellow" } elseif ($hp -ge 60) { "Magenta" } else { "Red" }
            $cycleDisplay = if ($row.CycleCount) { $row.CycleCount } else { "N/A" }
            Write-Host ("  {0,-20} " -f $row.Timestamp) -NoNewline
            Write-Host ("{0,7}% " -f $row.HealthPct) -ForegroundColor $color -NoNewline
            Write-Host ("{0,8} {1,10}" -f $row.FullCharge, $cycleDisplay)
        }

        Write-Host ""

        $parsed = @()
        foreach ($row in $battRows) {
            [double]$hp = 0
            $ts = $null
            try { $ts = [datetime]::Parse($row.Timestamp) } catch { continue }
            if (-not [double]::TryParse($row.HealthPct, [ref]$hp)) { continue }
            $parsed += [PSCustomObject]@{ Timestamp = $ts; HealthPct = $hp }
        }

        if ($parsed.Count -lt 2) {
            Write-Host "  Not enough data yet — run the tool again on a different day." -ForegroundColor DarkGray
            Write-Host ""
            continue
        }

        $first = $parsed[0]
        $last  = $parsed[-1]
        $spanDays = ($last.Timestamp - $first.Timestamp).TotalDays

        if ($spanDays -lt 1) {
            Write-Host "  All records from the same day — come back tomorrow for trend data." -ForegroundColor DarkGray
            Write-Host ""
            continue
        }

        if ($spanDays -lt 14) {
            Write-Host "  Need at least 14 days of data for a reliable trend." -ForegroundColor DarkGray
            Write-Host "  So far: $([math]::Round($spanDays)) days  ($($parsed.Count) records)" -ForegroundColor DarkGray
            Write-Host "  Current Health   : " -NoNewline
            $hpColor = if ($last.HealthPct -ge 80) { "Green" } elseif ($last.HealthPct -ge 70) { "Yellow" } elseif ($last.HealthPct -ge 60) { "Magenta" } else { "Red" }
            Write-Host ("{0:N2}%" -f $last.HealthPct) -ForegroundColor $hpColor
            Write-Host ""
            continue
        }

        $totalDrop    = $first.HealthPct - $last.HealthPct
        $dropPerDay   = $totalDrop / $spanDays
        $dropPerMonth = [math]::Round($dropPerDay * 30.44, 2)
        $currentHP    = $last.HealthPct

        $rateColor = if ($dropPerMonth -le 1.5) { "Green" } elseif ($dropPerMonth -le 3.0) { "Yellow" } else { "Red" }
        $rateLabel = if ($dropPerMonth -le 1.5) { "Normal" } elseif ($dropPerMonth -le 3.0) { "Elevated" } else { "Accelerated" }

        Write-Host "  Wear Rate        : " -NoNewline
        Write-Host ("{0:N2}%/month  ({1})" -f $dropPerMonth, $rateLabel) -ForegroundColor $rateColor

        Write-Host "  Tracked for      : $([math]::Round($spanDays)) days  ($($parsed.Count) records)"
        Write-Host "  Current Health   : " -NoNewline
        $hpColor = if ($currentHP -ge 80) { "Green" } elseif ($currentHP -ge 70) { "Yellow" } elseif ($currentHP -ge 60) { "Magenta" } else { "Red" }
        Write-Host ("{0:N2}%" -f $currentHP) -ForegroundColor $hpColor

        if ($dropPerMonth -gt 0 -and $currentHP -gt 80) {
            $monthsTo80 = [math]::Round(($currentHP - 80) / $dropPerMonth, 1)
            $dateTo80   = (Get-Date).AddDays($monthsTo80 * 30.44)
            Write-Host "  Est. 80% reached : in ~$monthsTo80 months  (around $($dateTo80.ToString('MMM yyyy')))" -ForegroundColor $rateColor
        } elseif ($currentHP -le 80) {
            Write-Host "  Battery is already at or below 80% — consider replacing it soon." -ForegroundColor Yellow
        } else {
            Write-Host "  Wear rate is too low to project a timeline." -ForegroundColor DarkGray
        }

        Write-Host ""
    }

    Write-Host "Saved to: $logPath" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press ENTER"
}

function Get-BatteryCapacityHistory {
    <#
    .SYNOPSIS
    Parses the capacity history table from powercfg /batteryreport /XML.

    .DESCRIPTION
    The battery report XML contains a CapacityHistory section that records
    full charge capacity and design capacity at regular intervals (typically
    weekly). This gives a real observed degradation curve rather than a
    single snapshot, enabling slope-based age estimation.

    Returns an array of PSCustomObjects sorted oldest-first:
      Timestamp       - DateTime of the measurement
      FullCharge      - Full charge capacity in mWh
      DesignCapacity  - Design capacity in mWh
      HealthPct       - FullCharge / DesignCapacity * 100 (rounded to 2dp)

    Returns an empty array if the report cannot be generated or parsed.
    The XML file is written to %TEMP% and deleted immediately after parsing.
    #>

    $entries = @()

    try {
        $xmlPath = Join-Path $env:TEMP "battery_history_$PID.xml"
        $null = & powercfg /batteryreport /XML /OUTPUT $xmlPath 2>$null

        if (-not (Test-Path $xmlPath)) { return $entries }

        [xml]$report = Get-Content $xmlPath -ErrorAction Stop
        Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue

        $historyNodes = $report.BatteryReport.History.HistoryEntry
        if (-not $historyNodes) { return $entries }

        foreach ($node in $historyNodes) {
            try {
                [int64]$fc = 0
                [int64]$dc = 0
                if (-not ([int64]::TryParse($node.FullChargeCapacity, [ref]$fc))) { continue }
                if (-not ([int64]::TryParse($node.DesignCapacity,     [ref]$dc))) { continue }
                if ($dc -le 0 -or $fc -le 0) { continue }

                # powercfg uses ISO 8601 UTC timestamps in StartDate.
                # Use LocalStartDate when available for a human-readable local time.
                $ts = $null
                $tsStr = if ($node.LocalStartDate) { $node.LocalStartDate } else { $node.StartDate }
                try { $ts = [datetime]::Parse($tsStr) } catch { continue }

                # CycleCount is present per entry — capture it for the newest entry
                [int64]$cc = 0
                [int64]::TryParse($node.CycleCount, [ref]$cc) | Out-Null

                $entries += [PSCustomObject]@{
                    Timestamp      = $ts
                    FullCharge     = $fc
                    DesignCapacity = $dc
                    HealthPct      = [math]::Round(($fc / $dc) * 100, 2)
                    CycleCount     = $cc
                }
            }
            catch {}
        }

        # Sort oldest first
        $entries = $entries | Sort-Object Timestamp
    }
    catch {}

    return $entries
}

function Get-BatteryAgeEstimate {
    <#
    .SYNOPSIS
    Estimates battery age in years using three independent methods and cross-validates them.

    Method 1 - Cycle-based:
      Average laptop user does ~200-300 cycles/year (partial cycles, plugged in sometimes).
      Conservative baseline: 250 cycles/year.
      Source: ASUS, Lenovo, and independent battery longevity studies.

    Method 2 - Capacity-based:
      Li-ion batteries lose ~20% capacity after ~300 cycles (approx 1 year typical use).
      Capacity loss is non-linear: faster early, stabilizes mid-life, accelerates near EOL.
      We use a simplified linear model: 100% health = 0 years, 0% health = ~5 years.
      This is conservative and matches real-world laptop battery observations.

    Method 3 - History-based (slope):
      Uses the capacity history from powercfg /batteryreport /XML to compute the
      observed degradation rate (% per day) across all recorded entries. Extrapolates
      how long ago health was at 100% using that real measured slope. This is the
      most accurate method when sufficient history is available (>= 4 weeks of data).

    If methods agree within 1 year, confidence is High.
    If they diverge by 1-2 years, confidence is Medium with a likely explanation.
    If they diverge by more than 2 years, confidence is Low with usage pattern notes.
    #>
    param(
        [object]$CycleCount,
        [object]$HealthPercent,
        [object[]]$CapacityHistory = @()
    )

    $result = [PSCustomObject]@{
        CycleBasedYears    = $null
        CapacityBasedYears = $null
        HistoryBasedYears  = $null
        EstimatedYears     = $null
        Confidence         = "Unknown"
        Note               = ""
        HasCycleData       = $false
        HasCapacityData    = $false
        HasHistoryData     = $false
        HistoryEntryCount  = 0
        HistorySpanDays    = 0
    }

    # --- Method 1: Cycle-based estimate ---
    [int64]$cycles = 0
    if ($CycleCount -and [int64]::TryParse(($CycleCount -replace '\D',''), [ref]$cycles) -and $cycles -gt 0) {
        # 250 cycles/year is the conservative average for a typical laptop user
        $result.CycleBasedYears = [math]::Round($cycles / 250, 1)
        $result.HasCycleData = $true
    }

    # --- Method 2: Capacity-based estimate ---
    [double]$health = 0
    if ($HealthPercent -and [double]::TryParse([string]$HealthPercent, [ref]$health) -and $health -gt 0) {
        $capacityLoss = 100 - $health
        $result.CapacityBasedYears = [math]::Round($capacityLoss / 4.0, 1)
        if ($result.CapacityBasedYears -lt 0) { $result.CapacityBasedYears = 0 }
        $result.HasCapacityData = $true
    }

    # --- Method 3: History-based slope estimate ---
    # Requires at least 2 entries spanning at least 7 days to produce a
    # meaningful slope. With fewer entries the slope is too noisy to trust.
    if ($CapacityHistory -and $CapacityHistory.Count -ge 2) {
        $result.HistoryEntryCount = $CapacityHistory.Count

        $oldest = $CapacityHistory[0]
        $newest = $CapacityHistory[-1]
        $spanDays = ($newest.Timestamp - $oldest.Timestamp).TotalDays
        $result.HistorySpanDays = [math]::Round($spanDays, 0)

        if ($spanDays -ge 7) {
            $healthDrop = $oldest.HealthPct - $newest.HealthPct

            if ($healthDrop -gt 0) {
                $ratePerDay = $healthDrop / $spanDays
                $totalDaysToReach100 = (100 - $newest.HealthPct) / $ratePerDay
                $result.HistoryBasedYears = [math]::Round($totalDaysToReach100 / 365, 1)
                if ($result.HistoryBasedYears -lt 0) { $result.HistoryBasedYears = 0 }
                $result.HasHistoryData = $true
            }
            elseif ($healthDrop -le 0) {
                # No measurable degradation in the history window
                $result.HistoryBasedYears = 0
                $result.HasHistoryData = $true
            }
        }
    }

    # --- Cross-validate and determine confidence ---
    $availableEstimates = @()
    if ($result.HasCycleData)    { $availableEstimates += $result.CycleBasedYears }
    if ($result.HasCapacityData) { $availableEstimates += $result.CapacityBasedYears }
    if ($result.HasHistoryData)  { $availableEstimates += $result.HistoryBasedYears }

    if ($availableEstimates.Count -eq 0) {
        $result.Confidence = "Unknown"
        $result.Note = "Insufficient data. Neither cycle count, health percentage, nor capacity history is available."
        return $result
    }

    if ($availableEstimates.Count -eq 1) {
        $result.EstimatedYears = $availableEstimates[0]
        $result.Confidence = "Medium"
        if ($result.HasHistoryData) {
            $result.Note = "History-based estimate only ($($result.HistoryEntryCount) entries, $($result.HistorySpanDays) days). No other data available for cross-validation."
        }
        elseif ($result.HasCycleData) {
            $result.Note = "Cycle-based estimate only. No capacity data available for cross-validation."
        }
        else {
            $result.Note = "Capacity-based estimate only. No cycle data available for cross-validation."
        }
        return $result
    }

    $avg  = ($availableEstimates | Measure-Object -Average).Average
    $max  = ($availableEstimates | Measure-Object -Maximum).Maximum
    $min  = ($availableEstimates | Measure-Object -Minimum).Minimum
    $diff = $max - $min

    # Weight history more heavily when span is long (>= 90 days)
    if ($result.HasHistoryData -and $result.HistorySpanDays -ge 90) {
        $otherCount    = $availableEstimates.Count - 1
        $historyWeight = 0.50
        $otherWeight   = if ($otherCount -gt 0) { 0.50 / $otherCount } else { 0 }
        $weighted = $result.HistoryBasedYears * $historyWeight
        if ($result.HasCycleData)    { $weighted += $result.CycleBasedYears    * $otherWeight }
        if ($result.HasCapacityData) { $weighted += $result.CapacityBasedYears * $otherWeight }
        $result.EstimatedYears = [math]::Round($weighted, 1)
    }
    else {
        $result.EstimatedYears = [math]::Round($avg, 1)
    }

    if ($diff -le 1.0) {
        $result.Confidence = "High"
        $methodCount = $availableEstimates.Count
        $result.Note = "All $methodCount methods agree. Estimate is reliable."
    }
    elseif ($diff -le 2.0) {
        $result.Confidence = "Medium"
        if ($result.HasHistoryData -and $result.HasCycleData -and
            $result.HistoryBasedYears -lt $result.CycleBasedYears) {
            $result.Note = "History shows slower degradation than cycles suggest. Battery may have been partially cycled or used in cool conditions."
        }
        elseif ($result.HasCapacityData -and $result.HasCycleData -and
                $result.CycleBasedYears -lt $result.CapacityBasedYears) {
            $result.Note = "Capacity has degraded more than cycles suggest. Battery may have experienced high temperatures or prolonged storage at full charge."
        }
        else {
            $result.Note = "Methods show minor divergence. Estimate is approximate."
        }
    }
    else {
        $result.Confidence = "Low"
        if ($result.HasHistoryData -and $result.HistorySpanDays -lt 30) {
            $result.Note = "Significant divergence. Capacity history span is short ($($result.HistorySpanDays) days) -- slope estimate may not be representative yet."
        }
        elseif ($result.HasCycleData -and $result.HasCapacityData -and
                $result.CycleBasedYears -lt $result.CapacityBasedYears) {
            $result.Note = "Significant divergence: capacity loss greatly exceeds cycle count. Likely calendar aging from long-term storage or heat exposure."
        }
        else {
            $result.Note = "Significant divergence between methods. Actual age may vary considerably."
        }
    }

    return $result
}

function BatteryAgeEstimation {
    Show-Header
    Write-Host "[ Battery Age Estimation ]"
    Write-Host ""

    $cycles = $null
    $health = $null

    # --- Get cycle count from Lenovo EC ---
    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
        try {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            $cycles = ([string]$odo.Battery_cycles -replace '\D','')
            Write-Host "Charge Cycles            : $cycles"
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Odometer" $_
            Write-Host "Charge Cycles            : Not available right now" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Charge Cycles            : Not available on this device" -ForegroundColor Yellow
    }

    # --- Get capacity and calculate health ---
    try {
        $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
        if ($batteries -and $batteries.Count -gt 0) {
            $fcRaw = Get-SafeWmiProperty -Object $batteries[0] -PropertyName "FullChargedCapacity"
            $dcRaw = Get-DesignCapacity -Index 0

            [int64]$fc = 0
            [int64]$dc = 0
            if ($null -ne $fcRaw -and $null -ne $dcRaw -and
                [int64]::TryParse($fcRaw.ToString(), [ref]$fc) -and
                [int64]::TryParse($dcRaw.ToString(), [ref]$dc) -and
                $dc -gt 0) {
                $health = [math]::Round(($fc / $dc) * 100, 2)
                Write-Host "Max Charge Now           : $fc mWh"
                Write-Host "Original Capacity        : $dc mWh"
                Write-Host "Battery Health           : $health%"
            }
            else {
                Write-Host "Battery Health           : Not available" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Battery Health           : Not available (no battery found)" -ForegroundColor Yellow
        }
    }
    catch {
        Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
        Write-Host "Battery Health           : Not available" -ForegroundColor Yellow
    }

    # --- Collect capacity history from powercfg battery report ---
    Write-Host ""
    Write-Host "[ Capacity History ]"
    Write-Host "Reading battery history..." -ForegroundColor Cyan
    $history = Get-BatteryCapacityHistory

    if ($history -and $history.Count -ge 2) {
        $spanDays = [math]::Round(($history[-1].Timestamp - $history[0].Timestamp).TotalDays, 0)
        Write-Host "  Records  : $($history.Count)  (spanning $spanDays days)" -ForegroundColor DarkGray
        Write-Host "  Oldest   : $($history[0].Timestamp.ToString('yyyy-MM-dd'))  Health: $($history[0].HealthPct)%" -ForegroundColor DarkGray
        Write-Host "  Newest   : $($history[-1].Timestamp.ToString('yyyy-MM-dd'))  Health: $($history[-1].HealthPct)%" -ForegroundColor DarkGray
        $totalDrop = [math]::Round($history[0].HealthPct - $history[-1].HealthPct, 2)
        if ($totalDrop -gt 0) {
            Write-Host "  Degraded : $totalDrop% over $spanDays days" -ForegroundColor DarkGray
        }
        elseif ($totalDrop -le 0) {
            Write-Host "  Degraded : No measurable degradation in recorded history" -ForegroundColor DarkGray
        }

        if ($null -eq $cycles -and $history[-1].CycleCount -gt 0) {
            $cycles = $history[-1].CycleCount
            Write-Host "  Charge Cycles (from history): $cycles" -ForegroundColor DarkGray
        }
    }
    elseif ($history -and $history.Count -eq 1) {
        Write-Host "  Only 1 record found — not enough to calculate a trend." -ForegroundColor Yellow
    }
    else {
        Write-Host "  No history available." -ForegroundColor Yellow
    }

    Write-Host ""

    # --- Run estimation ---
    $estimate = Get-BatteryAgeEstimate -CycleCount $cycles -HealthPercent $health -CapacityHistory $history

    if ($estimate.Confidence -eq "Unknown") {
        Write-Host "[ Age Estimate ]" -ForegroundColor Yellow
        Write-Host "Not enough data to estimate battery age." -ForegroundColor Yellow
        Write-Host $estimate.Note -ForegroundColor Yellow
    }
    else {
        Write-Host "[ Age Estimate ]"
        Write-Host ("-" * 50)

        if ($estimate.HasCycleData) {
            Write-Host "Based on charge cycles   : ~$($estimate.CycleBasedYears) year(s)"
        }
        if ($estimate.HasCapacityData) {
            Write-Host "Based on capacity loss   : ~$($estimate.CapacityBasedYears) year(s)"
        }
        if ($estimate.HasHistoryData) {
            Write-Host "Based on history trend   : ~$($estimate.HistoryBasedYears) year(s)  ($($estimate.HistoryEntryCount) records, $($estimate.HistorySpanDays) days)"
        }

        Write-Host ""

        $confidenceColor = switch ($estimate.Confidence) {
            "High"   { "Green" }
            "Medium" { "Yellow" }
            "Low"    { "Magenta" }
            default  { "White" }
        }

        Write-Host "Estimated Age            : " -NoNewline
        Write-Host "~$($estimate.EstimatedYears) year(s)" -ForegroundColor $confidenceColor
        Write-Host "Confidence               : " -NoNewline
        Write-Host $estimate.Confidence -ForegroundColor $confidenceColor
        Write-Host ""
        Write-Host "Note: $($estimate.Note)" -ForegroundColor $confidenceColor
        Write-Host ""
        Write-Host "This is a statistical estimate based on typical battery aging." -ForegroundColor DarkGray
        Write-Host "Actual age may vary depending on usage habits and temperature." -ForegroundColor DarkGray
    }

    Write-Host ""
    Read-Host "Press ENTER"
}

function Show-About {
    Show-Header
    Write-Host "[ About ]"
    Write-Host ""

    Write-Host "Lenovo Device Health Check" -ForegroundColor Cyan
    Write-Host "Version 1.1" -ForegroundColor White
    Write-Host ""

    Write-Host "Built with Windows PowerShell, ChatGPT GPT 5.4, and Claude Sonnet 4.6" -ForegroundColor DarkGray
    Write-Host ""

    Write-Host "A free, read-only tool for checking battery health, age," -ForegroundColor Gray
    Write-Host "and memory on Lenovo consumer devices." -ForegroundColor Gray
    Write-Host "It reads information from your device — it does not" -ForegroundColor Gray
    Write-Host "change any settings or send any data anywhere." -ForegroundColor Gray
    Write-Host ""

    Write-Host ("=" * 47)
    Write-Host ""

    Write-Host "Legal" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This tool is not affiliated with or endorsed" -ForegroundColor DarkGray
    Write-Host "by Lenovo." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "It makes no modifications to your device," -ForegroundColor DarkGray
    Write-Host "firmware, or settings." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Third-party battery replacements are a consumer" -ForegroundColor DarkGray
    Write-Host "right under EU Battery Regulation (2027)." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Battery age estimates are approximations only." -ForegroundColor DarkGray
    Write-Host "Actual age may vary based on usage and service history." -ForegroundColor DarkGray
    Write-Host ""

    Write-Host ("=" * 47)
    Write-Host ""

    Write-Host "Privacy" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Everything runs locally on your device." -ForegroundColor DarkGray
    Write-Host "No data is sent, collected, or stored anywhere" -ForegroundColor DarkGray
    Write-Host "other than the optional report you can save locally." -ForegroundColor DarkGray
    Write-Host ""

    Read-Host "Press ENTER"
}


function Test-LenovoNamespace {
    try {
        $lenovoNS = Get-CimInstance -CimSession $script:CimSession -Namespace root -ClassName __Namespace -ErrorAction SilentlyContinue | Where-Object Name -eq "Lenovo"
        return $null -ne $lenovoNS
    }
    catch {
        # Fallback: try to query a Lenovo class directly
        try {
            $null = Get-CimClass -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue
            return $true
        }
        catch {
            return $false
        }
    }
}

function Test-CommercialVantage {
    <#
    .SYNOPSIS
    Detects whether Lenovo Commercial Vantage is installed on this system.

    .DESCRIPTION
    Commercial Vantage is the enterprise-focused Lenovo management app that,
    when installed on top of SIF, deploys the CDRT MOF and registers the
    Lenovo_Odometer WMI class fields used for CPU uptime, shock event, and
    thermal event tracking (Options 11 / CDRT Odometer).

    Detection strategy (most-to-least reliable):
      1. Registry uninstall key scan (HKLM Uninstall hives, 64-bit and 32-bit).
         Matches on "Commercial Vantage" in DisplayName without invoking
         Win32_Product, which triggers MSI consistency checks.
      2. Lenovo-specific registry key written by the Commercial Vantage installer.
      3. Commercial Vantage service presence (LenovoVantageService).
    Returns $true if Commercial Vantage appears to be installed, $false otherwise.
    #>

    # Strategy 1: Registry uninstall key scan.
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    try {
        foreach ($regPath in $uninstallPaths) {
            if (-not (Test-Path $regPath)) { continue }
            $found = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue |
                ForEach-Object { Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue } |
                Where-Object { $_.DisplayName -match "Commercial Vantage" } |
                Select-Object -First 1
            if ($found) { return $true }
        }
    } catch {}

    # Strategy 2: Lenovo-specific registry key written by the Commercial Vantage installer.
    try {
        if (Test-Path "HKLM:\SOFTWARE\Lenovo\CommercialVantage") { return $true }
        if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Lenovo\CommercialVantage") { return $true }
    } catch {}

    # Strategy 3: Commercial Vantage service presence.
    try {
        $svc = Get-Service -Name "LenovoVantageService" -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -ne $null } |
            Select-Object -First 1
        if ($svc) { return $true }
    } catch {}

    return $false
}

function Test-SifInstalled {
    <#
    .SYNOPSIS
    Detects whether the Lenovo System Interface Foundation (SIF) driver package
    is installed on this system.

    .DESCRIPTION
    SIF (formerly "Lenovo System Interface Foundation" or "System Interface
    Foundation") is the Lenovo driver package that registers the root\Lenovo
    WMI namespace and the EC-facing WMI classes (Lenovo_Odometer,
    Lenovo_BiosSetting, Lenovo_WarrantyInformation, etc.).

    Options 2 (Lenovo Battery Cycles) and 5 (Comprehensive Battery Analysis)
    depend on SIF. Without it, the root\Lenovo namespace does not exist and
    every WMI call against it will fail with 0x8004100E (Invalid Namespace).

    Detection strategy (most-to-least reliable):
      1. Registry uninstall key scan (HKLM Uninstall hives, 64-bit and 32-bit).
         Reads DisplayName values directly from the Windows installer database
         without invoking Win32_Product, which triggers a full MSI consistency
         check and can cause unintended repair actions on every installed package.
      2. Lenovo-specific registry key written by the SIF installer.
      3. SIF service presence (ImControllerService / sifservice).
    Returns $true if SIF appears to be installed, $false otherwise.
    #>

    # Strategy 1: Registry uninstall key scan.
    # Checks both the 64-bit and 32-bit (WOW6432Node) uninstall hives so the
    # detection works regardless of whether SIF was installed as a 32-bit or
    # 64-bit package. Reading registry values is instantaneous and has no
    # side effects, unlike Win32_Product which triggers MSI reconfiguration.
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    try {
        foreach ($regPath in $uninstallPaths) {
            if (-not (Test-Path $regPath)) { continue }
            $found = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue |
                ForEach-Object { Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue } |
                Where-Object { $_.DisplayName -match "System Interface Foundation|Lenovo Service Bridge" } |
                Select-Object -First 1
            if ($found) { return $true }
        }
    } catch {}

    # Strategy 2: Lenovo-specific registry key written by the SIF installer.
    try {
        if (Test-Path "HKLM:\SOFTWARE\Lenovo\SystemInterfaceFoundation") { return $true }
        if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Lenovo\SystemInterfaceFoundation") { return $true }
    } catch {}

    # Strategy 3: SIF service presence.
    try {
        $svc = Get-Service -Name "ImControllerService","sifservice" -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -ne $null } |
            Select-Object -First 1
        if ($svc) { return $true }
    } catch {}

    return $false
}

function Test-WmiClass {
    param(
        [string]$Namespace = "root\cimv2",
        [string]$ClassName
    )
    
    try {
        $class = Get-CimClass -Namespace $Namespace -ClassName $ClassName -ErrorAction SilentlyContinue
        return $null -ne $class
    }
    catch {
        return $false
    }
}

function Get-SafeWmiProperty {
    param(
        [object]$Object,
        [string]$PropertyName,
        [object]$DefaultValue = "Unavailable"
    )
    
    try {
        if ($Object -and $null -ne $Object.$PropertyName) {
            return $Object.$PropertyName
        }
    }
    catch {
        # Silently fail and return default
    }
    
    return $DefaultValue
}

function Get-MemoryState {
    <#
    .SYNOPSIS
    Retrieves physical memory information from Win32_PhysicalMemory.
    
    .DESCRIPTION
    Queries system memory modules and returns capacity, speed, manufacturer info,
    and detects speed mismatches between modules.
    
    .OUTPUTS
    PSCustomObject with properties:
    - TotalRAMGB: Total system RAM in gigabytes
    - ModuleCount: Number of physical memory modules
    - SpeedMismatch: Boolean indicating if modules have different speeds
    - Severity: 0 for matched speeds, 1 for mismatched speeds
    - Modules: Array of module details (Capacity, Speed, Manufacturer)
    #>
    
    $result = [PSCustomObject]@{
        TotalRAMGB = 0
        ModuleCount = 0
        SpeedMismatch = $false
        Severity = 0
        Modules = @()
        Message = ""
    }
    
    try {
        if (-not (Test-WmiClass -ClassName "Win32_PhysicalMemory")) {
            $result.Message = "Win32_PhysicalMemory class not available on this system."
            return $result
        }
        
        $modules = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue)
        
        if (-not $modules -or $modules.Count -eq 0) {
            $result.Message = "No memory modules found."
            return $result
        }
        
        $result.ModuleCount = $modules.Count
        $speeds = @()
        $totalBytes = 0
        
        foreach ($module in $modules) {
            try {
                $capacity = Get-SafeWmiProperty -Object $module -PropertyName "Capacity"
                $speed = Get-SafeWmiProperty -Object $module -PropertyName "Speed"
                $manufacturer = Get-SafeWmiProperty -Object $module -PropertyName "Manufacturer"
                $deviceLocator = Get-SafeWmiProperty -Object $module -PropertyName "DeviceLocator"
                
                # Parse capacity to numeric value
                [int64]$capacityBytes = 0
                if ($capacity -and $capacity -ne "Unavailable" -and [int64]::TryParse($capacity.ToString(), [ref]$capacityBytes)) {
                    $totalBytes += $capacityBytes
                }
                
                # Parse speed to numeric value
                [int32]$speedValue = 0
                if ($speed -and $speed -ne "Unavailable" -and [int32]::TryParse($speed.ToString(), [ref]$speedValue)) {
                    $speeds += $speedValue
                }
                
                $result.Modules += [PSCustomObject]@{
                    DeviceLocator = $deviceLocator
                    CapacityGB = if ($capacityBytes -gt 0) { [math]::Round($capacityBytes / 1GB, 2) } else { "Unavailable" }
                    CapacityBytes = $capacityBytes
                    SpeedMHz = if ($speedValue -gt 0) { $speedValue } else { "Unavailable" }
                    Manufacturer = $manufacturer
                }
            }
            catch {
                Log-Error "root\cimv2" "Win32_PhysicalMemory" $_
            }
        }
        
        # Calculate total RAM in GB
        if ($totalBytes -gt 0) {
            $result.TotalRAMGB = [math]::Round($totalBytes / 1GB, 2)
        }
        
        # Detect speed mismatch
        if ($speeds.Count -gt 1) {
            $uniqueSpeeds = $speeds | Select-Object -Unique
            if ($uniqueSpeeds.Count -gt 1) {
                $result.SpeedMismatch = $true
                $result.Severity = 1
                $result.Message = "Warning: Memory modules are running at different speeds. This may reduce performance."
            }
            else {
                $result.Message = "All memory modules are running at the same speed."
            }
        }
        elseif ($speeds.Count -eq 1) {
            $result.Message = "Single memory module detected."
        }
        else {
            $result.Message = "Could not read memory speed information."
        }
        
        return $result
    }
    catch {
        Log-Error "root\cimv2" "Win32_PhysicalMemory" $_
        $result.Message = "Could not read memory information."
        return $result
    }
}

function MemoryInfo {
    Show-Header
    Write-Host "[ Memory (RAM) ]"
    Write-Host ""

    $memoryState = Get-MemoryState

    if (-not [string]::IsNullOrEmpty($memoryState.Message) -and $memoryState.ModuleCount -eq 0) {
        $failColor = if ($memoryState.Message -match "Failed") { "Red" } else { "Yellow" }
        Write-Host $memoryState.Message -ForegroundColor $failColor
    }
    else {
        Write-Host "Total RAM              : $($memoryState.TotalRAMGB) GB"
        Write-Host "Number of Modules      : $($memoryState.ModuleCount)"
        Write-Host ""

        $statusColor = if ($memoryState.SpeedMismatch) { "Yellow" } else { "Green" }
        Write-Host "Speed Mismatch         : " -NoNewline
        Write-Host $(if ($memoryState.SpeedMismatch) { "Yes — modules running at different speeds" } else { "No" }) -ForegroundColor $statusColor

        Write-Host ""
        Write-Host "[ Module Details ]"
        Write-Host ""

        $moduleIndex = 0
        foreach ($module in $memoryState.Modules) {
            Write-Host "Module $($moduleIndex + 1)"
            Write-Host ("-" * 50)
            Write-Host "Slot            : $($module.DeviceLocator)"
            Write-Host "Size            : $($module.CapacityGB) GB"
            Write-Host "Speed           : $($module.SpeedMHz) MHz"
            Write-Host "Manufacturer    : $($module.Manufacturer)"
            Write-Host ""
            $moduleIndex++
        }

        Write-Host "[ Status ]"
        Write-Host $memoryState.Message -ForegroundColor $statusColor
    }

    Read-Host "Press ENTER"
}

Show-Disclaimer

# Initialise the shared CIM session now that elevation is confirmed.
# Must be called after Show-Disclaimer so UAC self-elevation has already
# completed and the process token is elevated before connecting.
New-ScriptCimSession

# Query battery health once at startup and cache it in $script:BatteryAlert.
# This runs silently in the background — the result is surfaced in Show-Header
# and the main menu so reactive users see the alert the moment they open the script.
$null = Get-BatteryAlertState

# Silently append a health snapshot to the persistent trend log.
# Duplicate suppression inside the function prevents multiple entries per day.
Write-BatteryTrendLog

# ── One-time welcome notice ───────────────────────────────────────────────
# Shown only on the very first launch (flag file written to %TEMP% afterward).
# Gives first-time users immediate context on what this tool does and why
# battery health matters — without interrupting repeat users at all.
$welcomeFlag = "$env:TEMP\LenovoUtil_Notice.flag"
if (-not (Test-Path $welcomeFlag)) {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  Welcome to Lenovo Device Health Check" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This tool shows you hardware and battery information"
    Write-Host "that Windows doesn't normally display — read directly"
    Write-Host "from your device."
    Write-Host ""
    Write-Host "Most importantly: check your battery health early." -ForegroundColor Yellow
    Write-Host "A worn battery can cause short runtimes, unexpected"
    Write-Host "shutdowns, and in severe cases, swelling."
    Write-Host ""
    Write-Host "→ If your laptop dies faster than it used to, start with options 1–3." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Press ENTER to continue..."
    Read-Host | Out-Null
    "seen" | Out-File $welcomeFlag -Encoding ascii
}


function Show-ErrorLog {
    Show-Header
    Write-Host "[ Issue Log ]"
    Write-Host ""

    if ($script:ErrorLog.Count -eq 0) {
        Write-Host "No issues recorded this session." -ForegroundColor Green
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    Write-Host "Issues recorded this session: $($script:ErrorLog.Count)" -ForegroundColor Red
    Write-Host ""

    $index = 0
    foreach ($err in $script:ErrorLog) {
        $index++
        Write-Host ("=" * 50) -ForegroundColor DarkGray
        Write-Host "Issue #$index" -ForegroundColor Red
        Write-Host ("=" * 50) -ForegroundColor DarkGray

        Write-Host "Time      : $($err.TimeStamp)"
        Write-Host "Component : $($err.Namespace) / $($err.Class)"

        Write-Host "Detail    : " -NoNewline
        Write-Host $err.Message -ForegroundColor Yellow

        if ($err.Context) {
            Write-Host "Context   : $($err.Context)"
        }

        if ($err.ScriptName -or $err.ScriptLineNumber) {
            Write-Host "Location  : line $($err.ScriptLineNumber)" -ForegroundColor DarkGray
        }

        if ($err.BatteryDetails) {
            Write-Host "Battery   :" -ForegroundColor DarkGray
            Write-Host "  Max Charge Now  : $($err.BatteryDetails.FullChargedCapacity)" -ForegroundColor DarkGray
            Write-Host "  Original Cap.   : $($err.BatteryDetails.DesignedCapacity)" -ForegroundColor DarkGray
        }

        if ($err.StackTrace) {
            Write-Host "Trace     :" -ForegroundColor DarkGray
            $err.StackTrace -split "`n" | ForEach-Object {
                Write-Host "  $($_.Trim())" -ForegroundColor DarkGray
            }
        }

        Write-Host ""
    }

    Read-Host "Press ENTER"
}


function Get-LenovoHealthScore {
    <#
    .SYNOPSIS
    Computes a single 0-100 health score for this Lenovo device using data
    sources available on consumer devices (no SIF or CDRT required).

    Component weights (sum to 100):
      Battery Health %  65 pts  -- primary wear indicator
      Cycle Count       30 pts  -- corroborates battery age (ACPI only)
      Memory Mismatch    5 pts  -- configuration defect

    Shock and thermal events are excluded as they require Commercial Vantage
    (CDRT), which is not deployed on consumer Lenovo devices.
    #>

    $score      = 0
    $components = [ordered]@{}
    $sources    = @()

    # ── Component 1: Battery Health % (65 pts) ──────────────────────────────
    $batteryPct   = $null
    $batteryNotes = "No battery data found"

    try {
        $acpiBats = @(Get-WmiObject -Namespace root\wmi `
            -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
        if ($acpiBats -and $acpiBats.Count -gt 0) {
            $healths = @()
            for ($i = 0; $i -lt $acpiBats.Count; $i++) {
                $fc  = $acpiBats[$i].FullChargedCapacity
                $dcV = Get-DesignCapacity -Index $i
                [int64]$fcN = 0; [int64]$dcN = 0
                if ($fc -and $dcV -and
                    [int64]::TryParse($fc.ToString(),  [ref]$fcN) -and
                    [int64]::TryParse($dcV.ToString(), [ref]$dcN) -and
                    $dcN -gt 0) {
                    $healths += [math]::Round(($fcN / $dcN) * 100, 1)
                }
            }
            if ($healths.Count -gt 0) {
                $batteryPct   = ($healths | Measure-Object -Minimum).Minimum
                $batteryNotes = "Worst battery: $batteryPct%"
                $sources     += "Battery capacity"
            }
        }
    } catch {}

    $batteryScore = if ($null -ne $batteryPct) {
        [math]::Round(($batteryPct / 100) * 65, 1)
    } else {
        $batteryNotes = "No battery data — assumed neutral"
        [math]::Round(0.70 * 65, 1)
    }
    $components["Battery Health"] = [PSCustomObject]@{
        Score    = $batteryScore
        MaxScore = 65
        Note     = $batteryNotes
    }

    # ── Component 2: Cycle Count (30 pts) ────────────────────────────────────
    # Source: Win32_Battery.BatteryStatus is available on consumer devices
    # but cycle count is not exposed without root\Lenovo. We leave this
    # component at neutral (70%) when unavailable rather than penalising.
    $cycleScore = [math]::Round(0.70 * 30, 1)
    $cycleNotes = "Charge cycle count not available on this device — assumed neutral"
    $components["Cycle Count"] = [PSCustomObject]@{
        Score    = $cycleScore
        MaxScore = 30
        Note     = $cycleNotes
    }

    # ── Component 3: Memory Mismatch (5 pts) ─────────────────────────────────
    $memScore = 0
    $memNotes = "Memory data unavailable"
    try {
        $memState = Get-MemoryState
        if ($memState.ModuleCount -gt 0) {
            $sources += "Memory modules"
            if ($memState.SpeedMismatch) {
                $memScore = 0
                $memNotes = "Speed mismatch detected ($($memState.ModuleCount) modules)"
            } else {
                $memScore = 5
                $memNotes = "$($memState.ModuleCount) module(s) matched"
            }
        } else {
            $memScore = 5
            $memNotes = "Single or soldered module — no penalty"
        }
    } catch {}
    $components["Memory"] = [PSCustomObject]@{
        Score    = $memScore
        MaxScore = 5
        Note     = $memNotes
    }

    # ── Total ─────────────────────────────────────────────────────────────────
    $total = [math]::Round($batteryScore + $cycleScore + $memScore)
    $total = [math]::Max(0, [math]::Min(100, $total))

    $grade      = if     ($total -ge 85) { "Excellent" }
                  elseif ($total -ge 70) { "Good"      }
                  elseif ($total -ge 55) { "Fair"      }
                  elseif ($total -ge 35) { "Poor"      }
                  else                   { "Critical"  }

    $gradeColor = if     ($total -ge 85) { "Green"   }
                  elseif ($total -ge 70) { "Cyan"    }
                  elseif ($total -ge 55) { "Yellow"  }
                  elseif ($total -ge 35) { "Magenta" }
                  else                   { "Red"     }

    return [PSCustomObject]@{
        Score       = $total
        Grade       = $grade
        GradeColor  = $gradeColor
        Components  = $components
        DataSources = $sources
    }
}

function Show-LenovoHealthScore {
    Show-Header
    Write-Host "[ Lenovo Device Health Score ]"
    Write-Host ""
    Write-Host "Calculating..." -ForegroundColor Cyan
    Write-Host ""

    $hs = Get-LenovoHealthScore

    # ── Score banner ──────────────────────────────────────────────────────────
    Write-Host ("=" * 50)
    Write-Host "  Health Score : " -NoNewline
    Write-Host "$($hs.Score) / 100" -ForegroundColor $hs.GradeColor -NoNewline
    Write-Host "   [$($hs.Grade)]" -ForegroundColor $hs.GradeColor
    Write-Host ("=" * 50)
    Write-Host ""

    # ── Component breakdown ───────────────────────────────────────────────────
    Write-Host "[ Component Breakdown ]"
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    foreach ($key in $hs.Components.Keys) {
        $c         = $hs.Components[$key]
        $pct       = if ($c.MaxScore -gt 0) { [math]::Round(($c.Score / $c.MaxScore) * 100) } else { 0 }
        $compColor = if     ($pct -ge 85) { "Green"   }
                     elseif ($pct -ge 70) { "Cyan"    }
                     elseif ($pct -ge 55) { "Yellow"  }
                     elseif ($pct -ge 35) { "Magenta" }
                     else                 { "Red"     }
        $label = "$key".PadRight(20)
        Write-Host "  $label : " -NoNewline
        Write-Host "$($c.Score) / $($c.MaxScore) pts" -ForegroundColor $compColor -NoNewline
        Write-Host "   $($c.Note)" -ForegroundColor DarkGray
    }
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    Write-Host ""

    Write-Host "Note: Charge cycle count is not available on consumer Lenovo devices." -ForegroundColor DarkGray
    Write-Host "      Battery health is the main indicator used in this score." -ForegroundColor DarkGray
    Write-Host ""

    Read-Host "Press ENTER"
}



function Get-BiosUpdateInfo {
    <#
    .SYNOPSIS
    Checks whether a newer BIOS version is available for this device by
    querying Lenovo's public per-MTM XML catalog.

    .DESCRIPTION
    Lenovo publishes a machine-readable update catalog at:
        https://download.lenovo.com/catalog/<MTM>_Win<10|11>.xml

    The MTM (Machine Type) is the first four characters of the model number
    from Win32_ComputerSystemProduct.Name (e.g. '20X8' from '20X8S06P00').

    The catalog lists all available packages for the device. This function
    finds the BIOS package entry, follows its location URL to retrieve the
    exact version string and release date, then compares against the
    currently installed BIOS version from Win32_BIOS.SMBIOSBIOSVersion.

    Returns a PSCustomObject with:
      InstalledVersion  - Version string from Win32_BIOS
      InstalledDate     - BIOS release date from Win32_BIOS
      LatestVersion     - Latest version from Lenovo catalog
      LatestDate        - Release date of the latest version
      IsUpToDate        - $true if installed >= latest
      UpdateAvailable   - $true if a newer version exists
      MTM               - Four-character machine type used for the lookup
      CatalogUrl        - The catalog URL that was queried
      Available         - $false if the check could not be completed
      UnavailableReason - Populated when Available is $false
    #>

    $result = [PSCustomObject]@{
        InstalledVersion  = "Unknown"
        InstalledDate     = "Unknown"
        LatestVersion     = "Unknown"
        LatestDate        = "Unknown"
        IsUpToDate        = $false
        UpdateAvailable   = $false
        MTM               = "Unknown"
        CatalogUrl        = ""
        Available         = $false
        UnavailableReason = ""
    }

    # ── Installed BIOS from cached system info ───────────────────────────────
    $si = Get-SystemInfo
    $result.InstalledVersion = $si.BIOSVersion
    $result.InstalledDate    = $si.BIOSDate

    # ── Extract MTM from Win32_ComputerSystemProduct.Name ────────────────────
    # Name returns the full model number e.g. '20X8S06P00'.
    # The catalog uses only the first four characters: '20X8'.
    try {
        $csp = Get-CimInstance -CimSession $script:CimSession `
            -ClassName Win32_ComputerSystemProduct -ErrorAction Stop
        if ($csp -and $csp.Name -and $csp.Name.Length -ge 4) {
            $result.MTM = $csp.Name.Substring(0, 4).ToUpper()
        }
        else {
            $result.UnavailableReason = "Could not read model number from Win32_ComputerSystemProduct."
            return $result
        }
    }
    catch {
        $result.UnavailableReason = "Win32_ComputerSystemProduct query failed: $($_.Exception.Message)"
        return $result
    }

    # ── Detect Windows version for catalog URL suffix ────────────────────────
    $winSuffix = "Win10"
    try {
        $os = Get-CimInstance -CimSession $script:CimSession `
            -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os -and $os.Version -match "^10\.0\.2") {
            # Build 20000+ = Windows 11
            $winSuffix = "Win11"
        }
    }
    catch {}

    # ── Fetch catalog XML ─────────────────────────────────────────────────────
    $catalogUrl = "https://download.lenovo.com/catalog/$($result.MTM)_$winSuffix.xml"
    $result.CatalogUrl = $catalogUrl

    try {
        $response = Invoke-WebRequest -Uri $catalogUrl -UseBasicParsing `
            -TimeoutSec 15 -ErrorAction Stop
        # Decode raw bytes as UTF-8 explicitly -- Invoke-WebRequest may
        # misdetect encoding and return the UTF-8 BOM as garbage characters.
        # [System.Text.Encoding]::UTF8.GetString strips the BOM correctly.
        $catalogContent = [System.Text.Encoding]::UTF8.GetString($response.RawContent[($response.RawContent.IndexOf([byte]0x3C))..($response.RawContent.Length - 1)])
        [xml]$catalog = $catalogContent
    }
    catch {
        $result.UnavailableReason = "Failed to fetch catalog ($catalogUrl): $($_.Exception.Message)"
        return $result
    }

    # ── Find BIOS package in catalog ──────────────────────────────────────────
    # The category value in Lenovo's catalog is 'BIOS UEFI', not 'BIOS'.
    # We match on 'BIOS' as a substring to handle any future variation.
    $biosPackage = $null
    foreach ($pkg in $catalog.packages.package) {
        if ($pkg.Category -match 'BIOS') {
            $biosPackage = $pkg
            break
        }
    }

    if (-not $biosPackage) {
        $result.UnavailableReason = "No BIOS package found in catalog for MTM $($result.MTM)."
        return $result
    }

    # ── Fetch BIOS package descriptor for version and date ───────────────────
    $pkgUrl = $biosPackage.location
    if (-not $pkgUrl) {
        $result.UnavailableReason = "BIOS package entry has no location URL."
        return $result
    }

    try {
        $pkgResponse = Invoke-WebRequest -Uri $pkgUrl -UseBasicParsing `
            -TimeoutSec 15 -ErrorAction Stop
        # Same UTF-8 BOM workaround as the catalog fetch above.
        $pkgContent = [System.Text.Encoding]::UTF8.GetString($pkgResponse.RawContent[($pkgResponse.RawContent.IndexOf([byte]0x3C))..($pkgResponse.RawContent.Length - 1)])
        [xml]$pkgXml = $pkgContent
    }
    catch {
        $result.UnavailableReason = "Failed to fetch BIOS package descriptor ($pkgUrl): $($_.Exception.Message)"
        return $result
    }

    $latestVersion = $pkgXml.Package.version
    $latestDate    = $pkgXml.Package.ReleaseDate

    if (-not $latestVersion) {
        $result.UnavailableReason = "BIOS package descriptor did not contain a version string."
        return $result
    }

    $result.LatestVersion = $latestVersion.Trim()
    $result.LatestDate    = if ($latestDate) { $latestDate.Trim() } else { "Unknown" }
    $result.Available     = $true

    # ── Compare versions ──────────────────────────────────────────────────────
    # Win32_BIOS.SMBIOSBIOSVersion may return a string like 'H1CN50WW' or
    # 'H1CN50WW (1.50)'. The catalog returns only the numeric part e.g. '1.50'.
    # Plain string comparison always fails when the formats differ.
    #
    # Strategy: extract the numeric version from both sides.
    #   Installed: pull the decimal number from inside parentheses if present,
    #              otherwise fall back to any decimal number in the string.
    #   Catalog:   already a plain decimal string, just trim it.
    # Compare as [version] objects so 1.10 > 1.9 correctly.
    # If either side cannot be parsed as a version, fall back to string compare.

    $installedRaw = $result.InstalledVersion.Trim()
    $latestRaw    = $result.LatestVersion.Trim()

    $installedNumeric = $null
    if ($installedRaw -match '\(([\d\.]+)') {
        $installedNumeric = $matches[1].Trim()
    }
    elseif ($installedRaw -match '([\d]+\.[\d]+)') {
        $installedNumeric = $matches[1].Trim()
    }

    $latestNumeric = $null
    if ($latestRaw -match '([\d]+\.[\d]+)') {
        $latestNumeric = $matches[1].Trim()
    }

    $compared = $false
    if ($installedNumeric -and $latestNumeric) {
        try {
            $vInstalled = [version]$installedNumeric
            $vLatest    = [version]$latestNumeric
            if ($vInstalled -ge $vLatest) {
                $result.IsUpToDate      = $true
                $result.UpdateAvailable = $false
            }
            else {
                $result.IsUpToDate      = $false
                $result.UpdateAvailable = $true
            }
            $compared = $true
        }
        catch {}
    }

    # Fallback: plain string compare (catches identical strings)
    if (-not $compared) {
        if ($installedRaw.ToUpper() -eq $latestRaw.ToUpper()) {
            $result.IsUpToDate      = $true
            $result.UpdateAvailable = $false
        }
        else {
            $result.IsUpToDate      = $false
            $result.UpdateAvailable = $true
        }
    }

    return $result
}

function Show-BiosUpdateCheck {
    Show-Header
    Write-Host "[ BIOS Update Check ]"
    Write-Host ""
    Write-Host "Querying Lenovo update catalog..." -ForegroundColor Cyan
    Write-Host ""

    $bios = Get-BiosUpdateInfo

    Write-Host "[ Installed ]"
    Write-Host "  Version      : $($bios.InstalledVersion)"
    Write-Host "  Release Date : $($bios.InstalledDate)"
    Write-Host ""

    if (-not $bios.Available) {
        Write-Host "[ Latest ]"
        Write-Host "  Could not check for updates online." -ForegroundColor Yellow
        Write-Host "  Reason: $($bios.UnavailableReason)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Make sure you are connected to the internet and try again." -ForegroundColor DarkGray
        Write-Host "  Or check manually: support.lenovo.com → Drivers & Software" -ForegroundColor DarkGray
    }
    else {
        Write-Host "[ Latest ]"
        Write-Host "  Version      : $($bios.LatestVersion)"
        Write-Host "  Release Date : $($bios.LatestDate)"
        Write-Host ""

        Write-Host "[ Status ]"
        if ($bios.IsUpToDate) {
            Write-Host "  Your BIOS is up to date." -ForegroundColor Green
        }
        else {
            Write-Host "  An update is available." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Installed : $($bios.InstalledVersion)" -ForegroundColor DarkGray
            Write-Host "  Available : $($bios.LatestVersion)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Download from: support.lenovo.com → Drivers & Software" -ForegroundColor Cyan
            Write-Host "  Search for your model and filter by BIOS/UEFI." -ForegroundColor DarkGray
        }
    }

    Write-Host ""
    Write-Host "Note: If the version format is unusual, compare the dates manually." -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press ENTER"
}

function Show-StorageInfo {
    Show-Header
    Write-Host "[ Storage ]"
    Write-Host ""

    $sto = Get-StorageInfo

    if (-not $sto.Available) {
        Write-Host "Could not read drive information on this device." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    Write-Host "Type             : $($sto.Type)"
    Write-Host "Size             : $($sto.SizeGB) GB"
    Write-Host "Drive Model      : $($sto.Model)"
    Write-Host ""

    Write-Host "[ Tips ]" -ForegroundColor Cyan
    if ($sto.Type -match "NVMe") {
        Write-Host "  Your device has a fast NVMe SSD." -ForegroundColor DarkGray
        Write-Host "  Avoid filling it above 90% of capacity for best performance." -ForegroundColor DarkGray
    } elseif ($sto.Type -match "SSD") {
        Write-Host "  Your device has an SSD (Solid State Drive)." -ForegroundColor DarkGray
        Write-Host "  SSDs are fast and durable. Avoid filling them above 90% capacity." -ForegroundColor DarkGray
    } elseif ($sto.Type -match "HDD") {
        Write-Host "  Your device has a traditional hard drive (HDD)." -ForegroundColor DarkGray
        Write-Host "  HDDs are slower and more sensitive to bumps and drops." -ForegroundColor DarkGray
        Write-Host "  Consider backing up your data regularly." -ForegroundColor DarkGray
    } else {
        Write-Host "  Drive type could not be identified." -ForegroundColor DarkGray
        Write-Host "  Back up your data regularly regardless of drive type." -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  Always back up important files before upgrading or replacing storage." -ForegroundColor DarkGray
    Write-Host ""

    Read-Host "Press ENTER"
}

function Show-BatteryTips {
    Show-Header
    Write-Host "[ Battery Care Tips ]"
    Write-Host ""

    # Personalise intro based on current battery health
    if ($script:BatteryAlert -and $script:BatteryAlert.Available) {
        $worst = $script:BatteryAlert.WorstSeverity
        if ($worst -ge 3) {
            Write-Host "  Your battery is in poor condition. These tips can help slow further" -ForegroundColor Red
            Write-Host "  wear, but at this stage replacement is the best long-term option." -ForegroundColor Red
        } elseif ($worst -ge 1) {
            Write-Host "  Your battery is showing some wear. These habits can help" -ForegroundColor Yellow
            Write-Host "  slow further degradation and extend its remaining life." -ForegroundColor Yellow
        } else {
            Write-Host "  Your battery is in good shape. These habits will help" -ForegroundColor Green
            Write-Host "  keep it that way for as long as possible." -ForegroundColor Green
        }
        Write-Host ""
    }

    Write-Host "[ Charging ]" -ForegroundColor Cyan
    Write-Host "  • Keep your charge between 20% and 80% for day-to-day use." -ForegroundColor Gray
    Write-Host "    Staying at 100% for long periods wears the battery faster." -ForegroundColor DarkGray
    Write-Host "  • If mostly plugged in, enable Conservation Mode in Lenovo Vantage" -ForegroundColor Gray
    Write-Host "    (limits charging to ~60% to protect long-term battery health)." -ForegroundColor DarkGray
    Write-Host ""

    Write-Host "[ Heat ]" -ForegroundColor Cyan
    Write-Host "  • Heat is the biggest enemy of battery life." -ForegroundColor Gray
    Write-Host "    Avoid leaving your laptop in a hot car or direct sunlight." -ForegroundColor DarkGray
    Write-Host "  • Keep vents clear — use on hard, flat surfaces." -ForegroundColor DarkGray
    Write-Host ""

    Write-Host "[ Long-term Storage ]" -ForegroundColor Cyan
    Write-Host "  • If storing unused for weeks, keep it at around 50% charge." -ForegroundColor Gray
    Write-Host "    Storing fully charged or fully empty both accelerate wear." -ForegroundColor DarkGray
    Write-Host ""

    Write-Host "[ Replacement ]" -ForegroundColor Cyan
    Write-Host "  • Lenovo recommends considering replacement below 80% health." -ForegroundColor Gray
    Write-Host "  • Always use a genuine Lenovo battery for safety and compatibility." -ForegroundColor DarkGray
    Write-Host "    Find yours at: support.lenovo.com  →  Parts" -ForegroundColor DarkGray
    Write-Host ""

    Read-Host "Press ENTER"
}

do {
    Show-Header

    # ── Guided-mode auto-prompt (#3) ─────────────────────────────────────
    # After the user has browsed the menu 2 times without selecting a battery
    # option, and any battery is critically degraded (Severity >= 3), gently
    # surface the Comprehensive Battery Analysis once. Fires only once per
    # session so it never becomes intrusive to power users.
    $script:MenuLoopCount++
    if (-not $script:GuidedAnalysisShown -and
        $script:MenuLoopCount -ge 2 -and
        $script:BatteryAlert -and
        $script:BatteryAlert.Available -and
        $script:BatteryAlert.WorstSeverity -ge 3) {

        $script:GuidedAnalysisShown = $true
        $guideColor = switch ($script:BatteryAlert.WorstSeverity) {
            1 { "Cyan" }; 2 { "Yellow" }; 3 { "Magenta" }; 4 { "Red" }
            default { "Yellow" }
        }
        Write-Host "-----------------------------------------------" -ForegroundColor $guideColor
        Write-Host "  Battery concern detected. Would you like to run" -ForegroundColor $guideColor
        Write-Host "  Option 3 (Comprehensive Battery Analysis) now?" -ForegroundColor $guideColor
        Write-Host "  Press Y to run it, or any other key to skip." -ForegroundColor DarkGray
        Write-Host "-----------------------------------------------" -ForegroundColor $guideColor
        $guided = Read-Host "Run analysis now? [Y/any key to skip]"
        if ($guided.Trim().ToUpper() -eq "Y") {
            ComprehensiveBatteryAnalysis
            continue
        }
        Show-Header
    }

    # Build alert suffix for battery-related menu items
    # Shown next to battery options when any battery is degraded
    $alertSuffix = ""
    $alertColor  = "White"
    if ($script:BatteryAlert -and $script:BatteryAlert.Available -and $script:BatteryAlert.WorstSeverity -ge 1) {
        $alertSuffix = "  [!] $($script:BatteryAlert.SummaryLine)"
        $alertColor  = switch ($script:BatteryAlert.WorstSeverity) {
            1 { "Cyan" }; 2 { "Yellow" }; 3 { "Magenta" }; 4 { "Red" }
            default { "White" }
        }
    }

    # ── Battery status banner above the menu ────────────────────────────
    # Shows a colored summary line when any battery has Severity >= 1,
    # so the alert is visible before the user even reads the menu.
    if ($alertSuffix) {
        Write-Host "Battery Status Alert: $($script:BatteryAlert.SummaryLine)" -ForegroundColor $alertColor
        Write-Host ""
    }

    Write-Host "1. Battery Details"
    Write-Host -NoNewline "2. Battery Information"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host -NoNewline "3. Comprehensive Battery Analysis"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "4. Memory Info"
    Write-Host "5. Save Report"
    Write-Host "6. What This Tool Can See"
    Write-Host -NoNewline "7. Battery Age Estimation"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "8. Issue Log"
    Write-Host "9. Device Health Score"
    Write-Host "10. BIOS Update Check"
    Write-Host "11. About"
    Write-Host -NoNewline "12. Battery Health Trend"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "13. Storage Info"
    Write-Host -NoNewline "14. Battery Care Tips"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "15. Exit"
    Write-Host ""

    $choice = Read-Host "Select option"

    switch ($choice) {
        "1"  { BatteryStaticData }
        "2"  { FullCharge }
        "3"  { ComprehensiveBatteryAnalysis }
        "4"  { MemoryInfo }
        "5"  { ExportReport }
        "6"  { DiagnosticInfo }
        "7"  { BatteryAgeEstimation }
        "8"  { Show-ErrorLog }
        "9"  { Show-LenovoHealthScore }
        "10" { Show-BiosUpdateCheck }
        "11" { Show-About }
        "12" { Show-BatteryHealthTrend }
        "13" { Show-StorageInfo }
        "14" { Show-BatteryTips }
    }

} while ($choice -ne "15")

# Tear down the shared CIM session cleanly before exiting.
if ($script:CimSession) {
    Remove-CimSession -CimSession $script:CimSession -ErrorAction SilentlyContinue
    $script:CimSession = $null
}

Clear-Host
Write-Host "Exiting..."