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
        Write-Host "  ACCESS DENIED" -ForegroundColor Red
        Write-Host ""

        if ($_.Exception.Message -match "canceled by the user|was canceled") {
            # User clicked "No" on the UAC prompt
            Write-Host "  UAC elevation was cancelled." -ForegroundColor Yellow
        }
        else {
            # Any other failure (policy block, missing path, etc.)
            Write-Host "  Could not elevate automatically:" -ForegroundColor Yellow
            Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Right-click PowerShell and select" -ForegroundColor Gray
            Write-Host "  'Run as Administrator', then try again." -ForegroundColor Gray
        }

        Write-Host ""
        Write-Host "  Why elevation is required:" -ForegroundColor Cyan
        Write-Host "  This script reads data directly from your ThinkPad's BIOS and" -ForegroundColor Gray
        Write-Host "  Embedded Controller (EC) via the root\wmi and root\Lenovo WMI" -ForegroundColor Gray
        Write-Host "  namespaces. Windows restricts access to these interfaces to" -ForegroundColor Gray
        Write-Host "  Administrator processes only, to prevent unprivileged software" -ForegroundColor Gray
        Write-Host "  from reading or modifying firmware-level hardware state." -ForegroundColor Gray
        Write-Host ""
        Read-Host "  Press ENTER to exit"
        Exit
    }
}

$script:ErrorLog    = @()
$script:SystemInfo  = $null
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
      - Forward-compatible with remote ThinkPad management (change
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
    Returns the list of genuine battery manufacturer name tokens used for
    non-genuine detection in Get-BatteryClassification.

    .DESCRIPTION
    Tries to load the list from manufacturers.json in the same directory as
    this script. If the file is missing, unreadable, or contains no entries,
    falls back silently to the built-in hardcoded list so the script never
    breaks due to a missing or malformed JSON file.

    To extend detection without editing this script, add entries to the
    genuineManufacturers array in manufacturers.json:
      { "name": "NEWMFR", "fullName": "New Manufacturer Inc.", "region": "XX", "notes": "" }

    The "name" field is matched case-insensitively as a substring against
    BatteryStaticData.ManufactureName. Only the "name" field is required.
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

    # Detect non-genuine batteries
    # Load manufacturer list from manufacturers.json (falls back to built-in list if missing)
    $GenuineManufacturers = Get-GenuineManufacturerList
    $IsGenuine = $false

    if ($Manufacturer -and $Manufacturer -ne "Unknown" -and $Manufacturer -ne "N/A") {
        $normalizedManufacturer = $Manufacturer.Trim().ToUpper()
        foreach ($Genuine in $GenuineManufacturers) {
            if ($normalizedManufacturer -match [regex]::Escape($Genuine.ToUpper())) {
                $IsGenuine = $true
                break
            }
        }
    }

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
    if (-not $IsGenuine) {
        $Message += " Warning: The battery in use is not genuine Lenovo-made or authorized. Lenovo has no responsibility for the performance or safety of unauthorized batteries, and provides no warranties for failure or damage arising from their use."
    }

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

    # Environment detection: domain-joined vs standalone workgroup
    $environment = "Standalone / Workgroup"
    $domainName  = ""
    try {
        $cs2 = Get-CimInstance -CimSession $script:CimSession -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs2 -and $cs2.PartOfDomain) {
            $environment = "Domain-Joined (Enterprise)"
            $domainName  = if ($cs2.Domain) { $cs2.Domain } else { "" }
        }
    } catch {}

    $script:SystemInfo = [PSCustomObject]@{
        Model       = $model
        BIOSVersion = $biosVersion
        BIOSDate    = $biosDate
        Environment = $environment
        DomainName  = $domainName
    }

    return $script:SystemInfo
}


function Get-LenovoDeviceFamily {
    <#
    .SYNOPSIS
    Returns "ThinkPad" if this is a ThinkPad, otherwise "Other" or "Unknown".

    .DESCRIPTION
    The root\Lenovo WMI namespace is a ThinkPad-exclusive firmware feature.
    It is not present on any other Lenovo product line (IdeaPad, Yoga, Legion,
    IdeaCentre, ThinkBook, ThinkCentre, ThinkStation, etc.) by design —
    this is a firmware architecture decision, not a driver or configuration issue.

    Rather than enumerating every possible non-ThinkPad family name, the logic
    is intentionally binary:
      - "ThinkPad" : model string contains "ThinkPad" — namespace should be present
      - "Other"    : any other Lenovo device — namespace is absent by design
      - "Unknown"  : model string could not be read

    Uses $script:SystemInfo (already cached by Get-SystemInfo) — no extra WMI query.
    #>
    $si = Get-SystemInfo
    $model = $si.Model

    if (-not $model -or $model -eq "Unknown") { return "Unknown" }

    if ($model.ToUpper() -match "THINKPAD") { return "ThinkPad" }

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
            # ThinkPad — namespace should be present, absence is a driver problem
            Write-Host "root\Lenovo namespace not available." -ForegroundColor Red
            Write-Host ""
            if ($sifInstalled) {
                Write-Host "Lenovo System Interface Foundation (SIF) appears to be installed," -ForegroundColor Yellow
                Write-Host "but the root\Lenovo WMI namespace was not found." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Try:" -ForegroundColor Yellow
                Write-Host "  1. Restart the WMI service:  Restart-Service winmgmt"
                Write-Host "  2. Re-register the provider: mofcomp.exe"
                Write-Host "  3. Reinstall SIF from support.lenovo.com -> Drivers & Software"
            } else {
                Write-Host "Lenovo System Interface Foundation (SIF) is not installed." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "SIF registers the root\Lenovo WMI namespace used for $Feature." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Install SIF from Lenovo Support:" -ForegroundColor Cyan
                Write-Host "  support.lenovo.com  ->  Drivers & Software"
                Write-Host "  Search: 'System Interface Foundation'"
            }
        }

        "Other" {
            # Any non-ThinkPad Lenovo device — absent by firmware design
            Write-Host "root\Lenovo namespace is not available on this device." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Device: $($si.Model)" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "The Lenovo EC WMI interface used for $Feature is a" -ForegroundColor Yellow
            Write-Host "ThinkPad-exclusive firmware feature. It is not present on" -ForegroundColor Yellow
            Write-Host "other Lenovo devices by design — this is not a driver or" -ForegroundColor Yellow
            Write-Host "configuration issue and cannot be resolved by installing SIF." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "For battery and system information on this device," -ForegroundColor Cyan
            Write-Host "use the Lenovo Vantage app (available in the Microsoft Store)." -ForegroundColor Cyan
        }

        default {
            # Model string unreadable — generic fallback
            Write-Host "root\Lenovo namespace is not available on this system." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "This interface is only supported on Lenovo ThinkPad systems." -ForegroundColor Yellow
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

    PowerBridge ThinkPads (T440, T450, T460, W540 series and similar) have two
    batteries — an internal fixed cell and a hot-swappable external bay battery.
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
    $si = Get-SystemInfo
    Write-Host "==============================================="
    Write-Host "  THINKPAD UTILITY MANAGEMENT DASHBOARD V.1.2"
    Write-Host "==============================================="

    # ── [ SYSTEM INFO ] ──────────────────────────────────────────────────
    Write-Host "[ SYSTEM INFO ]" -ForegroundColor Cyan
    Write-Host "  Device      : $env:COMPUTERNAME"
    Write-Host "  Model       : $($si.Model)"
    Write-Host "  BIOS        : $($si.BIOSVersion)  ($($si.BIOSDate))"
    # Environment: show domain name when joined, plain label otherwise
    $envDisplay = if ($si.DomainName) { "$($si.Environment)  ($($si.DomainName))" } else { $si.Environment }
    $envColor   = if ($si.Environment -match "Domain") { "Cyan" } else { "DarkGray" }
    Write-Host "  Environment : " -NoNewline
    Write-Host $envDisplay -ForegroundColor $envColor
    Write-Host ""

    # ── [ BATTERY STATUS ] ───────────────────────────────────────────────
    Write-Host "[ BATTERY STATUS ]" -ForegroundColor Cyan
    if ($script:BatteryAlert -and $script:BatteryAlert.Available) {
        $worst  = $script:BatteryAlert.WorstSeverity
        $emoji  = @{0="  OK"; 1="⚠"; 2="⚠"; 3="❗"; 4="🔥"}[$worst]
        $phrase = @{0="Healthy"; 1="Below 80%"; 2="Poor"; 3="Critical"; 4="FAILURE"}[$worst]
        $color  = @{0="Green"; 1="Cyan"; 2="Yellow"; 3="Magenta"; 4="Red"}[$worst]
        Write-Host "  Health      : " -NoNewline
        Write-Host "$emoji $phrase" -ForegroundColor $color -NoNewline
        if ($worst -ge 1) {
            Write-Host "  → see option 5" -ForegroundColor DarkGray
        } else {
            Write-Host ""
        }
    } else {
        Write-Host "  Health      : " -NoNewline
        Write-Host "Not detected / limited info" -ForegroundColor DarkGray
    }

    if ($script:ErrorLog.Count -gt 0) {
        Write-Host "  Issues      : " -NoNewline
        Write-Host "$($script:ErrorLog.Count) error(s) logged this session  [Option 13 to view]" -ForegroundColor Red
    }
    Write-Host ""
    # Displayed whenever any battery has Severity >= 1 (health below 80%).
    # Shows each affected battery individually on its own line with its
    # classification and percentage, so PowerBridge dual-battery models
    # report both cells clearly.
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
                Write-Host "     $bLabel`: $($b.Classification) — $($b.HealthPercent)%  [Select Option 5 for full analysis]" -ForegroundColor $bannerColor
            }
        }
        Write-Host "-----------------------------------------------" -ForegroundColor $bannerColor
    }

    Write-Host ""
}

function Show-Disclaimer {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Yellow
    Write-Host "DISCLAIMER" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This diagnostic script uses publicly available Windows and ACPI interfaces to read system information. Some optional data sources referenced by the script—such as community forums, blogs, Discord servers, or third‑party utilities—are not Lenovo‑official resources."
    Write-Host ""
    Write-Host "Lenovo is not responsible for moderating, maintaining, or validating external tools like fan Discord servers, blogs, forums, social media, and other similar services."
    Write-Host ""
    Write-Host "This project utilizes publicly documented Lenovo WMI interfaces available on supported ThinkPad systems."
    Write-Host ""
    Write-Host "This script is provided for educational and diagnostic purposes only. Unauthorized modifications or firmware may adversely affect warranty coverage."
    Write-Host "This project is not affiliated with or endorsed by Lenovo."
    Write-Host ""
    Write-Host "ThinkPad is a trademark of Lenovo."
    Write-Host ""
    Write-Host "You must choose to Accept to continue. You may Exit without accepting if you do not agree."
    Write-Host ""
    do {
        $response = Read-Host "Type 'A' to Accept and continue, or 'E' to Exit without accepting"
        if ($null -ne $response) {
            switch ($response.Trim().ToUpper()) {
                "A" { return }          # proceed with script
                "E" { Clear-Host; Write-Host "Exiting..."; Exit } # exit immediately
                default { Write-Host "Invalid selection. Please enter 'A' or 'E'." -ForegroundColor Yellow }
            }
        }
        else {
            Write-Host "Invalid selection. Please enter 'A' or 'E'." -ForegroundColor Yellow
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
    Write-Host "[ Battery Static Data - root\wmi ]"
    Write-Host ""

    try {
        if (-not (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryStaticData")) {
            Write-Host "BatteryStaticData class not available on this system." -ForegroundColor Yellow
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

                    Write-Host "Battery #: $deviceID"
                    Write-Host ("-" * 50)
                    Write-Host "Manufacturer         : $manufacturer"
                    Write-Host "Serial Number        : $serial"
                    Write-Host "Designed Capacity    : $designCapacity"
                    
                    Write-Host ""
                    $index++
                }
                catch {
                    Log-BatteryError "root\wmi" "BatteryStaticData" $_ $index $($bat.DeviceName) $null
                    Write-Host "Battery processing failed." -ForegroundColor Red
                    Write-Host ""
                    $index++
                }
            }
        }
    }
    catch {
        Log-Error "root\wmi" "BatteryStaticData" $_
        Write-Host "Failed to query battery information." -ForegroundColor Red
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

function LenovoEC {
    Show-Header
    Write-Host "[ Lenovo EC - Lenovo_Odometer ]"
    Write-Host ""

    if (-not (Test-LenovoNamespace)) {
        Get-LenovoNamespaceUnavailableMessage -Feature "battery cycle count and EC data"
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    try {
        $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
        $cycles = ([string]$odo.Battery_cycles -replace '\D','')

        Write-Host "Cycle Count : $cycles"
        Write-Host "Shock Events: $($odo.Shock_events)"
        Write-Host "Thermal     : $($odo.Thermal_events)"
    }
    catch {
        Log-Error "root\Lenovo" "Lenovo_Odometer" $_
        Write-Host "Lenovo_Odometer not accessible." -ForegroundColor Red
    }

    # -- Battery Temperatures (Lenovo_Battery) -----------------------------------
    # Lenovo_Battery exposes a Temperature property per battery pack.
    # The raw value varies by firmware version and locale:
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
    Write-Host ""
    Write-Host "[ Battery Temperatures - Lenovo_Battery ]"
    if (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery") {
        try {
            $lbBatteries = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop)
            if ($lbBatteries -and $lbBatteries.Count -gt 0) {
                foreach ($lb in $lbBatteries) {
                    $batteryID = Get-SafeWmiProperty -Object $lb -PropertyName "BatteryID"
                    $tempRaw   = Get-SafeWmiProperty -Object $lb -PropertyName "Temperature"
                    $label     = if ($batteryID -and $batteryID -ne "Unavailable") { $batteryID } else { "Battery" }

                    Write-Host "  $label`: " -NoNewline

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
                        $tempLabel = if     ($tempNum -gt 60) { "  [!] CRITICAL - Exceeds safe limits" }
                                     elseif ($tempNum -ge 51) { "  [!] ELEVATED - Thermal stress range (51-60 C)" }
                                     elseif ($tempNum -ge 41) { "  [!] WARNING - High temperature (41-50 C)" }
                                     else                     { "  Normal" }
                        Write-Host "$tempNum C" -ForegroundColor $tempColor -NoNewline
                        Write-Host $tempLabel -ForegroundColor $tempColor
                    }
                    else {
                        Write-Host $tempRaw
                    }
                }
            }
            else {
                Write-Host "  No batteries returned by Lenovo_Battery." -ForegroundColor Yellow
            }
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Battery" $_
            Write-Host "  Lenovo_Battery not accessible." -ForegroundColor Red
        }
    }
    else {
        Write-Host "  Lenovo_Battery class not available." -ForegroundColor Yellow
    }

    Read-Host "Press ENTER"
}

function WarrantyInfo {
    Show-Header
    Write-Host "[ Warranty - Lenovo_WarrantyInformation ]"
    Write-Host ""

    if (-not (Test-LenovoNamespace)) {
        Get-LenovoNamespaceUnavailableMessage -Feature "warranty information"
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    try {
        $w = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_WarrantyInformation -ErrorAction Stop

        Write-Host "Serial : $($w.SerialNumber)"
        Write-Host "Start  : $($w.StartDate)"
        Write-Host "End    : $($w.EndDate)"
        Write-Host "Last Sync : $($w.LastUpdateTime)"
    }
    catch {
        Log-Error "root\Lenovo" "Lenovo_WarrantyInformation" $_
        Write-Host "Warranty data unavailable." -ForegroundColor Red
    }

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
                Write-Host "Source: Lenovo_Battery (root\Lenovo)" -ForegroundColor DarkGray
                Write-Host ""

                $index = 0
                foreach ($lb in $lbBatteries) {
                    try {
                        Write-Host "Battery #: $($lb.BatteryID)"
                        Write-Host ("=" * 50)

                        # ── Identity ─────────────────────────────────────────
                        Write-Host "[ Identity ]"
                        Write-Host "  Manufacturer         : $(Get-SafeWmiProperty -Object $lb -PropertyName 'Manufacturer')"
                        Write-Host "  FRU Part Number      : $(Get-SafeWmiProperty -Object $lb -PropertyName 'FRUPartNumber')"
                        Write-Host "  Barcode              : $(Get-SafeWmiProperty -Object $lb -PropertyName 'BarCode')"
                        Write-Host "  Device Chemistry     : $(Get-SafeWmiProperty -Object $lb -PropertyName 'DeviceChemistry')"
                        Write-Host "  Firmware Version     : $(Get-SafeWmiProperty -Object $lb -PropertyName 'FirmwareVersion')"
                        Write-Host "  Manufacture Date     : $(Get-SafeWmiProperty -Object $lb -PropertyName 'ManufactureDate')"
                        Write-Host "  First Use Date       : $(Get-SafeWmiProperty -Object $lb -PropertyName 'FirstUseDate')"
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
                        Write-Host "  Cycle Count          : $lbCycles"
                        Write-Host "  Design Capacity      : $lbDesign"
                        Write-Host "  Full Charge Capacity : $lbFull"
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
                        Write-Host "  Remaining Percentage : $(Get-SafeWmiProperty -Object $lb -PropertyName 'RemainingPercentage')"
                        Write-Host "  Remaining Time       : $(Get-SafeWmiProperty -Object $lb -PropertyName 'RemainingTime')"
                        Write-Host "  Charge Completion    : $(Get-SafeWmiProperty -Object $lb -PropertyName 'ChargeCompletionTime')"
                        Write-Host "  Adapter              : $(Get-SafeWmiProperty -Object $lb -PropertyName 'Adapter')"
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
                            $tempLabel = if     ($tempNum -gt 60) { "  [!] CRITICAL - Exceeds safe limits" }
                                         elseif ($tempNum -ge 51) { "  [!] ELEVATED - Thermal stress range (51-60 C)" }
                                         elseif ($tempNum -ge 41) { "  [!] WARNING - High temperature (41-50 C)" }
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

                        Write-Host "[ Classification ]"
                        Write-Host "  Classification       : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "  Status               : " -NoNewline
                        Write-Host $classification.Message -ForegroundColor $severityColor

                        if (-not $classification.IsGenuine) {
                            Write-Host "  WARNING: Non-genuine battery detected" -ForegroundColor Red
                        }
                        if ($classification.FailureFlag) {
                            Write-Host "  ALERT: Battery replacement needed immediately!" -ForegroundColor Red
                        }

                        Write-Host ""
                        $index++
                    }
                    catch {
                        Log-BatteryError "root\Lenovo" "Lenovo_Battery" $_ $index $($lb.BatteryID) $null
                        Write-Host "Battery #: $index - FAILED" -ForegroundColor Red
                        Write-Host ""
                        $index++
                    }
                }
            }
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Battery" $_
            Write-Host "Lenovo_Battery query failed — falling back to ACPI source." -ForegroundColor Yellow
            Write-Host ""
        }
    }

    # ── ACPI fallback ────────────────────────────────────────────────────
    # Used when Lenovo_Battery is unavailable: SIF not installed, older
    # firmware, or the Lenovo_Battery query above failed.
    if (-not $usedLenovoBattery) {
        try {
            if (-not (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryFullChargedCapacity")) {
                Write-Host "Neither Lenovo_Battery nor BatteryFullChargedCapacity is available." -ForegroundColor Yellow
                Write-Host "Install Lenovo SIF drivers for full battery data." -ForegroundColor Yellow
                Read-Host "Press ENTER"
                return
            }

            $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

            if (-not $batteries -or $batteries.Count -eq 0) {
                Write-Host "No battery capacity data found." -ForegroundColor Yellow
            }
            else {
                Write-Host "Source: BatteryFullChargedCapacity (root\wmi) — ACPI fallback" -ForegroundColor DarkGray
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

                        Write-Host "Battery #: $index"
                        Write-Host ("-" * 50)
                        Write-Host "Instance Name        : $($bat.InstanceName)"
                        Write-Host "Manufacturer         : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })"
                        Write-Host "Full Charge Capacity : $FullCharge mWh"
                        Write-Host "Design Capacity      : $DesignCap mWh"
                        Write-Host "Health %             : $Health"

                        $classification = Get-BatteryClassification `
                            -DesignCapacity     $DesignCapValue `
                            -FullChargeCapacity $FullCharge `
                            -CycleCount         $null `
                            -Manufacturer       $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })

                        $severityColor = switch ($classification.Severity) {
                            0 { "Green" }; 1 { "Cyan" }; 2 { "Yellow" }; 3 { "Magenta" }; 4 { "Red" }
                            default { "White" }
                        }

                        Write-Host "Classification       : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "Status               : " -NoNewline
                        Write-Host $classification.Message -ForegroundColor $severityColor

                        if (-not $classification.IsGenuine) {
                            Write-Host "WARNING: Non-genuine battery detected" -ForegroundColor Red
                        }
                        if ($classification.FailureFlag) {
                            Write-Host "ALERT: Battery replacement needed immediately!" -ForegroundColor Red
                        }

                        Write-Host ""
                        $index++
                    }
                    catch {
                        Log-BatteryError "root\wmi" "BatteryFullChargedCapacity" $_ $index $($bat.InstanceName) $null
                        Write-Host "Battery #: $index - FAILED" -ForegroundColor Red
                        Write-Host ""
                        $index++
                    }
                }
            }
        }
        catch {
            Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
            Write-Host "ACPI battery data unavailable." -ForegroundColor Yellow
        }
    }

    Read-Host "Press ENTER"
}

function DiagnosticInfo {
    Show-Header
    Write-Host "[ Diagnostic Information ]"
    Write-Host ""

    Write-Host "Checking system WMI configuration..." -ForegroundColor Cyan
    Write-Host ""

    # Check BatteryStaticData via Get-WmiObject (CimInstance fails on this class)
    Write-Host "BatteryStaticData (ACPI)               : " -NoNewline
    if (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryStaticData") {
        Write-Host "Available" -ForegroundColor Green
    }
    else {
        Write-Host "Not available" -ForegroundColor Red
    }

    # Show design capacity resolution diagnostic
    Write-Host ""
    Write-Host "[ Design Capacity Resolution ]" -ForegroundColor Cyan

    # Test Source 1: Get-WmiObject BatteryStaticData
    Write-Host "  Source 1 - BatteryStaticData (WmiObject) : " -NoNewline
    try {
        $sd = @(Get-WmiObject -Namespace root\wmi -Class BatteryStaticData -ErrorAction SilentlyContinue)
        if ($sd -and $sd.Count -gt 0 -and $sd[0].DesignedCapacity -gt 0) {
            Write-Host "$($sd[0].DesignedCapacity) mWh" -ForegroundColor Green
        } else {
            Write-Host "Returned 0 or empty" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Test Source 2: powercfg XML
    Write-Host "  Source 2 - powercfg /batteryreport /XML : " -NoNewline
    try {
        $xmlPath = Join-Path $env:TEMP "battery_diag_$PID.xml"
        $null = & powercfg /batteryreport /XML /OUTPUT $xmlPath 2>$null
        if (Test-Path $xmlPath) {
            [xml]$report = Get-Content $xmlPath -ErrorAction SilentlyContinue
            Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
            $batNode = @($report.BatteryReport.Batteries.Battery)
            if ($batNode -and $batNode[0].DesignCapacity -gt 0) {
                Write-Host "$($batNode[0].DesignCapacity) mWh" -ForegroundColor Green
            } else {
                Write-Host "Returned 0 or empty" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Report file not created" -ForegroundColor Red
        }
    } catch {
        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Show final resolved value
    Write-Host "  Resolved Design Capacity              : " -NoNewline
    $dcVal = Get-DesignCapacity -Index 0
    if ($dcVal) {
        Write-Host "$dcVal mWh" -ForegroundColor Green
    } else {
        Write-Host "Not found from any source" -ForegroundColor Red
    }
    Write-Host ""

    # Check Lenovo System Interface Foundation (SIF)
    # SIF is the driver package that registers root\Lenovo. Options 2 and 5
    # depend on it. We check SIF installation and namespace availability
    # independently so a partial install (SIF present but namespace broken)
    # is reported accurately rather than lumped with "not installed".
    Write-Host "[ Lenovo System Interface Foundation ]" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "  SIF Installed                         : " -NoNewline
    $sifInstalled = Test-SifInstalled
    $deviceFamily = Get-LenovoDeviceFamily
    if ($sifInstalled) {
        Write-Host "Yes" -ForegroundColor Green
    }
    else {
        Write-Host "No" -ForegroundColor Red
        Write-Host ""
        if ($deviceFamily -eq "Other") {
            Write-Host "  Note: SIF is a ThinkPad-exclusive driver package." -ForegroundColor Yellow
            Write-Host "  root\Lenovo is not available on this device by design." -ForegroundColor Yellow
        } else {
            Write-Host "  SIF is required for Options 2 (Lenovo Battery Cycles)" -ForegroundColor Yellow
            Write-Host "  and 5 (Comprehensive Battery Analysis)." -ForegroundColor Yellow
            Write-Host "  Install SIF from support.lenovo.com -> Drivers & Software." -ForegroundColor Yellow
        }
    }

    Write-Host "  root\Lenovo Namespace                 : " -NoNewline
    if (Test-LenovoNamespace) {
        Write-Host "Available" -ForegroundColor Green
        
        Write-Host ""
        Write-Host "  Lenovo Classes:"
        
        Write-Host "    - Lenovo_Odometer                 : " -NoNewline
        if (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer") {
            Write-Host "Available" -ForegroundColor Green
        }
        else {
            Write-Host "Not available" -ForegroundColor Red
        }
        
        Write-Host "    - Lenovo_WarrantyInformation      : " -NoNewline
        if (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_WarrantyInformation") {
            Write-Host "Available" -ForegroundColor Green
        }
        else {
            Write-Host "Not available" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Not available" -ForegroundColor Red
        Write-Host ""
        if ($deviceFamily -eq "Other") {
            Write-Host "  root\Lenovo is a ThinkPad-exclusive firmware feature." -ForegroundColor Yellow
            Write-Host "  It is not present on this device by design." -ForegroundColor Yellow
        } elseif ($sifInstalled) {
            Write-Host "  SIF appears installed but root\Lenovo was not found." -ForegroundColor Yellow
            Write-Host "  To resolve:" -ForegroundColor Yellow
            Write-Host "    1. Restart WMI service:  Restart-Service winmgmt"
            Write-Host "    2. Reinstall SIF from support.lenovo.com"
        }
        else {
            Write-Host "  SIF is not installed. This is expected — the namespace" -ForegroundColor Yellow
            Write-Host "  is created by SIF during driver installation." -ForegroundColor Yellow
        }
    }

    Write-Host ""

    # ── CDRT Odometer (Commercial Vantage) ──────────────────────────────
    # The CDRT Odometer is deployed by Commercial Vantage on top of SIF.
    # We report installation status and which CDRT fields are resolvable.
    Write-Host "[ CDRT Odometer (Commercial Vantage) ]" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "  Commercial Vantage Detected           : " -NoNewline
    $cvInstalled = Test-CommercialVantage
    if ($cvInstalled) {
        Write-Host "Yes" -ForegroundColor Green
    }
    else {
        Write-Host "No" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Commercial Vantage is required for Option 11 (CDRT Odometer)." -ForegroundColor Yellow
        Write-Host "  It deploys the CDRT MOF that adds CPU uptime, shock, and thermal" -ForegroundColor Yellow
        Write-Host "  event tracking to Lenovo_Odometer." -ForegroundColor Yellow
    }

    Write-Host "  CDRT Fields Available                 : " -NoNewline
    if ($cvInstalled -and (Test-LenovoNamespace)) {
        $cdrtData = Get-CdrtOdometerData
        if ($cdrtData.Available) {
            Write-Host "Yes" -ForegroundColor Green
            Write-Host "    - CPU Uptime    : " -NoNewline
            Write-Host $(if ($null -ne $cdrtData.CpuUptimeMinutes) { "~$([math]::Round($cdrtData.CpuUptimeMinutes / 60, 1)) hours ($($cdrtData.CpuUptimeMinutes) min raw)" } else { "Not found" }) -ForegroundColor $(if ($null -ne $cdrtData.CpuUptimeMinutes) { "Green" } else { "Yellow" })
            Write-Host "    - Shock Events  : " -NoNewline
            Write-Host $(if ($null -ne $cdrtData.ShockEvents) { $cdrtData.ShockEvents } else { "Not found" }) -ForegroundColor $(if ($null -ne $cdrtData.ShockEvents) { "Green" } else { "Yellow" })
            Write-Host "    - Thermal Events: " -NoNewline
            Write-Host $(if ($null -ne $cdrtData.ThermalEvents) { $cdrtData.ThermalEvents } else { "Not found" }) -ForegroundColor $(if ($null -ne $cdrtData.ThermalEvents) { "Green" } else { "Yellow" })
        }
        else {
            Write-Host "No" -ForegroundColor Yellow
            Write-Host "    $($cdrtData.UnavailableReason)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "No" -ForegroundColor Yellow
    }

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
    if (Test-LenovoNamespace) {
        try {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            $cycles = ([string]$odo.Battery_cycles -replace '\D','')
            $hasLenovoData = $true
            
            Write-Host "[ Lenovo EC Data ]"
            Write-Host "Cycle Count : $cycles"
            Write-Host "Shock Events: $($odo.Shock_events)"
            Write-Host "Thermal     : $($odo.Thermal_events)"
            Write-Host ""
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Odometer" $_
            Write-Host "Lenovo cycle data unavailable." -ForegroundColor Yellow
            Write-Host ""
        }
    }
    else {
        # root\Lenovo namespace missing — show device-aware explanation.
        # ACPI capacity data will still be shown below if available.
        Write-Host "[ Lenovo EC Data ]" -ForegroundColor Yellow
        Get-LenovoNamespaceUnavailableMessage -Feature "battery cycle count and EC data"
        Write-Host ""
    }

    # Try to get ACPI battery capacity data
    if (Test-WmiClass -Namespace "root\wmi" -ClassName "BatteryFullChargedCapacity") {
        try {
            $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)

            if ($batteries -and $batteries.Count -gt 0) {
                $hasACPIData = $true
                Write-Host "[ ACPI Battery Capacity Data ]"
                Write-Host ""

                $index = 0
                foreach ($bat in $batteries) {
                    try {
                        $FullCharge = Get-SafeWmiProperty -Object $bat -PropertyName "FullChargedCapacity"
                        
                        # Get design capacity and manufacturer via multi-source helpers
                        $DesignCapValue = Get-DesignCapacity -Index $index
                        $Manufacturer   = Get-BatteryManufacturer -Index $index
                        
                        $DesignCap = if ($DesignCapValue) { $DesignCapValue } else { "Unavailable" }
                        
                        Write-Host "Battery #: $index"
                        Write-Host ("-" * 50)
                        Write-Host "Instance Name        : $($bat.InstanceName)"
                        Write-Host "Manufacturer         : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })"
                        Write-Host "Full Charge Capacity : $FullCharge"
                        Write-Host "Design Capacity      : $DesignCap"
                        
                        # Get comprehensive classification with cycle data
                        $classification = Get-BatteryClassification -DesignCapacity $DesignCapValue -FullChargeCapacity $FullCharge -CycleCount $cycles -Manufacturer $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })
                        
                        Write-Host ""
                        Write-Host "[ Battery Health Classification ]"
                        
                        $severityColor = switch ($classification.Severity) {
                             0 { "Green" }
                            1 { "Cyan" }
                            2 { "Yellow" }
                            3 { "Magenta" }
                            4 { "Red" }
                            default { "White" }
                        }
                        
                        Write-Host "Health Percentage    : " -NoNewline
                        Write-Host "$($classification.HealthPercent)%" -ForegroundColor $severityColor
                        
                        Write-Host "Classification       : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "Status               : " -NoNewline
                        Write-Host $classification.Message -ForegroundColor $severityColor
                        
                        if (-not $classification.IsGenuine) {
                            Write-Host "WARNING: Non-genuine battery detected" -ForegroundColor Red
                        }
                        
                        if ($classification.FailureFlag) {
                            Write-Host "ALERT: Battery replacement needed immediately!" -ForegroundColor Red
                        }
                        
                        Write-Host ""
                        $index++
                    }
                    catch {
                        Log-BatteryError "root\wmi" "BatteryFullChargedCapacity" $_ $index $($bat.InstanceName) $null
                        Write-Host "Battery #: $index - Analysis failed" -ForegroundColor Red
                        Write-Host ""
                        $index++
                    }
                }
            }
            else {
                Write-Host "No ACPI battery capacity data available." -ForegroundColor Yellow
                Write-Host ""
            }
        }
        catch {
            Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
            Write-Host "ACPI battery data unavailable." -ForegroundColor Yellow
            Write-Host ""
        }
    }

    # Try Win32_Battery as fallback
    if (-not $hasACPIData) {
        try {
            $batteries = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery -ErrorAction SilentlyContinue)

            if ($batteries -and $batteries.Count -gt 0) {
                Write-Host "[ Windows Battery (Fallback) ]"
                Write-Host ""

                foreach ($bat in $batteries) {
                    try {
                        $percent = Get-SafeWmiProperty -Object $bat -PropertyName "EstimatedChargeRemaining"
                        $deviceID = Get-SafeWmiProperty -Object $bat -PropertyName "DeviceID"
                        $manufacturer = Get-SafeWmiProperty -Object $bat -PropertyName "Manufacturer"
                        $status = Get-SafeWmiProperty -Object $bat -PropertyName "Status"

                        Write-Host "Battery #: $deviceID"
                        Write-Host "Estimated Charge : $percent %"
                        
                        $classification = Get-BatteryClassification -DesignCapacity $null -FullChargeCapacity $percent -CycleCount $cycles -BatteryStatus $status -Manufacturer $manufacturer
                        
                        $severityColor = switch ($classification.Severity) {
                             0 { "Green" }
                            1 { "Cyan" }
                            2 { "Yellow" }
                            3 { "Magenta" }
                            4 { "Red" }
                            default { "White" }
                        }
                        
                        Write-Host "Classification     : " -NoNewline
                        Write-Host $classification.Classification -ForegroundColor $severityColor
                        Write-Host "Battery Status      : $status"
                        
                        if (-not $classification.IsGenuine) {
                            Write-Host "WARNING: Non-genuine battery detected" -ForegroundColor Red
                        }
                        
                        Write-Host ""
                    }
                    catch {
                        Log-BatteryError "root\cimv2" "Win32_Battery" $_ $null $($bat.DeviceID) $null
                        Write-Host "Battery analysis failed." -ForegroundColor Red
                        Write-Host ""
                    }
                }
            }
        }
        catch {
            Log-Error "root\cimv2" "Win32_Battery" $_
            Write-Host "No battery data available from any source." -ForegroundColor Red
        }
    }
    
    if (-not $hasLenovoData -and -not $hasACPIData) {
        Write-Host "Unable to retrieve comprehensive battery data." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This requires:" -ForegroundColor Yellow
        Write-Host "  • Lenovo WMI provider (for cycle data)" -ForegroundColor Yellow
        Write-Host "  • ACPI battery support (for capacity data)" -ForegroundColor Yellow
    }

    Read-Host "Press ENTER"
}

function ExportReport {

    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = "$env:USERPROFILE\Desktop\ThinkPad_Report_$timestamp.txt"
    $errorPath  = "$env:USERPROFILE\Desktop\ThinkPad_ErrorLog_$timestamp.txt"

    [System.Collections.Generic.List[string]]$Report = @()

    $si = Get-SystemInfo

    $Report += "==============================================="
    $Report += " THINKPAD UTILITY MANAGEMENT — DIAGNOSTIC REPORT"
    $Report += " Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $Report += " Device     : $env:COMPUTERNAME"
    $Report += " Model      : $($si.Model)"
    $Report += " BIOS       : $($si.BIOSVersion)  ($($si.BIOSDate))"
    $Report += " Environment: $($si.Environment)$(if ($si.DomainName) { "  ($($si.DomainName))" })"
    $Report += "==============================================="
    $Report += ""

    # ════════════════════════════════════════════════
    # [ SYSTEM INFO ]
    # ════════════════════════════════════════════════
    $Report += "[ SYSTEM INFO ]"
    $Report += "  Device      : $env:COMPUTERNAME"
    $Report += "  Model       : $($si.Model)"
    $Report += "  BIOS        : $($si.BIOSVersion)  ($($si.BIOSDate))"
    $Report += "  Environment : $($si.Environment)$(if ($si.DomainName) { "  ($($si.DomainName))" })"
    $Report += ""

    # ════════════════════════════════════════════════
    # [ BATTERY RAW DATA ]
    # ════════════════════════════════════════════════
    $Report += "[ BATTERY RAW DATA ]"
    try {
        if (Test-WmiClass -ClassName "Win32_Battery") {
            $batteries = @(Get-CimInstance -CimSession $script:CimSession -ClassName Win32_Battery -ErrorAction SilentlyContinue)
            if (-not $batteries -or $batteries.Count -eq 0) {
                $Report += "  No batteries found."
            } else {
                foreach ($bat in $batteries) {
                    $Report += "  Battery ID     : $(Get-SafeWmiProperty -Object $bat -PropertyName 'DeviceID')"
                    $Report += "  Manufacturer   : $(Get-SafeWmiProperty -Object $bat -PropertyName 'Manufacturer')"
                    $Report += "  Serial         : $(Get-SafeWmiProperty -Object $bat -PropertyName 'SerialNumber')"
                    $Report += "  Charge         : $(Get-SafeWmiProperty -Object $bat -PropertyName 'EstimatedChargeRemaining') %"
                    $Report += ""
                }
            }
        } else {
            $Report += "  Win32_Battery not available."
        }
    } catch {
        Log-Error "root\cimv2" "Win32_Battery" $_
        $Report += "  Win32_Battery query failed."
    }

    if (Test-LenovoNamespace) {
        try {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            $cycles = ([string]$odo.Battery_cycles -replace '\D','')
            $Report += "  Cycle Count    : $cycles"
            $Report += "  Shock Events   : $($odo.Shock_events)"
            $Report += "  Thermal Events : $($odo.Thermal_events)"
        } catch {
            Log-Error "root\Lenovo" "Lenovo_Odometer" $_
            $Report += "  Lenovo EC data unavailable."
        }
    } else {
        $Report += "  Lenovo EC      : root\Lenovo namespace not available."
    }
    $Report += ""

    # ════════════════════════════════════════════════
    # [ BATTERY ANALYSIS ]
    # ════════════════════════════════════════════════
    $Report += "[ BATTERY ANALYSIS ]"
    try {
        $reportCycles = 0
        if (Test-LenovoNamespace) {
            try {
                $odo2 = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue
                $reportCycles = ([string]$odo2.Battery_cycles -replace '\D','')
            } catch {}
        }

        $batteries = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
        if ($batteries -and $batteries.Count -gt 0) {
            $index = 0
            foreach ($bat in $batteries) {
                try {
                    $FullCharge     = Get-SafeWmiProperty -Object $bat -PropertyName "FullChargedCapacity"
                    $DesignCapValue = Get-DesignCapacity -Index $index
                    $Manufacturer   = Get-BatteryManufacturer -Index $index

                    $classification = Get-BatteryClassification `
                        -DesignCapacity $DesignCapValue -FullChargeCapacity $FullCharge `
                        -CycleCount $reportCycles -Manufacturer $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })

                    $healthStatus = if ($classification.Severity -eq 0)     { "PASS" }
                                    elseif ($classification.Severity -le 1)  { "WARN" }
                                    else                                      { "FAIL" }

                    $cycleStatus  = if ($reportCycles -eq 0)         { "INFO" }
                                    elseif ($reportCycles -lt 300)   { "PASS" }
                                    elseif ($reportCycles -lt 500)   { "WARN" }
                                    else                             { "FAIL" }

                    $genuineStatus = if ($classification.IsGenuine) { "PASS" } else { "WARN" }

                    $Report += "  Battery $($index + 1)"
                    $Report += "  Manufacturer   : $(if ($Manufacturer) { $Manufacturer } else { "Unknown" })"
                    $Report += "  Health         : $($classification.HealthPercent)%  [$healthStatus]"
                    $Report += "  Cycle Count    : $(if ($reportCycles -gt 0) { "$reportCycles" } else { "N/A" })  [$cycleStatus]"
                    $Report += "  Rating         : $($classification.Classification)"
                    $Report += "  Genuine        : $(if ($classification.IsGenuine) { "Yes" } else { "No" })  [$genuineStatus]"
                    $Report += "  Summary        : $($classification.Message)"
                    $Report += ""
                    $index++
                } catch {
                    $Report += "  Battery $($index + 1) — analysis failed."
                    $Report += ""
                    $index++
                }
            }
        } else {
            $Report += "  No ACPI battery capacity data available."
        }
    } catch {
        $Report += "  Battery analysis could not be completed."
    }
    $Report += ""

    # ════════════════════════════════════════════════
    # [ LENOVO EC STATUS ]
    # ════════════════════════════════════════════════
    $Report += "[ LENOVO EC STATUS ]"
    if (Test-LenovoNamespace) {
        try {
            $w = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_WarrantyInformation -ErrorAction Stop
            $Report += "  Warranty Serial : $($w.SerialNumber)"
            $Report += "  Warranty Start  : $($w.StartDate)"
            $Report += "  Warranty End    : $($w.EndDate)"
            $Report += "  Last Sync       : $($w.LastUpdateTime)"
        } catch {
            $Report += "  Warranty data unavailable."
        }
        $cdrt = Get-CdrtOdometerData
        if ($cdrt.Available) {
            $Report += "  CPU Uptime      : $(if ($null -ne $cdrt.CpuUptimeMinutes) { "~$([math]::Round($cdrt.CpuUptimeMinutes / 60, 1)) hours" } else { "N/A" })"
            $Report += "  Shock Events    : $(if ($null -ne $cdrt.ShockEvents) { $cdrt.ShockEvents } else { "N/A" })"
            $Report += "  Thermal Events  : $(if ($null -ne $cdrt.ThermalEvents) { $cdrt.ThermalEvents } else { "N/A" })"
        } else {
            $Report += "  CDRT Odometer   : Vantage not installed — skipped."
        }
    } else {
        $Report += "  root\Lenovo namespace not available on this system."
    }
    $Report += ""

    # ════════════════════════════════════════════════
    # [ MEMORY INFO ]
    # ════════════════════════════════════════════════
    $Report += "[ MEMORY INFO ]"
    try {
        $mem = Get-MemoryState
        if ($mem.ModuleCount -gt 0) {
            $memStatus = if ($mem.SpeedMismatch) { "WARN" } else { "PASS" }
            $Report += "  Total RAM      : $($mem.TotalRAMGB) GB  [$memStatus]"
            $Report += "  Modules        : $($mem.ModuleCount)"
            $Report += "  Speed Mismatch : $(if ($mem.SpeedMismatch) { "Yes  [WARN]" } else { "No  [PASS]" })"
            $modIdx = 0
            foreach ($m in $mem.Modules) {
                $Report += "  Module $($modIdx + 1)       : $($m.CapacityGB) GB  $($m.SpeedMHz) MHz  ($($m.Manufacturer))"
                $modIdx++
            }
        } else {
            $Report += "  $($mem.Message)"
        }
    } catch {
        $Report += "  Memory data unavailable."
    }
    $Report += ""

    # ════════════════════════════════════════════════
    # [ DEPLOYMENT READINESS ]
    # ════════════════════════════════════════════════
    $Report += "[ DEPLOYMENT READINESS ]"
    $dr = Get-DeploymentReadiness
    foreach ($key in $dr.Checks.Keys) {
        $chk   = $dr.Checks[$key]
        $label = $key.PadRight(20)
        $Report += "  $label : $($chk.Value)  [$($chk.Status)]"
    }
    $Report += ""
    $Report += "  Status : $($dr.Status)"
    $Report += ""

    # ── Save ─────────────────────────────────────────────────────────────
    $Report | Out-File $reportPath -Encoding UTF8

    if ($script:ErrorLog.Count -gt 0) {
        $ErrorOutput = @()
        foreach ($err in $script:ErrorLog) {
            $ErrorOutput += ("=" * 70)
            $ErrorOutput += "TimeStamp       : $($err.TimeStamp)"
            $ErrorOutput += "Namespace       : $($err.Namespace)"
            $ErrorOutput += "Class           : $($err.Class)"
            $ErrorOutput += "Message         : $($err.Message)"
            $ErrorOutput += "ScriptName      : $($err.ScriptName)"
            $ErrorOutput += "ScriptLineNumber: $($err.ScriptLineNumber)"
            if ($err.Context)       { $ErrorOutput += "Context         : $($err.Context)" }
            if ($err.BatteryDetails) {
                $ErrorOutput += "Battery Details :"
                $ErrorOutput += "  - InstanceName       : $($err.BatteryDetails.InstanceName)"
                $ErrorOutput += "  - FullChargedCapacity: $($err.BatteryDetails.FullChargedCapacity)"
                $ErrorOutput += "  - DesignedCapacity   : $($err.BatteryDetails.DesignedCapacity)"
            }
            $ErrorOutput += "StackTrace      : $($err.StackTrace)"
            $ErrorOutput += ""
        }
        $ErrorOutput | Out-File $errorPath -Encoding UTF8
    }

    Show-Header
    Write-Host "Report saved to:" -ForegroundColor Green
    Write-Host "  $reportPath"

    if ($script:ErrorLog.Count -gt 0) {
        Write-Host ""
        Write-Host "Error log saved to:" -ForegroundColor Yellow
        Write-Host "  $errorPath"
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
            $newRows | Export-Csv -Path $logPath -Append -NoTypeInformation -Force
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
        Write-Host "No trend data found yet." -ForegroundColor Yellow
        Write-Host "The log is written silently each time you run this script." -ForegroundColor DarkGray
        Write-Host "Run the script a few times over days or weeks to build a history." -ForegroundColor DarkGray
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    $rows = @()
    try { $rows = @(Import-Csv $logPath) } catch {
        Write-Host "Failed to read trend log: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    if ($rows.Count -eq 0) {
        Write-Host "Trend log is empty." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    # Group by BatteryIndex
    $indices = $rows | Select-Object -ExpandProperty BatteryIndex -Unique | Sort-Object

    foreach ($idx in $indices) {
        $battRows = @($rows | Where-Object { $_.BatteryIndex -eq $idx } | Sort-Object Timestamp)

        $battID = $battRows[-1].BatteryID
        Write-Host "Battery: $battID  (Index $idx)" -ForegroundColor Cyan
        Write-Host ("-" * 50)

        # Table header
        Write-Host ("  {0,-20} {1,8} {2,8} {3,10}" -f "Date", "Health%", "FullChg", "Cycles")
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

        # Trend analysis — need at least 2 entries with parseable timestamps
        $parsed = @()
        foreach ($row in $battRows) {
            [double]$hp = 0
            $ts = $null
            try { $ts = [datetime]::Parse($row.Timestamp) } catch { continue }
            if (-not [double]::TryParse($row.HealthPct, [ref]$hp)) { continue }
            $parsed += [PSCustomObject]@{ Timestamp = $ts; HealthPct = $hp }
        }

        if ($parsed.Count -lt 2) {
            Write-Host "  Not enough data points for trend analysis yet." -ForegroundColor DarkGray
            Write-Host "  (Need at least 2 entries recorded on different days)" -ForegroundColor DarkGray
            Write-Host ""
            continue
        }

        $first = $parsed[0]
        $last  = $parsed[-1]
        $spanDays = ($last.Timestamp - $first.Timestamp).TotalDays

        if ($spanDays -lt 1) {
            Write-Host "  All entries recorded on the same day — run again later for trend data." -ForegroundColor DarkGray
            Write-Host ""
            continue
        }

        if ($spanDays -lt 14) {
            Write-Host "  Insufficient span — need at least 14 days of data for a reliable rate." -ForegroundColor DarkGray
            Write-Host "  Span so far: $([math]::Round($spanDays)) days  ($($parsed.Count) entries)" -ForegroundColor DarkGray
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

        # Rate colour coding
        $rateColor = if ($dropPerMonth -le 1.5) { "Green" } elseif ($dropPerMonth -le 3.0) { "Yellow" } else { "Red" }
        $rateLabel = if ($dropPerMonth -le 1.5) { "Normal" } elseif ($dropPerMonth -le 3.0) { "Elevated" } else { "Accelerated" }

        Write-Host "  Degradation Rate : " -NoNewline
        Write-Host ("{0:N2}%/month  ({1})" -f $dropPerMonth, $rateLabel) -ForegroundColor $rateColor

        Write-Host "  Span             : $([math]::Round($spanDays)) days  ($($parsed.Count) entries)"
        Write-Host "  Current Health   : " -NoNewline
        $hpColor = if ($currentHP -ge 80) { "Green" } elseif ($currentHP -ge 70) { "Yellow" } elseif ($currentHP -ge 60) { "Magenta" } else { "Red" }
        Write-Host ("{0:N2}%" -f $currentHP) -ForegroundColor $hpColor

        # Projection to 80% threshold
        if ($dropPerMonth -gt 0 -and $currentHP -gt 80) {
            $monthsTo80 = [math]::Round(($currentHP - 80) / $dropPerMonth, 1)
            $dateTo80   = (Get-Date).AddDays($monthsTo80 * 30.44)
            Write-Host "  Est. 80% reached : in ~$monthsTo80 months  (approx. $($dateTo80.ToString('MMM yyyy')))" -ForegroundColor $rateColor
        } elseif ($currentHP -le 80) {
            Write-Host "  Battery is already at or below the 80% replacement threshold." -ForegroundColor Yellow
        } else {
            Write-Host "  Rate too low to project meaningfully." -ForegroundColor DarkGray
        }

        Write-Host ""
    }

    Write-Host "Log file: $logPath" -ForegroundColor DarkGray
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
            Write-Host "Cycle Count (Lenovo EC)  : $cycles"
        }
        catch {
            Log-Error "root\Lenovo" "Lenovo_Odometer" $_
            Write-Host "Cycle Count              : Unavailable (Lenovo EC not accessible)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Cycle Count              : Unavailable (Lenovo namespace not found)" -ForegroundColor Yellow
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
                Write-Host "Full Charge Capacity     : $fc mWh"
                Write-Host "Design Capacity          : $dc mWh"
                Write-Host "Health Percentage        : $health%"
            }
            else {
                Write-Host "Health Percentage        : Unavailable" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Health Percentage        : Unavailable (no ACPI battery data)" -ForegroundColor Yellow
        }
    }
    catch {
        Log-Error "root\wmi" "BatteryFullChargedCapacity" $_
        Write-Host "Health Percentage        : Unavailable" -ForegroundColor Yellow
    }

    # --- Collect capacity history from powercfg battery report ---
    Write-Host ""
    Write-Host "[ Capacity History ]"
    Write-Host "Reading battery report..." -ForegroundColor Cyan
    $history = Get-BatteryCapacityHistory

    if ($history -and $history.Count -ge 2) {
        $spanDays = [math]::Round(($history[-1].Timestamp - $history[0].Timestamp).TotalDays, 0)
        Write-Host "  Entries  : $($history.Count)  (spanning $spanDays days)" -ForegroundColor DarkGray
        Write-Host "  Oldest   : $($history[0].Timestamp.ToString('yyyy-MM-dd'))  Health: $($history[0].HealthPct)%" -ForegroundColor DarkGray
        Write-Host "  Newest   : $($history[-1].Timestamp.ToString('yyyy-MM-dd'))  Health: $($history[-1].HealthPct)%" -ForegroundColor DarkGray
        $totalDrop = [math]::Round($history[0].HealthPct - $history[-1].HealthPct, 2)
        if ($totalDrop -gt 0) {
            Write-Host "  Degraded : $totalDrop% over $spanDays days" -ForegroundColor DarkGray
        }
        elseif ($totalDrop -le 0) {
            Write-Host "  Degraded : No measurable degradation in recorded history" -ForegroundColor DarkGray
        }

        # Use cycle count from newest history entry as fallback when EC is unavailable
        if ($null -eq $cycles -and $history[-1].CycleCount -gt 0) {
            $cycles = $history[-1].CycleCount
            Write-Host "  Cycle Count (from report): $cycles" -ForegroundColor DarkGray
        }
    }
    elseif ($history -and $history.Count -eq 1) {
        Write-Host "  Only 1 entry found — insufficient for slope calculation." -ForegroundColor Yellow
    }
    else {
        Write-Host "  No capacity history available (powercfg report returned no entries)." -ForegroundColor Yellow
    }

    Write-Host ""

    # --- Run estimation ---
    $estimate = Get-BatteryAgeEstimate -CycleCount $cycles -HealthPercent $health -CapacityHistory $history

    if ($estimate.Confidence -eq "Unknown") {
        Write-Host "[ Age Estimation ]" -ForegroundColor Yellow
        Write-Host "Insufficient data to estimate battery age." -ForegroundColor Yellow
        Write-Host $estimate.Note -ForegroundColor Yellow
    }
    else {
        Write-Host "[ Age Estimation ]"
        Write-Host ("-" * 50)

        if ($estimate.HasCycleData) {
            Write-Host "Cycle-based estimate     : ~$($estimate.CycleBasedYears) year(s)"
        }
        if ($estimate.HasCapacityData) {
            Write-Host "Capacity-based estimate  : ~$($estimate.CapacityBasedYears) year(s)"
        }
        if ($estimate.HasHistoryData) {
            Write-Host "History-based estimate   : ~$($estimate.HistoryBasedYears) year(s)  ($($estimate.HistoryEntryCount) entries, $($estimate.HistorySpanDays) days)"
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
        Write-Host "Disclaimer: This is a statistical estimate based on typical Li-ion" -ForegroundColor DarkGray
        Write-Host "degradation patterns. Actual age may vary with usage habits, temperature," -ForegroundColor DarkGray
        Write-Host "and whether the battery has been replaced." -ForegroundColor DarkGray
    }

    Write-Host ""
    Read-Host "Press ENTER"
}

function Get-ChargeThreshold {
    <#
    .SYNOPSIS
    Reads battery charge threshold settings from Lenovo_BiosSetting (root\wmi).
    Returns a PSCustomObject with StartThreshold, StopThreshold, and raw setting names.
    Returns $null if the class is unavailable or no threshold settings are found.

    The CurrentSetting property format is: "SettingName,CurrentValue"
    Setting names vary by ThinkPad model and firmware generation.
    Known names: CustomChargeThreshold, BatteryChargeThresholdStart, BatteryChargeThresholdStop.
    We scan for any setting containing "threshold" or "charge" to maximize compatibility.
    #>

    $result = [PSCustomObject]@{
        StartThreshold     = $null
        StopThreshold      = $null
        StartSettingName   = $null
        StopSettingName    = $null
        RawStart           = $null
        RawStop            = $null
        Supported          = $false
        UnsupportedReason  = ""
    }

    if (-not (Test-LenovoNamespace)) {
        $result.UnsupportedReason = "Lenovo WMI namespace not available."
        return $result
    }

    try {
        $settings = @(Get-WmiObject -Namespace root\wmi -Class Lenovo_BiosSetting -ErrorAction Stop)

        if (-not $settings -or $settings.Count -eq 0) {
            $result.UnsupportedReason = "Lenovo_BiosSetting returned no settings."
            return $result
        }

        # Scan for start threshold — known setting names and patterns
        $startNames = @("BatteryChargeThresholdStart", "CustomThresholdStart", "ChargeStartThreshold")
        $startSetting = $null
        foreach ($name in $startNames) {
            $startSetting = $settings | Where-Object { $_.CurrentSetting -like "$name,*" } | Select-Object -First 1
            if ($startSetting) { break }
        }
        # Broad fallback scan
        if (-not $startSetting) {
            $startSetting = $settings | Where-Object {
                $_.CurrentSetting -match '(?i)(start|begin).*(threshold|charge)' -or
                $_.CurrentSetting -match '(?i)(threshold|charge).*(start|begin)'
            } | Select-Object -First 1
        }

        # Scan for stop threshold — known setting names and patterns
        $stopNames = @("BatteryChargeThresholdStop", "CustomThresholdStop", "ChargeStopThreshold")
        $stopSetting = $null
        foreach ($name in $stopNames) {
            $stopSetting = $settings | Where-Object { $_.CurrentSetting -like "$name,*" } | Select-Object -First 1
            if ($stopSetting) { break }
        }
        # Broad fallback scan
        if (-not $stopSetting) {
            $stopSetting = $settings | Where-Object {
                $_.CurrentSetting -match '(?i)(stop|end).*(threshold|charge)' -or
                $_.CurrentSetting -match '(?i)(threshold|charge).*(stop|end)'
            } | Select-Object -First 1
        }

        if (-not $startSetting -and -not $stopSetting) {
            $result.UnsupportedReason = "No charge threshold settings found in Lenovo_BiosSetting. This ThinkPad model may not support threshold control via WMI, or the setting names differ from known patterns."
            return $result
        }

        $result.Supported = $true

        if ($startSetting) {
            $parts = $startSetting.CurrentSetting -split ',', 2
            $result.StartSettingName = $parts[0].Trim()
            $result.RawStart         = $startSetting.CurrentSetting
            [int]$sv = 0
            if ($parts.Count -ge 2 -and [int]::TryParse($parts[1].Trim(), [ref]$sv)) {
                $result.StartThreshold = $sv
            }
        }

        if ($stopSetting) {
            $parts = $stopSetting.CurrentSetting -split ',', 2
            $result.StopSettingName = $parts[0].Trim()
            $result.RawStop         = $stopSetting.CurrentSetting
            [int]$sv = 0
            if ($parts.Count -ge 2 -and [int]::TryParse($parts[1].Trim(), [ref]$sv)) {
                $result.StopThreshold = $sv
            }
        }

        return $result
    }
    catch {
        Log-Error "root\wmi" "Lenovo_BiosSetting" $_
        $result.UnsupportedReason = "Failed to query Lenovo_BiosSetting: $($_.Exception.Message)"
        return $result
    }
}

function Set-ChargeThreshold {
    <#
    .SYNOPSIS
    Writes a charge threshold value to Lenovo_SetBiosSetting and commits it via
    Lenovo_SaveBiosSettings. Returns the return code string from the WMI method.

    The setting string format is: "SettingName,Value" (case-sensitive).
    Changes take effect on next reboot — the EC reads thresholds at boot time.

    Return codes from Lenovo WMI:
      "Success"          - Change staged successfully, reboot required
      "Not Supported"    - This setting is not supported on this model
      "Invalid Parameter"- Value or setting name is wrong
      "Access Denied"    - Supervisor password required
      "BIOS Error"       - Firmware-level failure
    #>
    param(
        [string]$SettingName,
        [int]$Value,
        [string]$SupervisorPassword = ""
    )

    try {
        $setter = Get-WmiObject -Namespace root\wmi -Class Lenovo_SetBiosSetting -ErrorAction Stop

        # Build the setting string — with or without supervisor password
        $settingString = if ($SupervisorPassword -ne "") {
            "$SettingName,$Value,$SupervisorPassword,ascii,us"
        } else {
            "$SettingName,$Value"
        }

        $setResult = $setter.SetBiosSetting($settingString)
        $returnCode = $setResult.Return

        if ($returnCode -eq "Success") {
            # Commit the staged change
            $saver = Get-WmiObject -Namespace root\wmi -Class Lenovo_SaveBiosSettings -ErrorAction Stop
            if ($SupervisorPassword -ne "") {
                $saver.SaveBiosSettings("$SupervisorPassword,ascii,us") | Out-Null
            } else {
                $saver.SaveBiosSettings() | Out-Null
            }
        }

        return $returnCode
    }
    catch {
        Log-Error "root\wmi" "Lenovo_SetBiosSetting" $_
        return "Exception: $($_.Exception.Message)"
    }
}

function ChargeThresholdManager {
    Show-Header
    Write-Host "[ Battery Charge Threshold Manager ]"
    Write-Host ""

    # ── Check Lenovo namespace ───────────────────────────────────────────
    if (-not (Test-LenovoNamespace)) {
        Get-LenovoNamespaceUnavailableMessage -Feature "charge threshold control"
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    # ── Read current thresholds ──────────────────────────────────────────
    Write-Host "Reading current charge thresholds..." -ForegroundColor Cyan
    $thresholds = Get-ChargeThreshold

    if (-not $thresholds.Supported) {
        Write-Host ""
        Write-Host "Charge threshold control is not available on this system." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Reason: $($thresholds.UnsupportedReason)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Note: Not all ThinkPad models expose threshold settings via WMI." -ForegroundColor DarkGray
        Write-Host "      If your model supports thresholds, use Lenovo Vantage instead." -ForegroundColor DarkGray
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    Write-Host ""
    Write-Host "[ Current Thresholds ]"
    Write-Host ("-" * 50)

    if ($null -ne $thresholds.StartThreshold) {
        Write-Host "Start Charging At : " -NoNewline
        Write-Host "$($thresholds.StartThreshold)%" -ForegroundColor Cyan
        Write-Host "  Setting name    : $($thresholds.StartSettingName)" -ForegroundColor DarkGray
    } else {
        Write-Host "Start Threshold   : Not available" -ForegroundColor Yellow
    }

    if ($null -ne $thresholds.StopThreshold) {
        Write-Host "Stop Charging At  : " -NoNewline
        Write-Host "$($thresholds.StopThreshold)%" -ForegroundColor Cyan
        Write-Host "  Setting name    : $($thresholds.StopSettingName)" -ForegroundColor DarkGray
    } else {
        Write-Host "Stop Threshold    : Not available" -ForegroundColor Yellow
    }

    Write-Host ""

    # ── Warning block ────────────────────────────────────────────────────
    Write-Host ("=" * 50) -ForegroundColor Yellow
    Write-Host "  WARNING — READ BEFORE CHANGING THRESHOLDS" -ForegroundColor Yellow
    Write-Host ("=" * 50) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  • Changes are written directly to BIOS/EC firmware." -ForegroundColor Yellow
    Write-Host "  • The new values take effect on the NEXT REBOOT." -ForegroundColor Yellow
    Write-Host "  • Start threshold MUST be lower than Stop threshold." -ForegroundColor Yellow
    Write-Host "  • Recommended range: Start 40-75%, Stop 60-80%." -ForegroundColor Yellow
    Write-Host "  • Setting Stop to 100% disables the upper limit." -ForegroundColor Yellow
    Write-Host "  • A BIOS Supervisor Password may be required." -ForegroundColor Yellow
    Write-Host "  • This tool is NOT affiliated with Lenovo." -ForegroundColor Yellow
    Write-Host ""
    Write-Host ("=" * 50) -ForegroundColor Yellow
    Write-Host ""

    $proceed = Read-Host "Change thresholds? Type 'Y' to continue or any other key to cancel"
    if ($proceed.Trim().ToUpper() -ne "Y") {
        Write-Host "Cancelled." -ForegroundColor DarkGray
        Read-Host "Press ENTER"
        return
    }

    # ── Collect new values ───────────────────────────────────────────────
    Write-Host ""

    # Start threshold
    $newStart = $null
    if ($thresholds.StartSettingName) {
        do {
            $input = Read-Host "Enter new Start threshold (1-99, or press ENTER to keep $($thresholds.StartThreshold)%)"
            if ($input.Trim() -eq "") {
                $newStart = $null  # keep current
                break
            }
            [int]$parsed = 0
            if ([int]::TryParse($input.Trim(), [ref]$parsed) -and $parsed -ge 1 -and $parsed -le 99) {
                $newStart = $parsed
                break
            }
            Write-Host "Invalid value. Enter a number between 1 and 99." -ForegroundColor Red
        } while ($true)
    }

    # Stop threshold
    $newStop = $null
    if ($thresholds.StopSettingName) {
        do {
            $input = Read-Host "Enter new Stop threshold (2-100, or press ENTER to keep $($thresholds.StopThreshold)%)"
            if ($input.Trim() -eq "") {
                $newStop = $null  # keep current
                break
            }
            [int]$parsed = 0
            if ([int]::TryParse($input.Trim(), [ref]$parsed) -and $parsed -ge 2 -and $parsed -le 100) {
                $newStop = $parsed
                break
            }
            Write-Host "Invalid value. Enter a number between 2 and 100." -ForegroundColor Red
        } while ($true)
    }

    # ── Validate start < stop ────────────────────────────────────────────
    $effectiveStart = if ($null -ne $newStart) { $newStart } else { $thresholds.StartThreshold }
    $effectiveStop  = if ($null -ne $newStop)  { $newStop  } else { $thresholds.StopThreshold  }

    if ($null -ne $effectiveStart -and $null -ne $effectiveStop -and $effectiveStart -ge $effectiveStop) {
        Write-Host ""
        Write-Host "Error: Start threshold ($effectiveStart%) must be lower than Stop threshold ($effectiveStop%)." -ForegroundColor Red
        Write-Host "No changes were made." -ForegroundColor Red
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    if ($null -eq $newStart -and $null -eq $newStop) {
        Write-Host ""
        Write-Host "No changes entered. Nothing was written." -ForegroundColor DarkGray
        Read-Host "Press ENTER"
        return
    }

    # ── Supervisor password ──────────────────────────────────────────────
    Write-Host ""
    Write-Host "If a BIOS Supervisor Password is set, enter it below." -ForegroundColor DarkGray
    Write-Host "Leave blank if no Supervisor Password is configured." -ForegroundColor DarkGray
    $svpInput = Read-Host "Supervisor Password (leave blank if none)"
    $svp = $svpInput.Trim()

    # ── Final confirmation ───────────────────────────────────────────────
    Write-Host ""
    Write-Host "Summary of changes to be written:" -ForegroundColor Cyan
    if ($null -ne $newStart) {
        Write-Host "  Start threshold : $($thresholds.StartThreshold)% -> $newStart%" -ForegroundColor White
    }
    if ($null -ne $newStop) {
        Write-Host "  Stop threshold  : $($thresholds.StopThreshold)% -> $newStop%" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "Changes take effect on next reboot." -ForegroundColor Yellow
    Write-Host ""

    $confirm = Read-Host "Type 'CONFIRM' to write to firmware, or any other key to cancel"
    if ($confirm.Trim().ToUpper() -ne "CONFIRM") {
        Write-Host "Cancelled. No changes were written." -ForegroundColor DarkGray
        Read-Host "Press ENTER"
        return
    }

    # ── Apply changes ────────────────────────────────────────────────────
    Write-Host ""
    $allSuccess = $true

    if ($null -ne $newStart -and $thresholds.StartSettingName) {
        Write-Host "Writing Start threshold..." -NoNewline
        $rc = Set-ChargeThreshold -SettingName $thresholds.StartSettingName -Value $newStart -SupervisorPassword $svp
        if ($rc -eq "Success") {
            Write-Host " OK" -ForegroundColor Green
        } else {
            Write-Host " FAILED ($rc)" -ForegroundColor Red
            $allSuccess = $false
        }
    }

    if ($null -ne $newStop -and $thresholds.StopSettingName) {
        Write-Host "Writing Stop threshold..." -NoNewline
        $rc = Set-ChargeThreshold -SettingName $thresholds.StopSettingName -Value $newStop -SupervisorPassword $svp
        if ($rc -eq "Success") {
            Write-Host " OK" -ForegroundColor Green
        } else {
            Write-Host " FAILED ($rc)" -ForegroundColor Red
            $allSuccess = $false
        }
    }

    Write-Host ""

    if ($allSuccess) {
        Write-Host "Thresholds written successfully." -ForegroundColor Green
        Write-Host "Reboot your ThinkPad to apply the new charge thresholds." -ForegroundColor Yellow
    } else {
        Write-Host "One or more values failed to write." -ForegroundColor Red
        Write-Host ""
        Write-Host "Common causes:" -ForegroundColor Yellow
        Write-Host "  • Incorrect or missing Supervisor Password" -ForegroundColor Yellow
        Write-Host "  • Setting not supported on this firmware version" -ForegroundColor Yellow
        Write-Host "  • Script not running as Administrator" -ForegroundColor Yellow
    }

    Write-Host ""
    Read-Host "Press ENTER"
}


function Show-About {
    Show-Header
    Write-Host "[ About ]"
    Write-Host ""

    # Version & build
    Write-Host "ThinkPad Utility Management" -ForegroundColor Cyan
    Write-Host "Version 1.1" -ForegroundColor White
    Write-Host ""

    # Built with
    Write-Host "Built with Windows PowerShell, ChatGPT GPT 5.4, and Claude Sonnet 4.6" -ForegroundColor DarkGray
    Write-Host ""

    # Description
    Write-Host "A read-only diagnostic tool for ThinkPad systems." -ForegroundColor Gray
    Write-Host "Reads battery health, cycle data, warranty, and memory" -ForegroundColor Gray
    Write-Host "information via Windows WMI and Lenovo EC interfaces." -ForegroundColor Gray
    Write-Host ""

    Write-Host ("=" * 47)
    Write-Host ""

    # Legal notices - plain text, no hyperlinks (console environment)
    Write-Host "Legal" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This project is not affiliated with or endorsed" -ForegroundColor DarkGray
    Write-Host "by Lenovo. ThinkPad is a trademark of Lenovo." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "This script is provided for educational and" -ForegroundColor DarkGray
    Write-Host "diagnostic purposes only. It makes no modifications" -ForegroundColor DarkGray
    Write-Host "to system configuration, firmware, or settings." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Lenovo is not responsible for the performance or" -ForegroundColor DarkGray
    Write-Host "safety of unauthorized batteries, and provides no" -ForegroundColor DarkGray
    Write-Host "warranties for failure or damage arising from" -ForegroundColor DarkGray
    Write-Host "their use." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Battery age estimation results are statistical" -ForegroundColor DarkGray
    Write-Host "approximations. Actual age may vary based on usage" -ForegroundColor DarkGray
    Write-Host "habits, temperature, and service history." -ForegroundColor DarkGray
    Write-Host ""

    Write-Host ("=" * 47)
    Write-Host ""

    # Third-party notices
    Write-Host "Third-Party Notices" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Battery degradation model based on published" -ForegroundColor DarkGray
    Write-Host "Li-ion cycle life research. Design capacity" -ForegroundColor DarkGray
    Write-Host "resolved via powercfg /batteryreport (Windows" -ForegroundColor DarkGray
    Write-Host "built-in utility) as a fallback data source." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Genuine manufacturer detection list sourced from" -ForegroundColor DarkGray
    Write-Host "Lenovo Material Safety Data Sheets and ThinkPad" -ForegroundColor DarkGray
    Write-Host "community hardware research." -ForegroundColor DarkGray
    Write-Host ""

    Write-Host ("=" * 47)
    Write-Host ""

    # Privacy notice
    Write-Host "Privacy" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This script runs entirely on the local machine." -ForegroundColor DarkGray
    Write-Host "No data is transmitted, collected, or stored" -ForegroundColor DarkGray
    Write-Host "outside of the optional local export report." -ForegroundColor DarkGray
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

function Test-CommercialVantage {
    <#
    .SYNOPSIS
    Detects whether Lenovo Commercial Vantage (or its predecessor, Lenovo Vantage
    Enterprise) is installed and has deployed the CDRT Odometer WMI classes.

    .DESCRIPTION
    The CDRT (Commercial Deployment Readiness Tool) Odometer solution is deployed
    by Lenovo Commercial Vantage. It registers custom WMI classes in root\Lenovo
    that track the cumulative life story of the ThinkPad hardware: CPU uptime,
    accelerometer shock events, and thermal throttle events.

    These classes are distinct from the SIF-provided Lenovo_Odometer fields —
    they are written by the CDRT MOF files and may or may not be present even
    when SIF is installed. Commercial Vantage is common in corporate/enterprise
    environments and is typically deployed via MDM or SCCM.

    Detection strategy:
      1. Lenovo-specific registry key written by the Commercial Vantage installer.
      2. Registry uninstall key scan (HKLM Uninstall hives, 64-bit and 32-bit).
         Reads DisplayName values directly — no Win32_Product, no MSI side effects.
      3. Service presence check for the Commercial Vantage service.
      4. Direct WMI class probe on Lenovo_Odometer in root\Lenovo —
         if the class exists and exposes CPU_uptime, CDRT is deployed regardless
         of how Commercial Vantage was detected.
    Returns $true if Commercial Vantage / CDRT appears present, $false otherwise.
    #>

    # Strategy 1: Lenovo-specific registry key written by the CV installer.
    try {
        if (Test-Path "HKLM:\SOFTWARE\Lenovo\Commercial Vantage")             { return $true }
        if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Lenovo\Commercial Vantage") { return $true }
    } catch {}

    # Strategy 2: Registry uninstall key scan.
    # Avoids Win32_Product entirely — reading DisplayName from the uninstall
    # hive is instantaneous and triggers no MSI reconfiguration or repair.
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    try {
        foreach ($regPath in $uninstallPaths) {
            if (-not (Test-Path $regPath)) { continue }
            $found = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue |
                ForEach-Object { Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue } |
                Where-Object { $_.DisplayName -match "Commercial Vantage|Vantage Enterprise" } |
                Select-Object -First 1
            if ($found) { return $true }
        }
    } catch {}

    # Strategy 3: Service
    try {
        $svc = Get-Service -Name "LenovoVantageService","ImControllerService" -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match "Vantage" } |
            Select-Object -First 1
        if ($svc) { return $true }
    } catch {}

    # Strategy 4: Direct WMI class probe — CDRT fields present on the class
    try {
        $cls = Get-CimClass -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction SilentlyContinue
        if ($cls -and ($cls.CimClassProperties | Where-Object { $_.Name -match "CPU_uptime|cpu_uptime" })) {
            return $true
        }
    } catch {}

    return $false
}

function Get-CdrtOdometerData {
    <#
    .SYNOPSIS
    Queries the CDRT Odometer fields from Lenovo_Odometer in root\Lenovo.

    .DESCRIPTION
    The CDRT solution extends Lenovo_Odometer with additional properties that
    track the cumulative hardware life story of the ThinkPad beyond what the
    base SIF Odometer exposes:

      CPU_uptime       - Total cumulative minutes the CPU has been active.
                         The raw WMI value is in minutes; this function converts
                         it to hours and exposes both values on the result object.
      Shock_events     - Cumulative count of accelerometer events that exceeded
                         the CDRT vibration threshold. This includes minor bumps,
                         desk impacts, and bag movement in addition to actual drops.
                         High counts are normal for a well-travelled machine.
      Thermal_events   - Number of times the system reached a critical temperature
                         and throttled the CPU.

    Property name casing may vary between firmware and CDRT versions. This
    function probes both common variants (CPU_uptime / cpu_uptime) and falls
    back gracefully if a property is absent.

    Returns a PSCustomObject with:
      CpuUptimeMinutes - Raw WMI value in minutes, or $null if unavailable
      CpuUptimeHours   - CpuUptimeMinutes / 60 (rounded), or $null if unavailable
      ShockEvents      - Numeric value, or $null if unavailable
      ThermalEvents   - Numeric value, or $null if unavailable
      RawObject       - The full WMI object for caller inspection
      Available       - $true if the class was queried successfully
      UnavailableReason - Populated when Available is $false
    #>

    $result = [PSCustomObject]@{
        CpuUptimeMinutes  = $null
        CpuUptimeHours    = $null
        ShockEvents       = $null
        ThermalEvents     = $null
        RawObject         = $null
        Available         = $false
        UnavailableReason = ""
    }

    if (-not (Test-LenovoNamespace)) {
        $result.UnavailableReason = "root\Lenovo namespace not available. SIF is not installed."
        return $result
    }

    if (-not (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Odometer")) {
        $result.UnavailableReason = "Lenovo_Odometer class not found in root\Lenovo."
        return $result
    }

    try {
        $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
        $result.RawObject = $odo
        $result.Available = $true

        # CPU uptime — probe both known property name variants
        foreach ($propName in @("CPU_uptime", "cpu_uptime", "CPUUptime")) {
            try {
                $val = $odo.$propName
                if ($null -ne $val) {
                    [int64]$parsed = 0
                    if ([int64]::TryParse(($val -replace '\D',''), [ref]$parsed)) {
                        # Raw WMI value is in minutes. Store both the
                        # raw minutes and the derived hours so callers can
                        # display either without re-dividing.
                        $result.CpuUptimeMinutes = $parsed
                        $result.CpuUptimeHours   = [math]::Round($parsed / 60, 1)
                        break
                    }
                }
            } catch {}
        }

        # Shock events
        foreach ($propName in @("Shock_events", "shock_events", "ShockEvents")) {
            try {
                $val = $odo.$propName
                if ($null -ne $val) {
                    [int64]$parsed = 0
                    if ([int64]::TryParse(($val -replace '\D',''), [ref]$parsed)) {
                        $result.ShockEvents = $parsed
                        break
                    }
                }
            } catch {}
        }

        # Thermal events
        foreach ($propName in @("Thermal_events", "thermal_events", "ThermalEvents")) {
            try {
                $val = $odo.$propName
                if ($null -ne $val) {
                    [int64]$parsed = 0
                    if ([int64]::TryParse(($val -replace '\D',''), [ref]$parsed)) {
                        $result.ThermalEvents = $parsed
                        break
                    }
                }
            } catch {}
        }

        # If CPU uptime is still null, the CDRT MOF was not deployed —
        # SIF Odometer is present but the CDRT extension fields are missing
        if ($null -eq $result.CpuUptimeMinutes -and $null -eq $result.ShockEvents -and $null -eq $result.ThermalEvents) {
            $result.Available         = $false
            $result.UnavailableReason = "Lenovo_Odometer exists but CDRT fields (CPU_uptime, Shock_events, Thermal_events) were not found. Commercial Vantage may not be installed or the CDRT MOF has not been deployed."
        }
    }
    catch {
        Log-Error "root\Lenovo" "Lenovo_Odometer (CDRT)" $_
        $result.Available         = $false
        $result.UnavailableReason = "Query failed: $($_.Exception.Message)"
    }

    return $result
}

function Show-CdrtOdometer {
    Show-Header
    Write-Host "[ CDRT Odometer - ThinkPad Life Story ]"
    Write-Host ""

    # ── Prerequisite: root\Lenovo namespace (SIF) ────────────────────────
    if (-not (Test-LenovoNamespace)) {
        Get-LenovoNamespaceUnavailableMessage -Feature "the CDRT Odometer"
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    # ── Check Commercial Vantage ─────────────────────────────────────────
    $cvInstalled = Test-CommercialVantage
    if (-not $cvInstalled) {
        Write-Host "Lenovo Commercial Vantage is not detected on this system." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "What is the CDRT Odometer?" -ForegroundColor Cyan
        Write-Host "The Commercial Deployment Readiness Tool (CDRT) Odometer is deployed"
        Write-Host "by Lenovo Commercial Vantage. It registers custom WMI classes that"
        Write-Host "record the cumulative hardware life story of your ThinkPad, tracking:"
        Write-Host "  - CPU Uptime      : Total hours the CPU has been active"
        Write-Host "  - Shock Events    : Accelerometer events above the CDRT vibration threshold"
        Write-Host "  - Thermal Events  : Times the CPU throttled due to critical temperature"
        Write-Host ""
        Write-Host "This data is valuable for enterprise asset management and pre-owned" 
        Write-Host "ThinkPad evaluation."
        Write-Host ""
        Write-Host "How to get Lenovo Commercial Vantage:" -ForegroundColor Cyan
        Write-Host "  - Microsoft Store: search 'Lenovo Commercial Vantage'"
        Write-Host "  - support.lenovo.com -> Drivers & Software"
        Write-Host "  - Enterprise: deploy via SCCM or MDM using the Lenovo CDRT package"
        Write-Host ""
        Write-Host "Note: Commercial Vantage is intended for business/enterprise ThinkPads." -ForegroundColor DarkGray
        Write-Host "Consumer models may use Lenovo Vantage instead, which does not" -ForegroundColor DarkGray
        Write-Host "include the CDRT Odometer." -ForegroundColor DarkGray
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    # ── Query CDRT data ──────────────────────────────────────────────────
    Write-Host "Querying CDRT Odometer data..." -ForegroundColor Cyan
    Write-Host ""

    $cdrt = Get-CdrtOdometerData

    if (-not $cdrt.Available) {
        Write-Host "CDRT Odometer data unavailable." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Reason: $($cdrt.UnavailableReason)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Commercial Vantage was detected on this system, but the CDRT" -ForegroundColor Yellow
        Write-Host "Odometer fields were not found in Lenovo_Odometer." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This can occur when:" -ForegroundColor Yellow
        Write-Host "  - Commercial Vantage is installed but the CDRT MOF has not been deployed"
        Write-Host "  - The CDRT package requires a separate deployment via SCCM/MDM"
        Write-Host "  - This ThinkPad model predates CDRT Odometer support"
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    Write-Host "[ ThinkPad Life Story ]"
    Write-Host ("-" * 50)

    # CPU Uptime
    # Raw WMI value is in minutes. Convert to hours and days for display.
    Write-Host "CPU Uptime             : " -NoNewline
    if ($null -ne $cdrt.CpuUptimeMinutes) {
        $totalHours = [math]::Round($cdrt.CpuUptimeMinutes / 60, 1)
        $days       = [math]::Floor($totalHours / 24)
        $hrs        = [math]::Round($totalHours % 24, 1)
        Write-Host "$($cdrt.CpuUptimeMinutes) min  (~$totalHours hours / $days days, $hrs hrs)" -ForegroundColor Cyan
    } else {
        Write-Host "Unavailable" -ForegroundColor Yellow
    }
    Write-Host ("-" * 50) -ForegroundColor DarkGray

    # Shock Events
    # The CDRT accelerometer threshold is sensitive — counts include minor bumps,
    # vibrations, and bag movement, not just actual drops.
    # Realistic lifetime ranges for ThinkPads:
    #   Mostly desk use     :   500 -  3,000
    #   Normal portable use :  5,000 - 20,000
    #   Heavy travel        : 20,000 - 60,000
    # Thresholds:
    #   0-20,000   Green  - Normal (desk to regular portable use)
    #   20,001-60,000 Yellow - Heavy travel, worth noting
    #   60,000+    Red    - Unusually high, inspect for damage history
    Write-Host "Shock Events           : " -NoNewline
    if ($null -ne $cdrt.ShockEvents) {
        $shockColor = if     ($cdrt.ShockEvents -le 20000) { "Green"  }
                      elseif ($cdrt.ShockEvents -le 60000) { "Yellow" }
                      else                                  { "Red"    }
        $shockLabel = if     ($cdrt.ShockEvents -le 20000) { "Normal" }
                      elseif ($cdrt.ShockEvents -le 60000) { "Heavy travel range" }
                      else                                  { "Unusually high - inspect for damage history" }
        Write-Host "$($cdrt.ShockEvents)  ($shockLabel)" -ForegroundColor $shockColor
        Write-Host "  Note: Counts accelerometer events above the CDRT vibration threshold." -ForegroundColor DarkGray
        Write-Host "        Includes minor bumps and movement — not only actual drops." -ForegroundColor DarkGray
    } else {
        Write-Host "Unavailable" -ForegroundColor Yellow
    }
    Write-Host ("-" * 50) -ForegroundColor DarkGray

    # Thermal Events
    # Thresholds are calibrated to lifetime count, not to zero-tolerance:
    #   0-50   Green  - Normal across any lifespan
    #   51-200 Yellow - Moderate, worth monitoring
    #   200+   Red    - Frequent throttling, investigate cooling
    Write-Host "Thermal Events         : " -NoNewline
    if ($null -ne $cdrt.ThermalEvents) {
        $thermalColor = if     ($cdrt.ThermalEvents -le 50)  { "Green"  }
                        elseif ($cdrt.ThermalEvents -le 200) { "Yellow" }
                        else                                  { "Red"    }
        $thermalLabel = if     ($cdrt.ThermalEvents -le 50)  { "Normal" }
                        elseif ($cdrt.ThermalEvents -le 200) { "Moderate - monitor cooling" }
                        else                                  { "Frequent throttling - investigate cooling" }
        Write-Host "$($cdrt.ThermalEvents)  ($thermalLabel)" -ForegroundColor $thermalColor
        Write-Host "  Note: Each event represents a CPU throttle due to critical temperature." -ForegroundColor DarkGray
        if ($null -ne $cdrt.CpuUptimeMinutes -and $cdrt.CpuUptimeMinutes -gt 0) {
            $uptimeYears = $cdrt.CpuUptimeMinutes / 525600
            if ($uptimeYears -gt 0) {
                $ratePerYear = [math]::Round($cdrt.ThermalEvents / $uptimeYears, 1)
                Write-Host "  Rate   : ~$ratePerYear events/year over $([math]::Round($uptimeYears, 1)) years of CPU uptime" -ForegroundColor DarkGray
            }
        }
    } else {
        Write-Host "Unavailable" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Source: CDRT Odometer via Lenovo_Odometer (root\Lenovo)" -ForegroundColor DarkGray
    Write-Host "Data reflects cumulative lifetime counts since CDRT deployment." -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press ENTER"
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

function Format-PassFail {
    <#
    .SYNOPSIS
    Writes a labelled metric line with a colour-coded PASS / WARN / FAIL badge.

    Parameters
      Label   — left-hand column label (padded to 22 chars)
      Value   — the raw value string to display
      Status  — "PASS", "WARN", or "FAIL"
      Indent  — optional leading spaces (default "  ")
    #>
    param(
        [string]$Label,
        [string]$Value,
        [ValidateSet("PASS","WARN","FAIL","INFO")]
        [string]$Status = "INFO",
        [string]$Indent = "  "
    )
    $badgeColor = switch ($Status) {
        "PASS" { "Green"  }
        "WARN" { "Yellow" }
        "FAIL" { "Red"    }
        "INFO" { "DarkGray" }
    }
    $badge = switch ($Status) {
        "PASS" { "✔ PASS" }
        "WARN" { "⚠ WARN" }
        "FAIL" { "✘ FAIL" }
        "INFO" { "  INFO" }
    }
    $labelPadded = $Label.PadRight(22)
    Write-Host "$Indent$labelPadded : $Value" -NoNewline
    Write-Host "  [$badge]" -ForegroundColor $badgeColor
}

function Get-DeploymentReadiness {
    <#
    .SYNOPSIS
    Evaluates whether this ThinkPad meets deployment / resale readiness criteria
    and returns a structured result object.

    Checks performed:
      Battery health >= 80%           PASS / FAIL
      Cycle count < 500               PASS / WARN (>300) / FAIL (>=500)
      No critical battery status       PASS / FAIL
      Memory speed matched             PASS / WARN
      BIOS not unknown                 PASS / WARN
      No shock/thermal alerts (CDRT)   PASS / WARN / INFO (if unavailable)

    Returns a PSCustomObject with:
      Ready       — $true if all hard checks pass
      Status      — "READY" / "REVIEW NEEDED" / "NOT READY"
      StatusColor — console colour
      Checks      — ordered hashtable of individual check results
    #>

    $checks  = [ordered]@{}
    $hardFail = $false

    # ── Battery health ───────────────────────────────────────────────────
    $battPct  = $null
    $battFail = $false
    try {
        # Prefer Lenovo_Battery, fall back to ACPI
        if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
            $lbs = @(Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop)
            $healths = @()
            foreach ($lb in $lbs) {
                $dcStr = ((Get-SafeWmiProperty -Object $lb -PropertyName "DesignCapacity") -replace '[^0-9,\.]','') -replace ',','.'
                $fcStr = ((Get-SafeWmiProperty -Object $lb -PropertyName "FullChargeCapacity") -replace '[^0-9,\.]','') -replace ',','.'
                [double]$dc = 0; [double]$fc = 0
                if ([double]::TryParse($dcStr,[System.Globalization.NumberStyles]::Any,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$dc) -and
                    [double]::TryParse($fcStr,[System.Globalization.NumberStyles]::Any,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$fc) -and
                    $dc -gt 0) { $healths += [math]::Round(($fc/$dc)*100,1) }
            }
            if ($healths.Count -gt 0) { $battPct = ($healths | Measure-Object -Minimum).Minimum }
        }
        if ($null -eq $battPct) {
            $acpi = @(Get-WmiObject -Namespace root\wmi -Class BatteryFullChargedCapacity -ErrorAction SilentlyContinue)
            $healths = @()
            for ($i = 0; $i -lt $acpi.Count; $i++) {
                $fc = $acpi[$i].FullChargedCapacity
                $dc = Get-DesignCapacity -Index $i
                [int64]$fcN = 0; [int64]$dcN = 0
                if ($fc -and $dc -and [int64]::TryParse($fc.ToString(),[ref]$fcN) -and [int64]::TryParse($dc.ToString(),[ref]$dcN) -and $dcN -gt 0) {
                    $healths += [math]::Round(($fcN/$dcN)*100,1)
                }
            }
            if ($healths.Count -gt 0) { $battPct = ($healths | Measure-Object -Minimum).Minimum }
        }
    } catch {}

    if ($null -ne $battPct) {
        $battStatus = if ($battPct -ge 80) { "PASS" } elseif ($battPct -ge 70) { "WARN" } else { "FAIL" }
        if ($battStatus -eq "FAIL") { $hardFail = $true }
        $checks["Battery Health"] = [PSCustomObject]@{ Value = "$battPct%"; Status = $battStatus }
    } else {
        $checks["Battery Health"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "WARN" }
    }

    # ── Cycle count ──────────────────────────────────────────────────────
    $cycles = $null
    try {
        if (Test-LenovoNamespace) {
            $odo = Get-CimInstance -CimSession $script:CimSession -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            [int64]$c = 0
            if ([int64]::TryParse(([string]$odo.Battery_cycles -replace '\D',''),[ref]$c)) { $cycles = $c }
        }
    } catch {}

    if ($null -ne $cycles) {
        $cycleStatus = if ($cycles -lt 300) { "PASS" } elseif ($cycles -lt 500) { "WARN" } else { "FAIL" }
        if ($cycleStatus -eq "FAIL") { $hardFail = $true }
        $checks["Cycle Count"] = [PSCustomObject]@{ Value = "$cycles cycles"; Status = $cycleStatus }
    } else {
        $checks["Cycle Count"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "INFO" }
    }

    # ── Memory ───────────────────────────────────────────────────────────
    try {
        $mem = Get-MemoryState
        if ($mem.ModuleCount -gt 0) {
            $memStatus = if ($mem.SpeedMismatch) { "WARN" } else { "PASS" }
            $checks["Memory"] = [PSCustomObject]@{ Value = "$($mem.TotalRAMGB) GB  ($($mem.ModuleCount) module(s))"; Status = $memStatus }
        } else {
            $checks["Memory"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "INFO" }
        }
    } catch {
        $checks["Memory"] = [PSCustomObject]@{ Value = "Unavailable"; Status = "INFO" }
    }

    # ── BIOS ─────────────────────────────────────────────────────────────
    $si = Get-SystemInfo
    $biosStatus = if ($si.BIOSVersion -ne "Unknown") { "PASS" } else { "WARN" }
    $checks["BIOS"] = [PSCustomObject]@{ Value = "$($si.BIOSVersion)  ($($si.BIOSDate))"; Status = $biosStatus }

    # ── CDRT shock/thermal (if available) ────────────────────────────────
    $cdrt = Get-CdrtOdometerData
    if ($cdrt.Available) {
        $shockStatus = if     ($null -eq $cdrt.ShockEvents)   { "INFO" }
                       elseif ($cdrt.ShockEvents -le 20000)   { "PASS" }
                       elseif ($cdrt.ShockEvents -le 60000)   { "WARN" }
                       else                                    { "WARN" }
        $checks["Shock Events"] = [PSCustomObject]@{
            Value  = if ($null -ne $cdrt.ShockEvents) { "$($cdrt.ShockEvents)" } else { "N/A" }
            Status = $shockStatus
        }

        $thermalStatus = if     ($null -eq $cdrt.ThermalEvents)  { "INFO" }
                         elseif ($cdrt.ThermalEvents -le 50)     { "PASS" }
                         elseif ($cdrt.ThermalEvents -le 200)    { "WARN" }
                         else                                     { "WARN" }
        $checks["Thermal Events"] = [PSCustomObject]@{
            Value  = if ($null -ne $cdrt.ThermalEvents) { "$($cdrt.ThermalEvents)" } else { "N/A" }
            Status = $thermalStatus
        }
    } else {
        $checks["Shock / Thermal"] = [PSCustomObject]@{ Value = "Vantage not installed — skipped"; Status = "INFO" }
    }

    # ── Overall verdict ──────────────────────────────────────────────────
    $anyWarn = $checks.Values | Where-Object { $_.Status -eq "WARN" }
    $status  = if ($hardFail)        { "NOT READY"     }
               elseif ($anyWarn)     { "REVIEW NEEDED" }
               else                  { "READY"         }
    $statusColor = if ($hardFail) { "Red" } elseif ($anyWarn) { "Yellow" } else { "Green" }

    return [PSCustomObject]@{
        Ready       = (-not $hardFail)
        Status      = $status
        StatusColor = $statusColor
        Checks      = $checks
    }
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
            $result.Message = "No physical memory modules found."
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
                $result.Message = "Warning: Memory modules have mismatched speeds. This may impact performance."
            }
            else {
                $result.Message = "All memory modules have matching speeds."
            }
        }
        elseif ($speeds.Count -eq 1) {
            $result.Message = "Single memory module detected."
        }
        else {
            $result.Message = "Unable to determine memory speeds."
        }
        
        return $result
    }
    catch {
        Log-Error "root\cimv2" "Win32_PhysicalMemory" $_
        $result.Message = "Failed to retrieve memory information."
        return $result
    }
}

function MemoryInfo {
    Show-Header
    Write-Host "[ Memory Information - Win32_PhysicalMemory ]"
    Write-Host ""

    $memoryState = Get-MemoryState

    if (-not [string]::IsNullOrEmpty($memoryState.Message) -and $memoryState.ModuleCount -eq 0) {
        $failColor = if ($memoryState.Message -match "Failed") { "Red" } else { "Yellow" }
        Write-Host $memoryState.Message -ForegroundColor $failColor
    }
    else {
        Write-Host "Total RAM              : $($memoryState.TotalRAMGB) GB"
        Write-Host "Module Count           : $($memoryState.ModuleCount)"
        Write-Host ""

        # Determine color based on speed mismatch
        $statusColor = if ($memoryState.SpeedMismatch) { "Yellow" } else { "Green" }
        Write-Host "Speed Mismatch         : " -NoNewline
        Write-Host $(if ($memoryState.SpeedMismatch) { "Yes" } else { "No" }) -ForegroundColor $statusColor

        if ($memoryState.SpeedMismatch) {
            Write-Host "Severity               : " -NoNewline
            Write-Host "Warning ($($memoryState.Severity)/1)" -ForegroundColor Yellow
        }

        Write-Host ""
        Write-Host "[ Memory Module Details ]"
        Write-Host ""

        $moduleIndex = 0
        foreach ($module in $memoryState.Modules) {
            Write-Host "Module #: $moduleIndex"
            Write-Host ("-" * 50)
            Write-Host "Location        : $($module.DeviceLocator)"
            Write-Host "Capacity        : $($module.CapacityGB) GB"
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
$welcomeFlag = "$env:TEMP\ThinkPadUtil_Notice.flag"
if (-not (Test-Path $welcomeFlag)) {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  Welcome to ThinkPad Utility Management" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This tool surfaces hardware and battery information that Windows"
    Write-Host "doesn't easily show — pulled directly from your ThinkPad's firmware"
    Write-Host "and Lenovo EC (Embedded Controller) WMI interface."
    Write-Host ""
    Write-Host "Most importantly: check battery health early." -ForegroundColor Yellow
    Write-Host "Degraded batteries cause short runtime, sudden shutdowns,"
    Write-Host "and in severe cases, swelling or safety risk."
    Write-Host ""
    Write-Host "→ If your ThinkPad feels like it dies too fast, start with options 1–5." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Press ENTER to continue..."
    Read-Host | Out-Null
    "seen" | Out-File $welcomeFlag -Encoding ascii
}


function Show-ErrorLog {
    Show-Header
    Write-Host "[ Error Log Viewer ]"
    Write-Host ""

    if ($script:ErrorLog.Count -eq 0) {
        Write-Host "No errors logged this session." -ForegroundColor Green
        Write-Host ""
        Read-Host "Press ENTER"
        return
    }

    Write-Host "Total errors this session: $($script:ErrorLog.Count)" -ForegroundColor Red
    Write-Host ""

    $index = 0
    foreach ($err in $script:ErrorLog) {
        $index++
        Write-Host ("=" * 50) -ForegroundColor DarkGray
        Write-Host "Error #$index" -ForegroundColor Red
        Write-Host ("=" * 50) -ForegroundColor DarkGray

        Write-Host "Time      : $($err.TimeStamp)"
        Write-Host "Namespace : $($err.Namespace)"
        Write-Host "Class     : $($err.Class)"

        Write-Host "Message   : " -NoNewline
        Write-Host $err.Message -ForegroundColor Yellow

        if ($err.Context) {
            Write-Host "Context   : $($err.Context)"
        }

        if ($err.ScriptName -or $err.ScriptLineNumber) {
            Write-Host "Location  : $($err.ScriptName) line $($err.ScriptLineNumber)" -ForegroundColor DarkGray
        }

        if ($err.BatteryDetails) {
            Write-Host "Battery   :"
            Write-Host "  InstanceName        : $($err.BatteryDetails.InstanceName)" -ForegroundColor DarkGray
            Write-Host "  FullChargedCapacity : $($err.BatteryDetails.FullChargedCapacity)" -ForegroundColor DarkGray
            Write-Host "  DesignedCapacity    : $($err.BatteryDetails.DesignedCapacity)" -ForegroundColor DarkGray
        }

        if ($err.StackTrace) {
            Write-Host "Stack     :" -ForegroundColor DarkGray
            # Print each stack frame on its own line, indented
            $err.StackTrace -split "`n" | ForEach-Object {
                Write-Host "  $($_.Trim())" -ForegroundColor DarkGray
            }
        }

        Write-Host ""
    }

    Read-Host "Press ENTER"
}


function Get-ThinkPadHealthScore {
    <#
    .SYNOPSIS
    Computes a single 0-100 health score for this ThinkPad by aggregating
    battery health, cycle count, shock events, thermal events, and memory
    configuration into weighted component scores.

    .DESCRIPTION
    Component weights (must sum to 100):
      Battery Health %  50 pts  -- primary wear indicator
      Cycle Count       20 pts  -- corroborates battery age
      Shock Events      15 pts  -- physical abuse history (CDRT, optional)
      Thermal Events    10 pts  -- sustained overheating history (CDRT, optional)
      Memory Mismatch    5 pts  -- configuration defect

    When CDRT data is unavailable the 25 pts allocated to shock and thermal
    are redistributed: 15 to battery health and 10 to cycle count, keeping
    the total at 100 and the score meaningful on systems without Commercial
    Vantage.

    Returns a PSCustomObject with:
      Score           - Integer 0-100
      Grade           - Excellent / Good / Fair / Poor / Critical
      GradeColor      - PowerShell console colour string
      Components      - Hashtable of per-component scores and notes
      CdrtAvailable   - Whether CDRT data contributed to the score
      DataSources     - Array of human-readable source descriptions used
    #>

    $score      = 0
    $components = [ordered]@{}
    $sources    = @()
    $cdrtAvail  = $false

    # ── Component 1: Battery Health % ───────────────────────────────────────
    # Try Lenovo_Battery first, fall back to ACPI BatteryFullChargedCapacity.
    # Weight: 50 pts normally; 65 pts when CDRT unavailable.
    $batteryPct   = $null
    $batteryNotes = "No battery data found"

    if ((Test-LenovoNamespace) -and (Test-WmiClass -Namespace "root\Lenovo" -ClassName "Lenovo_Battery")) {
        try {
            $lbBatteries = @(Get-CimInstance -CimSession $script:CimSession `
                -Namespace root\Lenovo -ClassName Lenovo_Battery -ErrorAction Stop)
            if ($lbBatteries -and $lbBatteries.Count -gt 0) {
                $healths = @()
                foreach ($lb in $lbBatteries) {
                    $dcStr = ((Get-SafeWmiProperty -Object $lb -PropertyName "DesignCapacity") `
                        -replace '[^0-9,\.]','') -replace ',','.'
                    $fcStr = ((Get-SafeWmiProperty -Object $lb -PropertyName "FullChargeCapacity") `
                        -replace '[^0-9,\.]','') -replace ',','.'
                    [double]$dc = 0; [double]$fc = 0
                    if ([double]::TryParse($dcStr, [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dc) -and
                        [double]::TryParse($fcStr, [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fc) -and
                        $dc -gt 0) {
                        $healths += [math]::Round(($fc / $dc) * 100, 1)
                    }
                }
                if ($healths.Count -gt 0) {
                    # Use worst battery — a chain is only as strong as its weakest link
                    $batteryPct   = ($healths | Measure-Object -Minimum).Minimum
                    $batteryNotes = "Worst battery: $batteryPct% (Lenovo_Battery)"
                    $sources     += "Lenovo_Battery (root\Lenovo)"
                }
            }
        } catch {}
    }

    if ($null -eq $batteryPct) {
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
                    $batteryNotes = "Worst battery: $batteryPct% (ACPI)"
                    $sources     += "BatteryFullChargedCapacity (root\wmi)"
                }
            }
        } catch {}
    }

    # ── Component 2: Cycle Count ─────────────────────────────────────────────
    # Source: Lenovo_Odometer.Battery_cycles.
    # Weight: 20 pts normally; 30 pts when CDRT unavailable.
    # Typical ThinkPad battery is rated 300-500 cycles; 80% health often
    # reached around 300. We score linearly: 0 cycles = full marks, 500+ = 0.
    $cycleCount  = $null
    $cycleNotes  = "Cycle count unavailable"

    if (Test-LenovoNamespace) {
        try {
            $odo = Get-CimInstance -CimSession $script:CimSession `
                -Namespace root\Lenovo -ClassName Lenovo_Odometer -ErrorAction Stop
            [int64]$parsed = 0
            if ([int64]::TryParse(([string]$odo.Battery_cycles -replace '\D',''), [ref]$parsed)) {
                $cycleCount  = $parsed
                $cycleNotes  = "$cycleCount cycles"
                $sources    += "Lenovo_Odometer (root\Lenovo)"
            }
        } catch {}
    }

    # ── Component 3 & 4: CDRT Shock and Thermal Events ──────────────────────
    $cdrt = Get-CdrtOdometerData
    if ($cdrt.Available) {
        $cdrtAvail = $true
        $sources  += "CDRT Odometer via Lenovo_Odometer"
    }

    # ── Calculate component scores ───────────────────────────────────────────
    # Redistribute weights when CDRT is absent.
    $wBattery = if ($cdrtAvail) { 50 } else { 65 }
    $wCycle   = if ($cdrtAvail) { 20 } else { 30 }
    $wShock   = if ($cdrtAvail) { 15 } else {  0 }
    $wThermal = if ($cdrtAvail) { 10 } else {  0 }
    $wMemory  = 5

    # Battery component (0–wBattery pts)
    $batteryScore = 0
    if ($null -ne $batteryPct) {
        # Direct linear mapping: health% maps to full weight
        $batteryScore = [math]::Round(($batteryPct / 100) * $wBattery, 1)
    }
    else {
        # No data — assume neutral (70%) rather than penalising unfairly
        $batteryScore  = [math]::Round(0.70 * $wBattery, 1)
        $batteryNotes  = "No battery data - assumed neutral"
    }
    $components["Battery Health"] = [PSCustomObject]@{
        Score   = $batteryScore
        MaxScore = $wBattery
        Note    = $batteryNotes
    }

    # Cycle component (0–wCycle pts)
    # 0 cycles = full marks; 500 cycles = 0 pts; linear between.
    $cycleScore = 0
    if ($null -ne $cycleCount) {
        $cycleFraction = [math]::Max(0, (500 - $cycleCount) / 500)
        $cycleScore    = [math]::Round($cycleFraction * $wCycle, 1)
    }
    else {
        $cycleScore = [math]::Round(0.70 * $wCycle, 1)
        $cycleNotes = "Cycle count unavailable - assumed neutral"
    }
    $components["Cycle Count"] = [PSCustomObject]@{
        Score    = $cycleScore
        MaxScore = $wCycle
        Note     = $cycleNotes
    }

    # Shock component (0–wShock pts) — only when CDRT available
    $shockScore = 0
    $shockNotes = "CDRT not available"
    if ($cdrtAvail -and $null -ne $cdrt.ShockEvents) {
        # 0-20,000 = full marks; 20,001-60,000 = linear decay to 50%;
        # 60,000+ = linear decay from 50% to 0 at 120,000
        $s = $cdrt.ShockEvents
        $shockFraction = if     ($s -le 20000) { 1.0 }
                         elseif ($s -le 60000) { 1.0 - (($s - 20000) / 80000) }
                         else                  { [math]::Max(0, 1.0 - (($s - 20000) / 100000)) }
        $shockScore = [math]::Round($shockFraction * $wShock, 1)
        $shockNotes = "$($cdrt.ShockEvents) shock events"
    }
    elseif ($cdrtAvail) {
        $shockNotes = "Shock events field not found in CDRT"
    }
    if ($wShock -gt 0) {
        $components["Shock Events"] = [PSCustomObject]@{
            Score    = $shockScore
            MaxScore = $wShock
            Note     = $shockNotes
        }
    }

    # Thermal component (0–wThermal pts) — only when CDRT available
    $thermalScore = 0
    $thermalNotes = "CDRT not available"
    if ($cdrtAvail -and $null -ne $cdrt.ThermalEvents) {
        # 0-50 = full marks; 51-200 = linear decay to 50%; 200+ = decay to 0 at 400
        $t = $cdrt.ThermalEvents
        $thermalFraction = if     ($t -le 50)  { 1.0 }
                           elseif ($t -le 200) { 1.0 - (($t - 50) / 300) }
                           else                { [math]::Max(0, 1.0 - (($t - 50) / 700)) }
        $thermalScore = [math]::Round($thermalFraction * $wThermal, 1)
        $thermalNotes = "$($cdrt.ThermalEvents) thermal throttle events"
    }
    elseif ($cdrtAvail) {
        $thermalNotes = "Thermal events field not found in CDRT"
    }
    if ($wThermal -gt 0) {
        $components["Thermal Events"] = [PSCustomObject]@{
            Score    = $thermalScore
            MaxScore = $wThermal
            Note     = $thermalNotes
        }
    }

    # Memory component (0–5 pts)
    $memScore = 0
    $memNotes = "Memory data unavailable"
    try {
        $memState = Get-MemoryState
        if ($memState.ModuleCount -gt 0) {
            $sources += "Win32_PhysicalMemory"
            if ($memState.SpeedMismatch) {
                $memScore = 0
                $memNotes = "Speed mismatch detected ($($memState.ModuleCount) modules)"
            }
            else {
                $memScore = $wMemory
                $memNotes = "$($memState.ModuleCount) module(s) matched"
            }
        }
        else {
            # Single soldered module or unreadable — give full marks, no penalty
            $memScore = $wMemory
            $memNotes = "Single/soldered module or unreadable - no penalty"
        }
    } catch {}
    $components["Memory"] = [PSCustomObject]@{
        Score    = $memScore
        MaxScore = $wMemory
        Note     = $memNotes
    }

    # ── Total ────────────────────────────────────────────────────────────────
    $total = [math]::Round($batteryScore + $cycleScore + $shockScore + $thermalScore + $memScore)
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
        Score         = $total
        Grade         = $grade
        GradeColor    = $gradeColor
        Components    = $components
        CdrtAvailable = $cdrtAvail
        DataSources   = $sources
    }
}

function Show-ThinkPadHealthScore {
    Show-Header
    Write-Host "[ ThinkPad Health Score ]"
    Write-Host ""
    Write-Host "Calculating..." -ForegroundColor Cyan
    Write-Host ""

    $hs = Get-ThinkPadHealthScore

    # ── Score banner ─────────────────────────────────────────────────────────
    Write-Host ("=" * 50)
    Write-Host "  Health Score : " -NoNewline
    Write-Host "$($hs.Score) / 100" -ForegroundColor $hs.GradeColor -NoNewline
    Write-Host "   [$($hs.Grade)]" -ForegroundColor $hs.GradeColor
    Write-Host ("=" * 50)
    Write-Host ""

    # ── Component breakdown ──────────────────────────────────────────────────
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
        $label     = "$key".PadRight(20)
        Write-Host "  $label : " -NoNewline
        Write-Host "$($c.Score) / $($c.MaxScore) pts" -ForegroundColor $compColor -NoNewline
        Write-Host "   $($c.Note)" -ForegroundColor DarkGray
    }
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    Write-Host ""

    # ── CDRT notice if absent ────────────────────────────────────────────────
    if (-not $hs.CdrtAvailable) {
        Write-Host "Note: Shock and thermal event data unavailable (CDRT/Commercial Vantage" -ForegroundColor DarkGray
        Write-Host "      not deployed). Their weight was redistributed to battery and cycle." -ForegroundColor DarkGray
        Write-Host ""
    }

    # ── Data sources ─────────────────────────────────────────────────────────
    Write-Host "[ Data Sources ]" -ForegroundColor DarkGray
    foreach ($src in $hs.DataSources) {
        Write-Host "  - $src" -ForegroundColor DarkGray
    }
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
    # The catalog <location> element points to a per-package XML with full details.
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

    # Extract version and date from package descriptor
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
    # Win32_BIOS.SMBIOSBIOSVersion returns a string like 'R1KET50W (1.35 )'.
    # The catalog returns only the numeric part e.g. '1.35'.
    # Direct string comparison always fails because the formats differ.
    #
    # Strategy: extract the numeric version from both sides.
    #   Installed: pull the decimal number from inside parentheses if present,
    #              otherwise fall back to any decimal number in the string.
    #   Catalog:   already a plain decimal string, just trim it.
    # Compare as [version] objects so 1.10 > 1.9 correctly.
    # If either side cannot be parsed as a version, fall back to string compare.

    $installedRaw = $result.InstalledVersion.Trim()
    $latestRaw    = $result.LatestVersion.Trim()

    # Extract numeric version from installed string
    # Try parentheses first: 'R1KET50W (1.35 )' -> '1.35'
    $installedNumeric = $null
    if ($installedRaw -match '\(([\d\.]+)') {
        $installedNumeric = $matches[1].Trim()
    }
    elseif ($installedRaw -match '([\d]+\.[\d]+)') {
        # Fallback: any decimal number in the string
        $installedNumeric = $matches[1].Trim()
    }

    # Extract numeric version from catalog string (usually already clean)
    $latestNumeric = $null
    if ($latestRaw -match '([\d]+\.[\d]+)') {
        $latestNumeric = $matches[1].Trim()
    }

    # Attempt [version] comparison
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
        Write-Host "  Could not retrieve catalog data." -ForegroundColor Yellow
        Write-Host "  Reason: $($bios.UnavailableReason)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Ensure the system has internet access and try again." -ForegroundColor DarkGray
        Write-Host "  Or check manually: support.lenovo.com -> Drivers & Software" -ForegroundColor DarkGray
    }
    else {
        Write-Host "[ Latest - Lenovo Catalog ]"
        Write-Host "  Version      : $($bios.LatestVersion)"
        Write-Host "  Release Date : $($bios.LatestDate)"
        Write-Host "  MTM          : $($bios.MTM)" -ForegroundColor DarkGray
        Write-Host "  Catalog      : $($bios.CatalogUrl)" -ForegroundColor DarkGray
        Write-Host ""

        Write-Host "[ Status ]"
        if ($bios.IsUpToDate) {
            Write-Host "  Up to date." -ForegroundColor Green
        }
        else {
            Write-Host "  Update available." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Installed : $($bios.InstalledVersion)" -ForegroundColor DarkGray
            Write-Host "  Available : $($bios.LatestVersion)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Download from: support.lenovo.com -> Drivers & Software" -ForegroundColor Cyan
            Write-Host "  Search for your model and filter by BIOS/UEFI." -ForegroundColor DarkGray
        }
    }

    Write-Host ""
    Write-Host "Note: Version comparison uses numeric extraction where possible." -ForegroundColor DarkGray
    Write-Host "      If the format is unrecognised, check the dates manually." -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press ENTER"
}

function Show-DeploymentReadiness {
    Show-Header
    Write-Host "[ DEPLOYMENT READINESS CHECK ]"
    Write-Host ""
    Write-Host "Evaluating device against deployment criteria..." -ForegroundColor Cyan
    Write-Host ""

    $dr = Get-DeploymentReadiness

    Write-Host ("-" * 50) -ForegroundColor DarkGray
    foreach ($key in $dr.Checks.Keys) {
        $chk = $dr.Checks[$key]
        Format-PassFail -Label $key -Value $chk.Value -Status $chk.Status
    }
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    Write-Host ""

    Write-Host "  Deployment Status : " -NoNewline
    Write-Host $dr.Status -ForegroundColor $dr.StatusColor

    Write-Host ""
    Write-Host "Criteria:" -ForegroundColor DarkGray
    Write-Host "  PASS  Battery >= 80%,  Cycles < 300,  Memory matched" -ForegroundColor DarkGray
    Write-Host "  WARN  Battery 70–79%,  Cycles 300–499,  Memory mismatch" -ForegroundColor DarkGray
    Write-Host "  FAIL  Battery < 70%,   Cycles >= 500" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press ENTER"
}

do {
    Show-Header

    # ── Guided-mode auto-prompt (#5) ─────────────────────────────────────
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
        Write-Host "  Option 5 (Comprehensive Battery Analysis) now?" -ForegroundColor $guideColor
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

    Write-Host "1. Battery Static Data"
    Write-Host "2. Lenovo Battery Cycles"
    Write-Host "3. Warranty Info"
    Write-Host -NoNewline "4. Full Charge Capacity"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host -NoNewline "5. Comprehensive Battery Analysis"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "6. Memory Info"
    Write-Host "7. Export Report"
    Write-Host "8. Diagnostic Info"
    Write-Host -NoNewline "9. Battery Age Estimation"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "10. Battery Charge Threshold"
    Write-Host "11. CDRT Odometer"
    Write-Host "12. About"
    Write-Host "13. Error Log Viewer"
    Write-Host "14. ThinkPad Health Score"
    Write-Host "15. BIOS Update Check"
    Write-Host -NoNewline "16. Battery Health Trend"
    if ($alertSuffix) { Write-Host $alertSuffix -ForegroundColor $alertColor } else { Write-Host "" }
    Write-Host "17. Deployment Readiness"
    Write-Host "18. Exit"
    Write-Host ""

    $choice = Read-Host "Select option"

    switch ($choice) {
        "1"  { BatteryStaticData }
        "2"  { LenovoEC }
        "3"  { WarrantyInfo }
        "4"  { FullCharge }
        "5"  { ComprehensiveBatteryAnalysis }
        "6"  { MemoryInfo }
        "7"  { ExportReport }
        "8"  { DiagnosticInfo }
        "9"  { BatteryAgeEstimation }
        "10" { ChargeThresholdManager }
        "11" { Show-CdrtOdometer }
        "12" { Show-About }
        "13" { Show-ErrorLog }
        "14" { Show-ThinkPadHealthScore }
        "15" { Show-BiosUpdateCheck }
        "16" { Show-BatteryHealthTrend }
        "17" { Show-DeploymentReadiness }
    }

} while ($choice -ne "18")

# Tear down the shared CIM session cleanly before exiting.
if ($script:CimSession) {
    Remove-CimSession -CimSession $script:CimSession -ErrorAction SilentlyContinue
    $script:CimSession = $null
}

Clear-Host
Write-Host "Exiting..."
