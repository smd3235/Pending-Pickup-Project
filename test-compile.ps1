# test-compile.ps1
# Compile-test harness for Pending_Pickup_Macro.bas
# Copies the template, injects a tiny probe macro, then runs it via COM.
# Running ANY macro forces VBA to compile the ENTIRE project, so a compile
# error anywhere blocks the probe.
#
# Safety:
#  - watch-compile-error.ps1 runs in a separate hidden process; it logs any
#    dialog text to log\watcher.log and auto-dismisses error dialogs.
#  - A 10-second watchdog force-kills this script's own Excel instance.
#    (The whole test normally finishes in a few seconds.)
#  - Excel is always killed/closed at the end.
#
# All output goes to log\compile-test.log with per-line timestamps.

param(
    [string]$Template = "$PSScriptRoot\Pending Pickup Template.xlsm"
)

$ErrorActionPreference = "Stop"
$logDir   = Join-Path $PSScriptRoot "log"
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$testCopy = Join-Path $PSScriptRoot "_compile_test.xlsm"
$marker   = Join-Path $PSScriptRoot "probe_ran.txt"
$watcher  = Join-Path $PSScriptRoot "watch-compile-error.ps1"
$watchLog = Join-Path $logDir "watcher.log"
$testLog  = Join-Path $logDir "compile-test.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -LiteralPath $testLog -Value $line
}

# 0. Fresh test copy
if (Test-Path -LiteralPath $testCopy) { Remove-Item -LiteralPath $testCopy -Force }
Copy-Item -LiteralPath $Template -Destination $testCopy -Force
Unblock-File -LiteralPath $testCopy
if (Test-Path -LiteralPath $marker) { Remove-Item -LiteralPath $marker -Force }
Log "=== Compile test started (template: $Template) ==="

# 1. Snapshot EXCEL pids, start dialog watcher
$beforePids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$watchArgs = "-ExecutionPolicy Bypass -File `"$watcher`""
$watchProc = Start-Process powershell -WindowStyle Hidden -PassThru -ArgumentList $watchArgs
Log "Dialog watcher started (pid $($watchProc.Id))."
Start-Sleep -Seconds 4

# 2. Start our own Excel COM instance
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$afterPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$newPids = @($afterPids | Where-Object { $beforePids -notcontains $_ })
$excelPid = if ($newPids.Count -gt 0) { $newPids[0] } else { $null }
Log "Excel COM instance pid(s): $(if ($newPids) { $newPids -join ',' } else { 'attached to existing instance' })"

# 3. 10-second watchdog: force-kill our Excel pid if anything takes too long
$watchdog = $null
if ($excelPid) {
    $watchdog = Start-Job -ScriptBlock {
        param($tp)
        Start-Sleep -Seconds 10
        if (Get-Process -Id $tp -ErrorAction SilentlyContinue) {
            Stop-Process -Id $tp -Force -ErrorAction SilentlyContinue
        }
    } -ArgumentList $excelPid
}

# 4. Run the probe
$logMark = 0
if (Test-Path -LiteralPath $watchLog) { $logMark = @(Get-Content -LiteralPath $watchLog).Count }
$runError = $null
$wb = $null
try {
    $wb = $excel.Workbooks.Open($testCopy)
    $vbp = $wb.VBProject
    $mod = $vbp.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
    $mod.Name = "zzCompileProbe"
    $probeCode = "Option Explicit" + "`r`n" +
                 "Public Sub Probe()" + "`r`n" +
                 "    Open `"$marker`" For Output As #1" + "`r`n" +
                 "    Print #1, `"ran`"" + "`r`n" +
                 "    Close #1" + "`r`n" +
                 "End Sub"
    $mod.CodeModule.AddFromString($probeCode)
    Log "Running Probe (forces full-project compile)..."
    $excel.Run("Probe")
    Log "Run call returned without COM error."
} catch {
    $runError = $_.Exception.Message
    Log "Run call raised: $runError"
}

# 5. Collect watcher log lines produced during this run
Start-Sleep -Seconds 2
$newLines = @()
if (Test-Path -LiteralPath $watchLog) {
    $all = @(Get-Content -LiteralPath $watchLog)
    if ($all.Count -gt $logMark) { $newLines = $all[$logMark..($all.Count - 1)] }
}
$compileOk = (Test-Path -LiteralPath $marker)
if ($compileOk) {
    Log "RESULT: COMPILE OK - probe executed, project compiled cleanly."
} else {
    Log "RESULT: COMPILE FAILED - probe did not execute."
    if ($newLines.Count -gt 0) {
        Log "Captured dialog text (log\watcher.log):"
        foreach ($l in $newLines) { Log "  $l" }
    } else {
        Log "No dialog captured. Run error: $runError"
    }
}

# 6. Cleanup: cancel watchdog, close workbook, quit/kill our Excel, stop watcher
if ($watchdog) { Stop-Job $watchdog -ErrorAction SilentlyContinue; Remove-Job $watchdog -Force -ErrorAction SilentlyContinue }
try { if ($wb) { $wb.Close($false) } } catch { Log "Close failed: $($_.Exception.Message)" }
try {
    if ($excelPid) {
        $excel.Quit()
        if (Get-Process -Id $excelPid -ErrorAction SilentlyContinue) {
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            Log "Killed lingering Excel pid $excelPid."
        }
    }
} catch { Log "Quit failed: $($_.Exception.Message)" }
try { Stop-Process -Id $watchProc.Id -Force -ErrorAction SilentlyContinue } catch {}
Remove-Item -LiteralPath $testCopy -Force -ErrorAction SilentlyContinue
Log "=== Compile test finished ==="