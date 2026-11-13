# test-compile.ps1
# Compile-test harness for Pending_Pickup_Macro.bas
# Copies the template, injects a tiny probe macro, then runs it via COM.
# Running ANY macro forces VBA to compile the ENTIRE project, so a compile
# error anywhere blocks the probe. A dialog watcher (watch-compile-error.ps1)
# runs in a separate process: it logs any dialog text to compile-error.log
# and auto-dismisses error dialogs so this script never hangs.
# A watchdog kills the harness's own Excel instance after 120s if needed.

param(
    [string]$Template = "$PSScriptRoot\Pending Pickup Template.xlsm"
)

$ErrorActionPreference = "Stop"
$testCopy  = Join-Path $PSScriptRoot "_compile_test.xlsm"
$marker    = Join-Path $PSScriptRoot "probe_ran.txt"
$watcher   = Join-Path $PSScriptRoot "watch-compile-error.ps1"
$watchLog  = Join-Path $PSScriptRoot "compile-error.log"
$testLog   = Join-Path $PSScriptRoot "compile-test.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
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
$watchProc = Start-Process powershell -WindowStyle Hidden -PassThru -ArgumentList @(
    '-ExecutionPolicy','Bypass','-File', $watcher)
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

# 3. Watchdog: force-kill our Excel pid after 120s (guaranteed no hang)
$watchdog = $null
if ($excelPid) {
    $watchdog = Start-Job -ScriptBlock {
        param($tp)
        Start-Sleep -Seconds 120
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
        Log "Captured dialog text:"
        foreach ($l in $newLines) { Log "  $l" }
    } else {
        Log "No dialog captured. Run error: $runError"
    }
}

# 6. Cleanup: close workbook, quit/kill our Excel, stop watcher and watchdog
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