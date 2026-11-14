# run-report.ps1
# Runs the report macro end-to-end against a temp copy of the template.
# Hard rules: Excel must be killed if the run stalls (no new report file within
# $MaxRunSeconds, default 25s) or if a modal dialog stays unanswered too long.
# The watchdog + dialog answerer run on separate runspace threads so they work
# even while the main thread is blocked inside the synchronous Excel.Run call.
# Logs: log\run-report.log, log\answer-dialogs.log
param(
    [string]$Template = "C:\Users\smd32\Downloads\Pending Pickup Project\Pending Pickup Template.xlsm",
    [string]$SourcePath = "C:\Users\smd32\Downloads\Pending Pickup Project\AllOpenTreatmentPlansResults618.xlsx",
    [string]$TempCopy = "C:\Users\smd32\AppData\Local\Temp\kilo\_run_temp.xlsm",
    [int]$MaxRunSeconds = 30,
    [int]$StallSeconds = 10
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class W {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, string lp);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
    delegate bool EnumProc(IntPtr h, IntPtr l);
    static string T(IntPtr h) { var sb = new StringBuilder(1024); GetWindowText(h, sb, 1024); return sb.ToString(); }
    static string C(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static List<IntPtr> ExcelWindows() {
        var list = new List<IntPtr>();
        var pids = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName("EXCEL")) pids.Add(p.Id);
        if (pids.Count == 0) return list;
        EnumWindows(delegate(IntPtr h, IntPtr l) {
            if (!IsWindowVisible(h)) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains((int)pid)) return true;
            list.Add(h);
            return true;
        }, IntPtr.Zero);
        return list;
    }
    public static IntPtr FindEdit(IntPtr h) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(h, delegate(IntPtr c, IntPtr l) {
            if (C(c) == "Edit") { found = c; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
    public static string Text(IntPtr h) { return T(h); }
    public static string Class(IntPtr h) { return C(h); }
    public static List<string> ChildTexts(IntPtr parent) {
        var texts = new List<string>();
        EnumChildWindows(parent, delegate(IntPtr c, IntPtr l) {
            var t = T(c);
            if (t.Length > 0) texts.Add(C(c) + ":" + t);
            return true;
        }, IntPtr.Zero);
        return texts;
    }
    public static void Focus(IntPtr h) { SetForegroundWindow(h); }
    public static void Type(IntPtr h, string s) { SendMessage(h, 0x000C, IntPtr.Zero, s); }
    public static void Enter(IntPtr h) { SendMessage(h, 0x0111, (IntPtr)1, IntPtr.Zero); }
    public static void Esc(IntPtr h) { SendMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); }
    public static void KillExcel() {
        foreach (var p in Process.GetProcessesByName("EXCEL")) { try { p.Kill(); } catch {} }
    }
}
"@
$projDir = "C:\Users\smd32\Downloads\Pending Pickup Project"
$reportsDir = Join-Path $projDir "Reports"
if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir | Out-Null }
$before = @(Get-ChildItem -LiteralPath $reportsDir -Filter *.xlsx | Select-Object -ExpandProperty Name)
$ansLog = "$projDir\log\answer-dialogs.log"
$runLog = "$projDir\log\run-report.log"
function Write-Log { param($p, $m) "$(Get-Date -Format 'HH:mm:ss') | $m" | Add-Content -LiteralPath $p }

Copy-Item -LiteralPath $Template -Destination $TempCopy -Force
$outcome = "UNKNOWN"
$excelPid = $null
$runToken = [System.Threading.CancellationTokenSource]::new()

$mainPool = [runspacefactory]::CreateRunspacePool(1, 8)
$mainPool.Open()
$psThreads = @()

# Thread A: watchdog. Polls for new report files and for stall; kills Excel.
$psA = [powershell]::Create()
$psA.RunspacePool = $mainPool
$psA.AddScript({
    param($reportsDir, $before, $runLog, $ansLog, $maxSec, $stallSec, $token)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $lastProgress = 0
    $lastDialogs = 0
    $lastReportCheck = ""
    while (-not $token.IsCancellationRequested -and $sw.Elapsed.TotalSeconds -lt $maxSec) {
        Start-Sleep -Milliseconds 300
        # New report file = success.
        $newFiles = @(Get-ChildItem -LiteralPath $reportsDir -Filter *.xlsx -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Name })
        if ($newFiles.Count -gt 0) { "$(Get-Date -Format 'HH:mm:ss') | REPORT OK: $($newFiles.Name -join ', ')" | Add-Content -LiteralPath $runLog; return "OK" }
        # Answerable dialog visible? (Only real dialogs reset the stall timer;
        # the Excel main window itself does not.)
        $dialogOpen = $false
        foreach ($h in [W]::ExcelWindows()) {
            $cls = [W]::Class($h)
            $title = [W]::Text($h)
            if ($cls -eq "bosa_sdm_MsoFiles" -or $cls -eq "#32770" -or $title -match "Select the CURRENT|Select the PRIOR|Choose where|Microsoft Excel") { $dialogOpen = $true }
        }
        if ($dialogOpen) { $lastDialogs = $sw.Elapsed.TotalSeconds; continue }
        # Stall check: no dialog for $stallSec consecutive seconds while Excel busy.
        $excelBusy = $false
        try { $excelBusy = (Get-Process EXCEL -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0 } catch {}
        if ($excelBusy -and ($sw.Elapsed.TotalSeconds - $lastDialogs) -gt $stallSec) {
            "$(Get-Date -Format 'HH:mm:ss') | KILLED: Excel busy > $stallSec s with no answerable dialog (elapsed=$([int]$sw.Elapsed.TotalSeconds)s)" | Add-Content -LiteralPath $runLog
            # Capture what windows exist at kill time.
            foreach ($h in [W]::ExcelWindows()) { "  window: class=$([W]::Class($h)) title=$([W]::Text($h))" | Add-Content -LiteralPath $runLog }
            [W]::KillExcel()
            return "KILLED"
        }
    }
    "$(Get-Date -Format 'HH:mm:ss') | KILLED: run exceeded $maxSec s without a new report file" | Add-Content -LiteralPath $runLog
    [W]::KillExcel()
    return "TIMEOUT"
}).AddArgument($reportsDir).AddArgument($before).AddArgument($runLog).AddArgument($ansLog).AddArgument($MaxRunSeconds).AddArgument($StallSeconds).AddArgument($runToken.Token) | Out-Null
$psA.BeginInvoke() | Out-Null

# Thread B: dialog answerer.
$psB = [powershell]::Create()
$psB.RunspacePool = $mainPool
$psB.AddScript({
    param($sourcePath, $ansLog, $shotPath, $token)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $answered = @{}
    while (-not $token.IsCancellationRequested -and $sw.Elapsed.TotalSeconds -lt 120) {
        Start-Sleep -Milliseconds 250
        $live = @{}
        foreach ($h in [W]::ExcelWindows()) { $live["$h"] = $h }
        foreach ($k in @($answered.Keys)) { if (-not $live.ContainsKey($k)) { [void]$answered.Remove($k) } }
        foreach ($k in $live.Keys) {
            if ($answered.ContainsKey($k)) { continue }
            $h = $live[$k]
            $cls = [W]::Class($h)
            $title = [W]::Text($h)
            "$(Get-Date -Format 'HH:mm:ss') | dialog seen: class=$cls title=$title" | Add-Content -LiteralPath $ansLog
            if ($cls -eq "#32770" -and $title -eq "Microsoft Excel") {
                # Log all child windows to find the error message text. Do NOT dismiss — leave it open for inspection.
                $children = [W]::ChildTexts($h)
                "$(Get-Date -Format 'HH:mm:ss') | ERROR DIALOG children: $($children -join ' | ')" | Add-Content -LiteralPath $ansLog
                $answered[$k] = 1
                continue
            }
            if ($cls -eq "bosa_sdm_MsoFiles" -or $cls -eq "#32770" -or $title -match "Select the CURRENT|Select the PRIOR|Choose where") {
                [W]::Focus($h) | Out-Null
                Start-Sleep -Milliseconds 300
                if ($title -match 'PRIOR') {
                    [W]::Esc($h) | Out-Null
                    "$(Get-Date -Format 'HH:mm:ss') | cancelled PRIOR picker" | Add-Content -LiteralPath $ansLog
                } elseif ($title -match 'CURRENT') {
                    $edit = [W]::FindEdit($h)
                    if ($edit -ne [IntPtr]::Zero) {
                        [W]::Type($edit, $sourcePath) | Out-Null
                        Start-Sleep -Milliseconds 300
                        [W]::Enter($h) | Out-Null
                    } else {
                        [System.Windows.Forms.SendKeys]::SendWait("^a")
                        [System.Windows.Forms.SendKeys]::SendWait($sourcePath)
                        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    }
                    "$(Get-Date -Format 'HH:mm:ss') | answered CURRENT picker -> $sourcePath" | Add-Content -LiteralPath $ansLog
                } else {
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    "$(Get-Date -Format 'HH:mm:ss') | dismissed dialog: $title" | Add-Content -LiteralPath $ansLog
                }
                $answered[$k] = 1
            }
        }
    }
}).AddArgument($SourcePath).AddArgument($ansLog).AddArgument("$projDir\log\screen.png").AddArgument($runToken.Token) | Out-Null
$psB.BeginInvoke() | Out-Null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($TempCopy)
    $excel.Run("CreatePendingPickupReport")
} catch {
    $outcome = "ERROR: $($_.Exception.Message)"
} finally {
    $runToken.Cancel()
    Start-Sleep -Milliseconds 500
    # Screenshot for diagnosis if it stalled (file not produced).
    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size)
        $bmp.Save("$projDir\log\screen.png")
        $g.Dispose(); $bmp.Dispose()
    } catch {}
    [W]::KillExcel()
    if ($excel) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
    Remove-Item -LiteralPath $TempCopy -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath "$projDir\~`$_run_temp.xlsm" -Force -ErrorAction SilentlyContinue
    Write-Log $runLog "run finished (outcome=$outcome)"
    # Print the macro debug log for diagnosis.
    $dbgLog = "$projDir\log\macro-debug.log"
    if (Test-Path $dbgLog) {
        Write-Log $runLog "--- macro-debug.log ---"
        Get-Content $dbgLog | ForEach-Object { Write-Log $runLog $_ }
    }
}
$mainPool.Close()