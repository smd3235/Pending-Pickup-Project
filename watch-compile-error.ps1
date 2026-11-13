# watch-compile-error.ps1
# Watches Excel/VBA modal dialogs (e.g. "Compile error: Variable not defined")
# and logs them to log\watcher.log with per-line timestamps.
# - Dismisses error dialogs (posts IDOK to the default button).
# - If an error dialog refuses to close, kills the owning EXCEL process.
#
# Usage:
#   1. Start this script:
#        powershell -ExecutionPolicy Bypass -File .\watch-compile-error.ps1
#   2. Run the macro in Excel. Dialogs are captured to log\watcher.log and closed.
#   3. Stop with Ctrl+C (auto-stops after 2h).

$ErrorActionPreference = "Stop"
$logDir = Join-Path $PSScriptRoot "log"
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logPath = Join-Path $logDir "watcher.log"

function Write-Log {
    param([string]$Text)
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $out = @()
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line.Trim().Length -gt 0) { $out += "$stamp | $line" }
    }
    if ($out.Count -gt 0) { Add-Content -LiteralPath $logPath -Value $out }
}

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

public static class ExcelDialogWatcher
{
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    delegate bool EnumProc(IntPtr h, IntPtr l);

    static Dictionary<IntPtr, int> stuck = new Dictionary<IntPtr, int>();

    static string WinText(IntPtr h) { var sb = new StringBuilder(1024); GetWindowText(h, sb, 1024); return sb.ToString(); }
    static string WinClass(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }

    static string DialogText(IntPtr h)
    {
        var sb = new StringBuilder();
        EnumChildWindows(h, delegate(IntPtr c, IntPtr l2)
        {
            var t = WinText(c);
            if (t.Length > 0) { sb.AppendLine(t); }
            return true;
        }, IntPtr.Zero);
        return WinText(h) + "|" + sb.ToString();
    }

    static HashSet<int> ExcelPids()
    {
        var set = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName("EXCEL")) set.Add(p.Id);
        return set;
    }

    // Poll once. Returns text of every NEW dialog seen (for logging).
    // Dismisses error dialogs; kills the owning EXCEL process if an error
    // dialog survives 5 consecutive polls (dismissal failed).
    // killedPids receives any pids killed this poll.
    public static List<string> Poll(HashSet<string> seen, out List<int> killedPids)
    {
        var result = new List<string>();
        var killed = new List<int>();
        var pids = ExcelPids();
        if (pids.Count > 0)
        {
            EnumWindows(delegate(IntPtr h, IntPtr l)
            {
                if (!IsWindowVisible(h)) return true;
                uint pid; GetWindowThreadProcessId(h, out pid);
                if (!pids.Contains((int)pid)) return true;
                if (WinClass(h) != "#32770") return true;
                var text = DialogText(h);
                bool isError = text.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0;
                if (!seen.Contains(text)) { seen.Add(text); result.Add(text); }
                if (!isError) return true;
                int c;
                stuck.TryGetValue(h, out c);
                stuck[h] = c + 1;
                if (stuck[h] >= 5)
                {
                    try { var proc = Process.GetProcessById((int)pid); proc.Kill(); killed.Add((int)pid); }
                    catch { }
                }
                else
                {
                    PostMessage(h, 0x0111, (IntPtr)1, IntPtr.Zero); // WM_COMMAND = IDOK
                }
                return true;
            }, IntPtr.Zero);
        }
        killedPids = killed;
        return result;
    }
}
"@

$seen = New-Object 'System.Collections.Generic.HashSet[string]'
Write-Log "watcher started (pid $PID)"
Write-Host "Watching Excel for dialogs -> $logPath"
Write-Host "Run the macro in Excel now. Stop with Ctrl+C."
$deadline = (Get-Date).AddHours(2)
try {
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 600
        $killed = New-Object 'System.Collections.Generic.List[int]'
        $hits = [ExcelDialogWatcher]::Poll($seen, [ref]$killed)
        foreach ($hit in $hits) {
            Write-Log "DIALOG: $hit"
            Write-Host "Captured: $hit"
        }
        foreach ($kp in $killed) {
            Write-Log "KILLED EXCEL pid $kp (dialog would not close)"
            Write-Host "KILLED EXCEL pid $kp (dialog would not close)"
        }
    }
    Write-Log "watcher stopped (2h limit reached)"
}
finally {
    Write-Log "watcher stopped"
}