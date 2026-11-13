# watch-compile-error.ps1
# Watches Excel/VBA modal dialogs (e.g. "Compile error: Variable not defined")
# and writes their text to compile-error.log so the error can be reviewed.
# Auto-dismisses error dialogs so the macro run does not hang.
#
# Usage:
#   1. Start this script in a separate PowerShell window:
#        powershell -ExecutionPolicy Bypass -File .\watch-compile-error.ps1
#   2. Run the macro in Excel (button or Alt+F8). When a dialog pops,
#      its text is captured to compile-error.log and the dialog is closed.
#   3. Stop the watcher with Ctrl+C.
#   4. Open compile-error.log (or tell Kilo to read it).

$ErrorActionPreference = "Stop"
$logPath = Join-Path $PSScriptRoot "compile-error.log"

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

    static string WinText(IntPtr h) { var sb = new StringBuilder(1024); GetWindowText(h, sb, 1024); return sb.ToString(); }
    static string WinClass(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }

    static HashSet<int> ExcelPids()
    {
        var set = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName("EXCEL")) set.Add(p.Id);
        return set;
    }

    // Returns text of every visible Excel dialog window not seen before.
    public static List<string> FindNewDialogs(HashSet<string> seen)
    {
        var result = new List<string>();
        var pids = ExcelPids();
        if (pids.Count == 0) return result;
        EnumWindows(delegate(IntPtr h, IntPtr l)
        {
            if (!IsWindowVisible(h)) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains((int)pid)) return true;
            if (WinClass(h) != "#32770") return true;
            var title = WinText(h);
            var sb = new StringBuilder();
            EnumChildWindows(h, delegate(IntPtr c, IntPtr l2)
            {
                var t = WinText(c);
                if (t.Length > 0) { sb.AppendLine(t); }
                return true;
            }, IntPtr.Zero);
            var key = title + "|" + sb.ToString();
            if (!seen.Contains(key)) { seen.Add(key); result.Add(key); }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    // Closes Excel dialogs that look like errors (posts IDOK to the default button).
    public static void DismissErrorDialogs()
    {
        var pids = ExcelPids();
        if (pids.Count == 0) return;
        EnumWindows(delegate(IntPtr h, IntPtr l)
        {
            if (!IsWindowVisible(h)) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains((int)pid)) return true;
            if (WinClass(h) != "#32770") return true;
            var title = WinText(h);
            var sb = new StringBuilder();
            EnumChildWindows(h, delegate(IntPtr c, IntPtr l2)
            {
                var t = WinText(c);
                if (t.Length > 0) { sb.AppendLine(t); }
                return true;
            }, IntPtr.Zero);
            var all = title + " " + sb.ToString();
            if (all.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                PostMessage(h, 0x0111, (IntPtr)1, IntPtr.Zero); // WM_COMMAND = IDOK
            }
            return true;
        }, IntPtr.Zero);
    }
}
"@

$seen = New-Object 'System.Collections.Generic.HashSet[string]'
Write-Host "Watching Excel for compile-error dialogs. Log: $logPath"
Write-Host "Run the macro in Excel now. Stop with Ctrl+C."
$deadline = (Get-Date).AddHours(2)
try {
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 600
        $hits = [ExcelDialogWatcher]::FindNewDialogs($seen)
        foreach ($hit in $hits) {
            $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -LiteralPath $logPath -Value "===== $stamp =====","$hit",""
            Write-Host "Captured:`n$hit`n"
        }
        if ($hits.Count -gt 0) {
            [ExcelDialogWatcher]::DismissErrorDialogs()
        }
    }
    Write-Host "Watch window ended (2h limit)."
}
finally {
    Write-Host "Watcher stopped."
}