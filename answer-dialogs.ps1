# answer-dialogs.ps1
# Answers the report macro's modal dialogs so a report run can complete unattended.
#  - File picker "Select the CURRENT ..."  -> types the source path and opens it.
#  - File picker "Select the PRIOR ..."    -> Cancel (Esc) unless -PriorPath given.
#  - Any other dialog (MsgBox, error box)  -> Enter (dismiss; text is logged by the watcher).
# Log: log\answer-dialogs.log (per-line timestamps).
param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [string]$PriorPath = "",
    [string]$LogPath = "$PSScriptRoot\log\answer-dialogs.log"
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class Dlg {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    delegate bool EnumProc(IntPtr h, IntPtr l);
    static string T(IntPtr h) { var sb = new StringBuilder(1024); GetWindowText(h, sb, 1024); return sb.ToString(); }
    static string C(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
    public static List<IntPtr> ExcelDialogs() {
        var list = new List<IntPtr>();
        var pids = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName("EXCEL")) pids.Add(p.Id);
        if (pids.Count == 0) return list;
        EnumWindows(delegate(IntPtr h, IntPtr l) {
            if (!IsWindowVisible(h)) return true;
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains((int)pid)) return true;
            if (C(h) != "#32770") return true;
            list.Add(h);
            return true;
        }, IntPtr.Zero);
        return list;
    }
    public static string Text(IntPtr h) { return T(h); }
    public static void Focus(IntPtr h) { SetForegroundWindow(h); }
}
"@
function Log { param($m) $line = "$(Get-Date -Format 'HH:mm:ss') | $m"; Write-Output $line; Add-Content -LiteralPath $LogPath -Value $line }

Log "answerer started (pid $PID)"
$answered = @{}
$deadline = (Get-Date).AddMinutes(10)
try {
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 400
        $live = @{}
        foreach ($h in [Dlg]::ExcelDialogs()) { $live["$h"] = $h }
        foreach ($k in @($answered.Keys)) { if (-not $live.ContainsKey($k)) { [void]$answered.Remove($k) } }
        foreach ($k in $live.Keys) {
            if ($answered.ContainsKey($k)) { continue }
            $title = [Dlg]::Text($live[$k])
            if ($title -match 'Select the CURRENT') {
                Start-Sleep -Milliseconds 500
                [Dlg]::Focus($live[$k]) | Out-Null
                Start-Sleep -Milliseconds 200
                [System.Windows.Forms.SendKeys]::SendWait("^a")
                [System.Windows.Forms.SendKeys]::SendWait($SourcePath)
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                $answered[$k] = 1
                Log "answered CURRENT picker -> $SourcePath"
            } elseif ($title -match 'Select the PRIOR') {
                if ($PriorPath -ne "") {
                    Start-Sleep -Milliseconds 400
                    [Dlg]::Focus($live[$k]) | Out-Null
                    [System.Windows.Forms.SendKeys]::SendWait("^a")
                    [System.Windows.Forms.SendKeys]::SendWait($PriorPath)
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    Log "answered PRIOR picker -> $PriorPath"
                } else {
                    Start-Sleep -Milliseconds 400
                    [Dlg]::Focus($live[$k]) | Out-Null
                    [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                    Log "cancelled PRIOR picker (no prior report available)"
                }
                $answered[$k] = 1
            } elseif ($title -match 'Choose where') {
                Start-Sleep -Milliseconds 400
                [Dlg]::Focus($live[$k]) | Out-Null
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                $answered[$k] = 1
                Log "accepted default output folder"
            } else {
                Start-Sleep -Milliseconds 300
                [Dlg]::Focus($live[$k]) | Out-Null
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                $answered[$k] = 1
                Log "dismissed dialog: $title"
            }
        }
    }
} finally { Log "answerer stopped" }