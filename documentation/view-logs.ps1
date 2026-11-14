# Reads runtime logs and writes dated review copies with newest entries first.
param(
    [string]$LogDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "log")
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $LogDir)) { throw "Log directory not found: $LogDir" }

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$logFiles = @(Get-ChildItem -LiteralPath $LogDir -File | Where-Object { $_.Extension -in ".log", ".txt" } | Sort-Object Name)
foreach ($logFile in $logFiles) {
    $lines = @(Get-Content -LiteralPath $logFile.FullName -ErrorAction SilentlyContinue)
    [System.Array]::Reverse($lines)
    $header = @(
        "$stamp | Log file: $($logFile.Name)",
        "$stamp | Display order: newest first",
        ""
    )
    ($header + $lines) | Set-Content -LiteralPath $logFile.FullName -Encoding UTF8
    Write-Host "Reordered newest first: $($logFile.FullName)"
}

Write-Host "Updated $($logFiles.Count) runtime log files."