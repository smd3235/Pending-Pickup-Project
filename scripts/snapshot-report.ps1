# snapshot-report.ps1
# Exports every sheet (Dashboard, manager tabs, Raw Data) of a generated
# Pending Pickup report as a PNG into documentation\scripts\screenshots\.
# Uses CopyPicture -> temp ChartObject -> Chart.Export so whole sheets are
# captured as images for the README.
param(
    [string]$Report = (Join-Path (Split-Path -Parent $PSScriptRoot) "Reports\Pending Pickup 2026-11-14.xlsx"),
    [string]$OutDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "documentation\screenshots")
)
$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $Report)) { throw "Report not found: $Report" }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

Add-Type -AssemblyName System.Windows.Forms
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
try {
    $wb = $excel.Workbooks.Open($Report)
    $outFiles = @()
    foreach ($ws in $wb.Worksheets) {
        $name = $ws.Name
        $safe = $name -replace '[\\/:*?"<>|]', '_'
        $png = Join-Path $OutDir "$safe.png"
        try {
            $ws.Activate() | Out-Null
            $used = $ws.UsedRange
            if ($used.Cells.Count -eq 1 -and [string]::IsNullOrWhiteSpace($used.Value2)) {
                Write-Host "SKIP empty sheet: $name"
                continue
            }
            $copied = $false
            for ($attempt = 1; $attempt -le 3 -and -not $copied; $attempt++) {
                try {
                    $null = $used.CopyPicture(1, 1)
                    $copied = $true
                }
                catch {
                    Start-Sleep -Milliseconds 200
                }
            }
            if (-not $copied) { throw "CopyPicture failed after 3 attempts" }
            $rows = [int]$used.Rows.Count
            $cols = [int]$used.Columns.Count
            $w = [Math]::Max(800, $cols * 70)
            $h = [Math]::Max(600, $rows * 22)
            # xlColumnClustered = -4109
            $shp = $ws.Shapes.AddChart(-4109, 0, 0, $w, $h)
            $cht = $shp.Chart
            try {
                $cht.Paste()
                $cht.Export($png)
                $outFiles += $png
                Write-Host "EXPORTED $name -> $png"
            }
            finally {
                $shp.Delete()
            }
        }
        catch {
            Write-Host "SKIP (failed) $name : $($_.Exception.Message)"
        }
    }
    $wb.Close($false)
    Write-Host "=== Exported $($outFiles.Count) sheet PNGs to $OutDir ==="
}
finally {
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
