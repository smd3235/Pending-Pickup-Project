# snapshot-report.ps1
# Exports every sheet (Dashboard, manager tabs, Raw Data) of a generated
# Pending Pickup report as a PNG into documentation\scripts\screenshots\.
# Uses CopyPicture -> standalone Chart sheet -> Chart.Export so whole sheets are
# captured as images for the README. Defaults to the newest workbook in Reports\.
param(
    [string]   $Path,
    [string[]] $Sheets,
    [string]   $OutFolder,
    [string]   $Prefix = "report",
    [switch]   $List
)
$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $candidates = @(Get-ChildItem -LiteralPath (Join-Path $ProjectRoot "Reports") -Filter "*.xlsx" -File -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending)
    if ($candidates.Count -eq 0) {
        Write-Host "No workbook given and Reports\ holds no .xlsx. Pass -Path." -ForegroundColor Red
        exit 2
    }
    $Path = $candidates[0].FullName
    Write-Host "Using newest report: $Path" -ForegroundColor DarkGray
}

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host "Workbook not found: $Path" -ForegroundColor Red
    exit 2
}
$Path = (Resolve-Path -LiteralPath $Path).Path

if (-not $OutFolder) { $OutFolder = Join-Path $ProjectRoot "documentation\screenshots" }
if (-not (Test-Path -LiteralPath $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder | Out-Null }

# Always start fresh: delete every existing PNG so committed screenshots
# reflect only the latest report run, never a stale mix of old/new sheets.
Get-ChildItem -LiteralPath $OutFolder -Filter *.png -File | Remove-Item -Force -ErrorAction SilentlyContinue

function Save-SheetImage {
    param($Workbook, $Worksheet, [string]$PngPath)

    $used = $Worksheet.UsedRange
    $range = $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells($used.Row + $used.Rows.Count - 1, $used.Column + $used.Columns.Count - 1))

    # Activate sheet and set zoom to 100% so CopyPicture captures at full resolution,
    # not zoomed out to 85% (which makes everything look squished).
    $Worksheet.Activate() | Out-Null
    $Workbook.Application.ActiveWindow.Zoom = 100

    $range.CopyPicture(1, 2) | Out-Null   # xlScreen=1, xlBitmap=2

    $rows = [int]$used.Rows.Count
    $cols = [int]$used.Columns.Count
    # Wider chart so content isn't squished onto one page width.
    $w = [Math]::Max(1800, $cols * 110)
    $h = [Math]::Max(900,  $rows * 30)

    $chart = $Workbook.Charts.Add()
    try {
        $chart.Paste() | Out-Null
        Start-Sleep -Milliseconds 250
        if (Test-Path -LiteralPath $PngPath) { Remove-Item -LiteralPath $PngPath -Force }
        $chart.Export($PngPath)
    }
    finally {
        $chart.Delete()
    }
}

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3
$wb = $null
$exported = 0
try {
    $wb = $xl.Workbooks.Open($Path, $false, $true)

    if ($List -or -not $Sheets) {
        Write-Host "Sheets in $(Split-Path -Leaf $Path):"
        foreach ($s in $wb.Worksheets) {
            $u = $s.UsedRange
            Write-Host ("  {0,-28} used {1}" -f $s.Name, $u.Address(0, 0))
        }
    }

    $targets = @()
    if ($Sheets) {
        foreach ($name in $Sheets) {
            $found = $null
            foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { $found = $s; break } }
            if ($found) { $targets += $found }
            else { Write-Host "Sheet not found, skipped: $name" -ForegroundColor Yellow }
        }
    }
    else {
        $targets = @($wb.Worksheets)
    }

    foreach ($ws in $targets) {
        $safe = ($ws.Name -replace "[^\w\-]+", "-").Trim("-").ToLower()
        $png = Join-Path $OutFolder ("{0}-{1}.png" -f $Prefix, $safe)
        Save-SheetImage -Workbook $wb -Worksheet $ws -PngPath $png
        $kb = [math]::Round((Get-Item -LiteralPath $png).Length / 1KB, 1)
        Write-Host ("  exported {0}  ({1} KB)" -f $png, $kb) -ForegroundColor Green
        $exported++
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

if ($exported -eq 0 -and -not $List -and $Sheets) {
    Write-Host "Nothing exported." -ForegroundColor Red
    exit 3
}
exit 0
