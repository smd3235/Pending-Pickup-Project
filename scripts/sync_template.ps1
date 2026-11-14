# sync_template.ps1
# Re-imports Pending_Pickup_Macro.bas into the template's embedded VBA project
# and ensures the styled CREATE PENDING PICKUP REPORT button is installed on START HERE.
# Run this after every change to the .bas file.
param(
    [string]$Template = (Join-Path (Split-Path -Parent $PSScriptRoot) "Pending Pickup Template.xlsm"),
    [string]$Bas      = (Join-Path (Split-Path -Parent $PSScriptRoot) "Pending_Pickup_Macro.bas"),
    [string]$ModuleName = "PendingPickupReport",
    [string]$ButtonName = "btnCreatePendingPickupReport"
)
$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $Template)) { throw "Template not found: $Template" }
if (-not (Test-Path -LiteralPath $Bas))      { throw "Bas not found: $Bas" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
try {
    $wb = $excel.Workbooks.Open($Template)
    $vbp = $wb.VBProject

    # Replace the module so there is never a duplicate.
    $existing = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $ModuleName) { $existing = $c; break } }
    if ($existing -ne $null) { $vbp.VBComponents.Remove($existing) }
    $null = $vbp.VBComponents.Import($Bas)

    # Ensure the styled button shape exists and points at the macro.
    $ws = $wb.Worksheets("START HERE")
    $btn = $null
    foreach ($s in $ws.Shapes) { if ($s.Name -eq $ButtonName) { $btn = $s; break } }
    if ($btn -eq $null) {
        $cta = $ws.Range("B14")
        $btn = $ws.Shapes.AddShape(5, $cta.Left + 20, $cta.Top + 8, 360, 48)
        $btn.Name = $ButtonName
        $btn.Fill.ForeColor.RGB = 7949855
        $btn.Line.ForeColor.RGB = 4116214
        $btn.TextFrame2.TextRange.Text = "CREATE PENDING PICKUP REPORT"
        $btn.TextFrame2.TextRange.Font.Bold = $true
        $btn.TextFrame2.TextRange.Font.Size = 14
        $btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        $btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    }
    $btn.OnAction = "CreatePendingPickupReport"

    $wb.Save()
    $wb.Close($true)
    Write-Host "Synced '$ModuleName' and button '$ButtonName' into $Template"
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}