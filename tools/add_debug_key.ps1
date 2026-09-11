# add_debug_key.ps1
#
# Idempotent helper: adds key DEBUG=<Value> to the first table on sheet "Variable"
# of a ReportMTO workbook, unless the key already exists (existing value is kept).
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less .ps1 as ANSI and cyrillic
# string literals would break parsing.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools\add_debug_key.ps1 `
#            -BookPath "build\ReportMTO v7.0.xlsm" -Value 0

param(
    [Parameter(Mandatory = $true)][string]$BookPath,
    [string]$Value = "0"
)

$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    $wb = $excel.Workbooks.Open((Resolve-Path $BookPath).Path)

    $ws = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq "Variable") { $ws = $s; break }
    }
    if ($null -eq $ws) { throw "Sheet 'Variable' not found in $BookPath" }
    if ($ws.ListObjects.Count -lt 1) { throw "No table (ListObject) on sheet 'Variable' in $BookPath" }

    $lo = $ws.ListObjects.Item(1)
    $found = $false
    foreach ($r in $lo.ListRows) {
        if ($r.Range.Cells.Item(1, 1).Text -eq "DEBUG") { $found = $true; break }
    }

    if ($found) {
        Write-Output "DEBUG key already present - nothing to do ($BookPath)"
    } else {
        $newRow = $lo.ListRows.Add()
        $newRow.Range.Cells.Item(1, 1).Value2 = "DEBUG"
        $newRow.Range.Cells.Item(1, 2).Value2 = $Value
        Write-Output "DEBUG=$Value added ($BookPath)"
    }

    $wb.Save()
    Write-Output "Saved."
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    try { $excel.Quit() } catch { }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
