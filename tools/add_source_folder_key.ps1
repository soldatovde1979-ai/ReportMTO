# add_source_folder_key.ps1
#
# Idempotent helper: adds key DATA/SOURCE_FOLDER=<Value> to the tblVariable table
# (sheet "Variable") of a ReportMTO workbook, unless the key already exists
# (existing value is kept). Used by tools\load_package.ps1.
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less .ps1 as ANSI and cyrillic
# string literals would break parsing.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools\add_source_folder_key.ps1 `
#            -BookPath "ReportMTO.xlsm" -Value "D:\GOOGLEDISK\PROJECTs\ReportMTO\data"

param(
    [Parameter(Mandatory = $true)][string]$BookPath,
    [Parameter(Mandatory = $true)][string]$Value
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

    $lo = $null
    foreach ($l in $ws.ListObjects) { if ($l.Name -eq "tblVariable") { $lo = $l; break } }
    if ($null -eq $lo) { throw "Table 'tblVariable' not found on sheet Variable in $BookPath" }

    $found = $false
    foreach ($r in $lo.ListRows) {
        if ($r.Range.Cells.Item(1, 1).Text -eq "DATA/SOURCE_FOLDER") { $found = $true; break }
    }

    if ($found) {
        Write-Output "DATA/SOURCE_FOLDER key already present - nothing to do ($BookPath)"
    } else {
        $newRow = $lo.ListRows.Add()
        $newRow.Range.Cells.Item(1, 1).Value2 = "DATA/SOURCE_FOLDER"
        $newRow.Range.Cells.Item(1, 2).Value2 = $Value
        Write-Output "DATA/SOURCE_FOLDER=$Value added ($BookPath)"
    }

    $wb.Save()
    Write-Output "Saved."
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    try { $excel.Quit() } catch { }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
