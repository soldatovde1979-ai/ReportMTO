# add_keep_weeks_key.ps1
# Idempotent: adds the DATA/KEEP_WEEKS = 52 row to the tblVariable table
# (sheet Variable) of the given workbooks. Does nothing when the key exists.
# ASCII-only by project rule (PowerShell 5.1 ANSI parsing).
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\add_keep_weeks_key.ps1
#       -> default books: build\ReportMTO_starter.xlsx, "build\ReportMTO v7.0.xlsm"
#   powershell ... -Books "C:\path\ReportMTO.xlsm"

param(
    [string[]]$Books = @(
        (Join-Path "build" "ReportMTO_starter.xlsx"),
        (Join-Path "build" "ReportMTO v7.0.xlsm")
    )
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    foreach ($b in $Books) {
        $full = $b
        if (-not [System.IO.Path]::IsPathRooted($full)) {
            $full = [System.IO.Path]::GetFullPath((Join-Path $root $full))
        }
        try {
            $wb = $excel.Workbooks.Open($full, 0, $false)

            $ws = $null
            foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "Variable") { $ws = $sh } }
            if ($null -eq $ws) { throw "no sheet Variable" }

            $lo = $null
            foreach ($l in $ws.ListObjects) { if ($l.Name -eq "tblVariable") { $lo = $l } }
            if ($null -eq $lo) { throw "no table tblVariable" }

            $found = $false
            $keyCol = $lo.ListColumns.Item("Key").DataBodyRange
            for ($r = 1; $r -le $keyCol.Rows.Count; $r++) {
                if ($keyCol.Cells.Item($r, 1).Text -eq "DATA/KEEP_WEEKS") { $found = $true }
            }

            if ($found) {
                Write-Output "EXISTS $full - DATA/KEEP_WEEKS already present"
            } else {
                $newRow = $lo.ListRows.Add() | Out-Null
                $n = $lo.ListRows.Count
                $lo.ListColumns.Item("Key").DataBodyRange.Cells.Item($n, 1).Value2 = "DATA/KEEP_WEEKS"
                $lo.ListColumns.Item("Value").DataBodyRange.Cells.Item($n, 1).Value2 = 52
                Write-Output "ADDED $full - DATA/KEEP_WEEKS = 52, tblVariable now " + $lo.Range.Rows.Count + " rows"
            }

            $wb.Save()
            $wb.Close($false)
        } catch {
            Write-Output "ERROR $full - $($_.Exception.Message)"
            try { $wb.Close($false) } catch { }
        }
    }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
