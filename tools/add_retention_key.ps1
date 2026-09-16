# add_retention_key.ps1
# Idempotent: adds the REPORT/RETENTION_WEEKS row (EMPTY value, reserved) to the
# tblVariable table (sheet Variable) of the given workbooks.
# Per P0-3 (TZ v1.2): the key is reserved, NOT read by code - history is stored
# entirely, no deletion. Does nothing when the key exists.
# ASCII-only by project rule (PowerShell 5.1 ANSI parsing).
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\add_retention_key.ps1
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
                if ($keyCol.Cells.Item($r, 1).Text -eq "REPORT/RETENTION_WEEKS") { $found = $true }
            }

            if ($found) {
                Write-Output "EXISTS $full - REPORT/RETENTION_WEEKS already present"
            } else {
                $newRow = $lo.ListRows.Add() | Out-Null
                $n = $lo.ListRows.Count
                $lo.ListColumns.Item("Key").DataBodyRange.Cells.Item($n, 1).Value2 = "REPORT/RETENTION_WEEKS"
                # Empty value on purpose: key is reserved, history is kept entirely (P0-3).
                $lo.ListColumns.Item("Value").DataBodyRange.Cells.Item($n, 1).Value2 = ""
                Write-Output "ADDED $full - REPORT/RETENTION_WEEKS (empty, reserved), tblVariable now " + $lo.Range.Rows.Count + " rows"
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
