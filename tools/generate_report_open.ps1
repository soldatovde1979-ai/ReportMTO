# generate_report_open.ps1
# Version 1.0 / 2026-09-10
# Builds pivots and generates the report in the ALREADY OPEN workbook (attaches via
# GetActiveObject). MsgBox calls were removed from VBA, so GenerateReport returns
# without blocking on a modal dialog.
#
# Usage from the project root:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\generate_report_open.ps1

param(
    [string]$BookName = "ReportMTO.xlsm"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
$resultDir = Join-Path $root "result"

function Say([string]$s) { Write-Output ((Get-Date -Format "HH:mm:ss") + "  " + $s) }
function Die([string]$s) { Say ("DONE_FAIL " + $s); exit 1 }

Say "RUN_START (attach to running Excel)"

$excel = $null
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Die "no running Excel instance found"
}

$wb = $null
try {
    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq $BookName) { $wb = $w; break }
    }
    if ($null -eq $wb) { Die ("workbook not open: " + $BookName) }

    Say ("ATTACHED book=" + $wb.FullName)
    $lo = $wb.Sheets.Item("tbDATA").ListObjects.Item("tbDATA")
    Say ("ROWS " + [int]$lo.ListRows.Count)

    Say "BuildPivots"
    $null = $excel.Run("modContentMTO.BuildPivots")
    Say "  BuildPivots ok"

    Say "GenerateReport"
    $startTime = Get-Date
    $null = $excel.Run("modMain.GenerateReport")
    Say "  GenerateReport returned"

    $latest = Get-ChildItem $resultDir -Filter "Report_*.html" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $latest) {
        Say "REPORT FAIL - no fresh Report_*.html in result\"
    } else {
        Say ("RESULT_FILE " + $latest.FullName + " | " + [int]($latest.Length / 1024) + " KB")
    }

    $wb.Save()
    Say "SAVED"
    Say "DONE_OK"
} catch {
    Say ("EXCEPTION " + $_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
