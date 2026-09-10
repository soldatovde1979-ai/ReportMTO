# run_load_and_report.ps1
# Version 1.0 / 2026-09-10
# Loads every data\*.json into the ALREADY OPEN production workbook (upsert via Power Query)
# and then builds pivots + generates the report. Attaches to the running Excel instance
# via GetActiveObject instead of opening a second copy of the book.
#
# Usage from the project root:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\run_load_and_report.ps1
#   powershell ... -File tools\run_load_and_report.ps1 -SkipLoad
#   powershell ... -File tools\run_load_and_report.ps1 -SkipReport

param(
    [switch]$SkipLoad,
    [switch]$SkipReport
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
$dataDir   = Join-Path $root "data"
$resultDir = Join-Path $root "result"

function Say([string]$s) { Write-Output ((Get-Date -Format "HH:mm:ss") + "  " + $s) }
function Die([string]$s) { Say ("DONE_FAIL " + $s); exit 1 }

Say "RUN_START (attach to running Excel)"

# Attach to the already running Excel instance.
$excel = $null
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Die "no running Excel instance found - open ReportMTO.xlsm first"
}

$wb = $null
try {
    # Find the workbook by name among the open workbooks.
    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq "ReportMTO.xlsm") { $wb = $w; break }
    }
    if ($null -eq $wb) { Die "ReportMTO.xlsm is not open in the running Excel" }

    Say ("ATTACHED book=" + $wb.FullName)

    $wsData = $wb.Sheets.Item("tbDATA")
    $lo = $wsData.ListObjects.Item("tbDATA")
    Say ("ROWS_START " + [int]$lo.ListRows.Count)

    # ---------------------------------------------------------------- LOAD
    if (-not $SkipLoad) {
        $files = @(Get-ChildItem -Path (Join-Path $dataDir "*.json") | Sort-Object Name)
        Say ("LOAD " + $files.Count + " json file(s)")
        if ($files.Count -eq 0) { Say "WARN no json files in data\" }
        foreach ($f in $files) {
            $t0 = Get-Date
            $rowsBefore = [int]$lo.ListRows.Count
            $safe = $f.FullName -replace '"', '""'
            $wb.Queries.Item("prmSourcePath").Formula =
                '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
            $qt = $lo.QueryTable
            $qt.BackgroundQuery = $false
            $null = $qt.Refresh($false)
            $rowsAfter = [int]$lo.ListRows.Count
            $sec = [int]((Get-Date) - $t0).TotalSeconds
            Say ("  LOADED " + $f.Name + " | rows " + $rowsBefore + " -> " + $rowsAfter + " | " + $sec + " s")
        }
        Say ("ROWS_AFTER_LOAD " + [int]$lo.ListRows.Count)
        $wb.Save()
    } else {
        Say "LOAD SKIP (-SkipLoad)"
    }

    # ---------------------------------------------------------------- REPORT
    if (-not $SkipReport) {
        Say "BuildPivots"
        $null = $excel.Run("modContentMTO.BuildPivots")
        Say "  BuildPivots ok"

        Say "GenerateReport"
        $startTime = Get-Date
        $null = $excel.Run("modMain.GenerateReport")

        $latest = Get-ChildItem $resultDir -Filter "Report_*.html" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-5) } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($null -eq $latest) {
            Say "REPORT FAIL - no fresh Report_*.html in result\"
        } else {
            Say ("RESULT_FILE " + $latest.FullName + " | " + [int]($latest.Length / 1024) + " KB")
        }
        $wb.Save()
    } else {
        Say "REPORT SKIP (-SkipReport)"
    }

    Say "DONE_OK"
} catch {
    Say ("EXCEPTION " + $_.Exception.Message)
    exit 1
} finally {
    # Do NOT close the workbook or quit Excel - it was already open by the user.
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
