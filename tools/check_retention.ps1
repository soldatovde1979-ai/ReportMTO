# check_retention.ps1
# Runs the ImportJSON pipeline on a COPY of the build book with the test JSON
# and prints tbDATA row counts under DATA/KEEP_WEEKS = 52 and = 1, each run
# twice (retention must be idempotent: reload gives the same row count).
# The original book is never opened; a watchdog kills ONLY the EXCEL process
# started by this script.
# ASCII-only by project rule (PowerShell 5.1 ANSI parsing).
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\check_retention.ps1
#   powershell ... -Book "C:\path\ReportMTO.xlsm" -Json "C:\path\sppr_tablet_1.json"

param(
    [string]$Book = (Join-Path "build" "ReportMTO v7.0.xlsm"),
    [string]$Json = (Join-Path "tests" "test_sppr_tablet_v1.json")
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
if (-not [System.IO.Path]::IsPathRooted($Book)) { $Book = [System.IO.Path]::GetFullPath((Join-Path $root $Book)) }
if (-not [System.IO.Path]::IsPathRooted($Json)) { $Json = [System.IO.Path]::GetFullPath((Join-Path $root $Json)) }

$tmp = Join-Path $env:TEMP "ReportMTO_retention_check.xlsm"
Copy-Item $Book $tmp -Force

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$job = Start-Job -ScriptBlock {
    param($sec, $ids)
    Start-Sleep -Seconds $sec
    foreach ($p in Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ($ids -notcontains $p.Id) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
} -ArgumentList 180, $before

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($tmp, 0, $false)

    function Set-KeepWeeks([object]$wbIn, [int]$val) {
        $ws = $wbIn.Sheets.Item("Variable")
        $tbl = $ws.ListObjects.Item("tblVariable")
        $keys = $tbl.ListColumns.Item("Key").DataBodyRange
        for ($r = 1; $r -le $keys.Rows.Count; $r++) {
            if ($keys.Cells.Item($r, 1).Text -eq "DATA/KEEP_WEEKS") {
                $null = $tbl.ListColumns.Item("Value").DataBodyRange.Cells.Item($r, 1).Value2 = $val
                return
            }
        }
        throw "DATA/KEEP_WEEKS not found in tblVariable"
    }

    $safe = $Json -replace '"', '""'
    $wb.Queries.Item("prmSourcePath").Formula = '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'

    $wsData = $wb.Sheets.Item("tbDATA")
    $lo = $wsData.ListObjects.Item("tbDATA")
    if ($lo.ListRows.Count -gt 0) { $lo.DataBodyRange.Delete() }
    $qt = $lo.QueryTable
    $qt.BackgroundQuery = $false

    Set-KeepWeeks $wb 0
    $null = $qt.Refresh($false)
    $n0 = [int]$lo.ListRows.Count
    $null = $qt.Refresh($false)
    $n0b = [int]$lo.ListRows.Count
    Write-Output ("KW0 first=" + $n0 + " reload=" + $n0b)

    Set-KeepWeeks $wb 52
    $null = $qt.Refresh($false)
    $n1 = [int]$lo.ListRows.Count
    $null = $qt.Refresh($false)
    $n1b = [int]$lo.ListRows.Count
    Write-Output ("KW52 first=" + $n1 + " reload=" + $n1b)

    Set-KeepWeeks $wb 1
    $null = $qt.Refresh($false)
    $n2 = [int]$lo.ListRows.Count
    $null = $qt.Refresh($false)
    $n2b = [int]$lo.ListRows.Count
    Write-Output ("KW1 first=" + $n2 + " reload=" + $n2b)

    $wb.Close($false)
} catch {
    Write-Output ("RETENTION_CHECK_FAIL - " + $_.Exception.Message)
} finally {
    try { $excel.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Stop-Job $job -ErrorAction SilentlyContinue
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}
