# compile_check.ps1
# Forces a full VBA compile of a workbook: copies it to %TEMP%, opens it via COM
# and runs the harmless public function HasColumn (modAggregate). VBA compiles
# the whole project before the first run, so a Compile error surfaces here as
# COMPILE_FAIL without touching the original book.
# A watchdog job kills ONLY the EXCEL processes started by this script (any
# Excel that existed before is never touched).
# ASCII-only by project rule (PowerShell 5.1 ANSI parsing).
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\compile_check.ps1
#   powershell ... -Book "C:\path\ReportMTO.xlsm"

param(
    [string]$Book = (Join-Path "build" "ReportMTO v7.0.xlsm")
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
if (-not [System.IO.Path]::IsPathRooted($Book)) {
    $Book = [System.IO.Path]::GetFullPath((Join-Path $root $Book))
}

$tmp = Join-Path $env:TEMP "ReportMTO_compile_check.xlsm"
Copy-Item $Book $tmp -Force

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$job = Start-Job -ScriptBlock {
    param($sec, $ids)
    Start-Sleep -Seconds $sec
    foreach ($p in Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ($ids -notcontains $p.Id) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
} -ArgumentList 120, $before

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($tmp, 0, $false)
    try {
        $res = $excel.Run("HasColumn", "x")
        Write-Output ("COMPILE_OK " + $Book + " - HasColumn returned " + $res)
    } catch {
        Write-Output ("COMPILE_FAIL " + $Book + " - " + $_.Exception.Message)
    }
    try { $wb.Close($false) } catch { }
} catch {
    Write-Output ("OPEN_FAIL " + $Book + " - " + $_.Exception.Message)
}
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Stop-Job $job -ErrorAction SilentlyContinue
Remove-Job $job -Force -ErrorAction SilentlyContinue
Remove-Item $tmp -Force -ErrorAction SilentlyContinue
